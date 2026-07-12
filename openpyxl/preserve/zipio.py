# paper-xlsx: deterministic, atomic zip writing + raw compressed-stream copy

"""Zip-layer machinery for the preserve-mode save.

- Deterministic output: every entry gets a fixed timestamp and uniform
  attributes, so part payloads (and, on the raw-copy path, compressed
  streams) are reproducible run-to-run.
- Raw compressed-stream copy: untouched parts are copied without
  recompression (measured 235x faster and byte-identical); guarded by the
  the raw-copy guards, with transparent fallback to recompression.
- Atomic targets: path targets are written temp-file-then-``os.replace``
  (in-place truncation is the measured corruption hazard); in-memory targets
  must be exact ``io.BytesIO`` instances or verified path-backed
  ``io.BufferedRandom`` handles (the form pandas uses for append mode).
"""

import io
import hashlib
import os
import stat
import struct
import tempfile
import zipfile
from dataclasses import dataclass

# Fixed timestamp for deterministic archives (zip epoch).
FIXED_DATE_TIME = (1980, 1, 1, 0, 0, 0)
_EXTERNAL_ATTR = 0o600 << 16


def _probe_private_zipfile_api():
    """The raw-copy fast path uses CPython-private ZipFile internals
    (validated on 3.9-3.13). Probe once; absence disables the fast path."""
    try:
        buf = io.BytesIO()
        zf = zipfile.ZipFile(buf, "w")
        ok = all(hasattr(zf, attr) for attr in
                 ("fp", "filelist", "NameToInfo", "start_dir", "_writecheck"))
        ok = ok and hasattr(zipfile.ZipInfo("x"), "FileHeader")
        zf.close()
        return ok
    except Exception:
        return False


RAW_COPY_AVAILABLE = _probe_private_zipfile_api()

_ZIP64_LIMIT = 0xFFFFFFFF
_RAW_COPY_ALLOWED_FLAGS = 0x0006 | 0x0800


@dataclass(frozen=True)
class PathIdentity:
    """Content-bound identity of one pathname and its resolved occupant."""

    requested: str
    resolved: str
    entry: object
    occupant: object
    digest: object

    @property
    def exists(self):
        return self.entry is not None


def _stat_identity(value):
    if value is None:
        return None
    return (
        value.st_dev, value.st_ino, stat.S_IFMT(value.st_mode),
        value.st_size, value.st_mtime_ns, value.st_ctime_ns,
    )


def _hash_file(path):
    digest = hashlib.sha256()
    with open(path, "rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def path_identity(path, *, allow_missing=False):
    """Capture a stable, content-bound identity for ``path``."""
    requested = os.path.abspath(os.fspath(path))
    try:
        entry_before = os.lstat(requested)
    except FileNotFoundError:
        if not allow_missing:
            raise
        return PathIdentity(
            requested, requested, None, None, None)
    resolved = os.path.realpath(requested)
    occupant_before = os.stat(resolved)
    if not stat.S_ISREG(occupant_before.st_mode):
        from openpyxl.errors import UnsupportedStructureError

        raise UnsupportedStructureError(
            "workbook path {0!r} is not a regular file; refusing a path "
            "that could block or change type. Nothing was written."
            .format(requested),
            kind="path-not-regular-file", anchor=requested)
    digest = _hash_file(resolved)
    entry_after = os.lstat(requested)
    occupant_after = os.stat(resolved)
    if (_stat_identity(entry_before) != _stat_identity(entry_after)
            or _stat_identity(occupant_before) !=
            _stat_identity(occupant_after)
            or os.path.realpath(requested) != resolved):
        _identity_refusal(requested)
    return PathIdentity(
        requested, resolved, _stat_identity(entry_after),
        _stat_identity(occupant_after), digest)


def read_path_snapshot(path, *, context="preserve-mode workbook"):
    """Read one stable regular-file snapshot and its path identity."""
    from openpyxl.preserve.limits import read_bounded

    before = path_identity(path)
    with open(before.resolved, "rb") as source:
        payload = read_bounded(source, context=context)
    after = path_identity(path)
    if before != after:
        _identity_refusal(before.requested)
    return payload, after


def read_path_handle_snapshot(handle, *, context="preserve-mode workbook"):
    """Return a stable snapshot for a named regular-file handle, or None."""
    try:
        requested = os.path.abspath(os.fspath(handle.name))
        descriptor = os.fstat(handle.fileno())
    except (AttributeError, OSError, TypeError, ValueError):
        return None
    if not stat.S_ISREG(descriptor.st_mode) or not os.path.exists(requested):
        return None
    identity = path_identity(requested)
    if not os.path.samestat(descriptor, os.stat(identity.resolved)):
        _identity_refusal(requested)
    position = handle.tell()
    try:
        handle.seek(0)
        from openpyxl.preserve.limits import read_bounded

        payload = read_bounded(handle, context=context)
    finally:
        handle.seek(position)
    after = path_identity(requested)
    if identity != after or not os.path.samestat(
            os.fstat(handle.fileno()), os.stat(after.resolved)):
        _identity_refusal(requested)
    return payload, after


def _identity_refusal(path, cause=None):
    from openpyxl.errors import UnsupportedStructureError

    raise UnsupportedStructureError(
        "destination {0!r} changed after its identity was captured; "
        "refusing to overwrite another occupant. Nothing was written."
        .format(path),
        kind="destination-identity-changed", anchor=path) from cause


def _same_occupant(left, right):
    def stable(stat_key):
        if stat_key is None:
            return None
        # Atomic exchange changes ctime even though the same inode and bytes
        # were displaced. All precommit checks still compare full identities.
        return stat_key[:5]

    return (
        left.exists == right.exists
        and left.digest == right.digest
        and stable(left.occupant) == stable(right.occupant)
    )


def _assert_path_identity(expected):
    try:
        current = path_identity(
            expected.requested, allow_missing=not expected.exists)
    except (FileNotFoundError, OSError) as exc:
        _identity_refusal(expected.requested, exc)
    binding_matches = (
        current.requested == expected.requested
        and current.resolved == expected.resolved
        and _same_occupant(current, expected)
    )
    if not binding_matches:
        _identity_refusal(expected.requested)


def _posix_exchange(first, second):
    """Atomically exchange two path entries on supported POSIX systems."""
    import ctypes
    import sys

    libc = ctypes.CDLL(None, use_errno=True)
    first_b = os.fsencode(first)
    second_b = os.fsencode(second)
    if sys.platform == "darwin" and hasattr(libc, "renamex_np"):
        result = libc.renamex_np(first_b, second_b, 0x00000002)
    elif hasattr(libc, "renameat2"):
        result = libc.renameat2(-100, first_b, -100, second_b, 0x2)
    else:
        raise OSError("atomic path exchange is unavailable")
    if result:
        error = ctypes.get_errno()
        raise OSError(error, os.strerror(error))


def _windows_replace(destination, replacement, backup):
    import ctypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    replace = kernel32.ReplaceFileW
    replace.argtypes = (
        ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_wchar_p,
        ctypes.c_uint, ctypes.c_void_p, ctypes.c_void_p)
    replace.restype = ctypes.c_int
    if not replace(destination, replacement, backup, 0, None, None):
        raise ctypes.WinError(ctypes.get_last_error())


def _rollback_exchange(displaced, destination, candidate):
    """Restore ``displaced`` only when the exchange removes our candidate."""
    _posix_exchange(displaced, destination)
    removed = path_identity(displaced)
    if _same_occupant(removed, candidate):
        return True
    # A newer writer arrived before rollback. Put it back; never replace it
    # with the older occupant this operation originally displaced.
    _posix_exchange(displaced, destination)
    return False


def _conditional_replace(tmp_path, destination, expected, precommit=None,
                         postcommit=None):
    """Install ``tmp_path`` only if ``expected`` is still the occupant."""
    _assert_path_identity(expected)
    candidate = path_identity(tmp_path)
    if precommit is not None:
        precommit()
    if not expected.exists:
        try:
            os.link(tmp_path, destination)
        except FileExistsError as exc:
            _identity_refusal(expected.requested, exc)
        os.unlink(tmp_path)
        installed = path_identity(expected.requested)
        if not _same_occupant(installed, candidate):
            _identity_refusal(expected.requested)
        if postcommit is not None:
            try:
                postcommit()
            except BaseException:
                current = path_identity(expected.requested)
                if _same_occupant(current, candidate):
                    os.unlink(destination)
                raise
        return

    if os.name == "nt":
        backup = tmp_path + ".previous"
        _windows_replace(destination, tmp_path, backup)
        displaced = backup
    else:
        _posix_exchange(tmp_path, destination)
        displaced = tmp_path
    try:
        previous = path_identity(displaced)
        binding_ok = (
            os.path.realpath(expected.requested) == expected.resolved)
        if not _same_occupant(previous, expected) or not binding_ok:
            if os.name == "nt":
                rollback_displaced = backup + ".candidate"
                _windows_replace(destination, backup, rollback_displaced)
                try:
                    os.unlink(rollback_displaced)
                except OSError:
                    pass
            else:
                _rollback_exchange(displaced, destination, candidate)
            _identity_refusal(expected.requested)
        installed = path_identity(expected.requested)
        if not _same_occupant(installed, candidate):
            # Another writer replaced the just-committed candidate. Do not
            # overwrite that newer occupant while attempting rollback.
            _identity_refusal(expected.requested)
        if postcommit is not None:
            try:
                postcommit()
            except BaseException:
                current = path_identity(expected.requested)
                if _same_occupant(current, candidate):
                    if os.name == "nt":
                        rollback_candidate = displaced + ".candidate"
                        _windows_replace(
                            destination, displaced, rollback_candidate)
                        try:
                            os.unlink(rollback_candidate)
                        except OSError:
                            pass
                    else:
                        _rollback_exchange(
                            displaced, destination, candidate)
                raise
        try:
            os.unlink(displaced)
        except OSError as exc:
            import warnings

            warnings.warn(
                "the workbook was saved correctly, but the displaced "
                "destination could not be removed ({0})".format(exc),
                RuntimeWarning, stacklevel=4)
    except BaseException:
        raise


def raw_copy_supported(info):
    """D10 guards: data-descriptor entries (GP flag bit 3), zip64-sized
    entries, and exotic compression methods take the recompression fallback."""
    if not RAW_COPY_AVAILABLE:
        return False
    if info.flag_bits & 0x8:            # data descriptor: sizes live after payload
        return False
    if info.flag_bits & ~_RAW_COPY_ALLOWED_FLAGS:
        return False
    # entries not read from an archive may lack size/offset attributes
    compress_size = getattr(info, "compress_size", None)
    header_offset = getattr(info, "header_offset", None)
    if compress_size is None or header_offset is None:
        return False
    if compress_size >= _ZIP64_LIMIT or info.file_size >= _ZIP64_LIMIT:
        return False
    if header_offset >= _ZIP64_LIMIT:
        return False
    if info.compress_type not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
        return False
    return True


def _read_raw_stream(zin, info):
    """Read one entry's compressed byte stream straight from the archive,
    verifying central-vs-local header AGREEMENT first (a zip-confusion payload shows different content to parsers
    that trust different headers). Returns None on disagreement — the
    caller falls back to recompression, which normalizes the entry to the
    central directory's view (the view zipfile and Excel read)."""
    f = zin.fp
    f.seek(info.header_offset)
    header = f.read(30)
    if header[:4] != b"PK\x03\x04":
        raise zipfile.BadZipFile(
            "bad local file header for {0!r}".format(info.filename))
    flags, method = struct.unpack("<HH", header[6:10])
    local_crc, local_csize, local_usize = struct.unpack("<LLL",
                                                        header[14:26])
    name_len, extra_len = struct.unpack("<HH", header[26:30])
    local_name = f.read(name_len)
    encoding = "utf-8" if flags & 0x0800 else "cp437"
    try:
        expected_name = info.filename.encode(encoding)
    except UnicodeEncodeError:
        return None
    if local_name != expected_name or flags != info.flag_bits \
            or method != info.compress_type:
        return None
    # sizes/CRC of 0 in the local header are legal when a data
    # descriptor follows — but descriptor entries never reach this path
    # (raw_copy_supported excludes GP bit 3), so disagreement is real
    if (local_crc, local_csize, local_usize) != (
            info.CRC, info.compress_size, info.file_size):
        return None
    f.seek(info.header_offset + 30 + name_len + extra_len)
    payload = f.read(info.compress_size)
    if len(payload) != info.compress_size:
        raise zipfile.BadZipFile(
            "truncated compressed stream for {0!r}".format(info.filename))
    return payload


def copy_entry(zin, info, zout):
    """Copy one entry from ``zin`` into ``zout``, raw when possible.

    Payload bytes are identical either way; the raw path also preserves the
    compressed stream. Entry metadata is normalized for determinism.
    """
    payload_stream = _read_raw_stream(zin, info) \
        if raw_copy_supported(info) else None
    if payload_stream is not None:
        new = zipfile.ZipInfo(info.filename, date_time=FIXED_DATE_TIME)
        new.compress_type = info.compress_type
        new.flag_bits = info.flag_bits
        new.external_attr = _EXTERNAL_ATTR
        new.file_size = info.file_size
        new.CRC = info.CRC
        new.compress_size = info.compress_size
        zout._writecheck(new)
        zout._didModify = True
        new.header_offset = zout.fp.tell()
        zout.fp.write(new.FileHeader())
        zout.fp.write(payload_stream)
        zout.start_dir = zout.fp.tell()
        zout.filelist.append(new)
        zout.NameToInfo[new.filename] = new
    else:
        write_entry(zout, info.filename, zin.read(info.filename),
                    compress_type=info.compress_type
                    if info.compress_type in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED)
                    else zipfile.ZIP_DEFLATED)


def write_entry(zout, name, payload, compress_type=zipfile.ZIP_DEFLATED):
    """Write one entry with deterministic metadata."""
    info = zipfile.ZipInfo(name, date_time=FIXED_DATE_TIME)
    info.compress_type = compress_type
    info.external_attr = _EXTERNAL_ATTR
    zout.writestr(info, payload)


def build_archive_bytes(build):
    """Run ``build(zout)`` against an in-memory archive; return its bytes."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as zout:
        build(zout)
    return buf.getvalue()


def validate_target(target):
    """Refuse unsupported file-like destinations before archive planning."""
    if not hasattr(target, "write"):
        os.fspath(target)
        return
    message = (
        "preserve-mode file-like save destinations must be open, exact "
        "io.BytesIO instances with no exported buffer views or verified "
        "path-backed io.BufferedRandom handles; use a filesystem path to "
        "write to another stream type"
    )
    if type(target) is io.BufferedRandom:
        if _path_backed_target(target) is None:
            raise TypeError(message)
        return
    if type(target) is not io.BytesIO:
        raise TypeError(message)
    try:
        # Even a no-op resize refuses while a buffer view is exported. Probe
        # that state without changing the bytes or cursor.
        payload, _position, _attributes = io.BytesIO.__getstate__(target)
        io.BytesIO.truncate(target, len(payload))
    except (BufferError, ValueError) as exc:
        raise TypeError(message) from exc


def _path_backed_target(target):
    """Return the live path owned by an exact ``BufferedRandom`` handle."""
    if type(target) is not io.BufferedRandom or target.closed:
        return None
    writable = getattr(target, "writable", None)
    if not callable(writable) or not writable():
        return None
    try:
        requested = os.path.abspath(os.fspath(target.name))
        if not os.path.samestat(os.fstat(target.fileno()),
                                os.stat(requested)):
            return None
        target.tell()
    except (OSError, TypeError, ValueError):
        return None
    return requested


def _reopen_buffered_random(target, path, position):
    """Rebind a closed exact ``BufferedRandom`` to ``path`` in place."""
    raw = io.FileIO(path, "r+b")
    try:
        io.BufferedRandom.__init__(target, raw)
    except BaseException:
        raw.close()
        raise
    target.seek(position)


def _replace_for_open_path_windows(tmp_path, destination, handle,
                                   reopen_path, expected=None,
                                   precommit=None, postcommit=None):
    """Windows cannot replace a path while its ordinary handle is open."""
    if expected is None:
        expected = path_identity(destination)
    original_position = handle.tell()
    handle.close()
    try:
        _conditional_replace(
            tmp_path, destination, expected, precommit=precommit,
            postcommit=postcommit)
    except BaseException:
        _reopen_buffered_random(handle, reopen_path, original_position)
        raise
    try:
        _reopen_buffered_random(
            handle, reopen_path, os.path.getsize(destination))
    except BaseException as exc:
        # The file commit is already complete. Raising here would tell the
        # caller the save failed even though disk changed, so surface the one
        # legal post-commit outcome: an explicit warning that the handle is
        # closed and must be reopened by the caller.
        import warnings

        from openpyxl.errors import HandleRebindWarning

        warnings.warn(HandleRebindWarning(
            "the workbook was saved correctly, but the destination handle "
            "could not be reopened after atomic replacement ({0}); the "
            "handle is closed".format(exc)), stacklevel=4)


def _replace_for_open_path(tmp_path, destination, handle, reopen_path,
                           expected, precommit=None, postcommit=None):
    """Replace a path and keep pandas' exact BufferedRandom usable."""
    if handle is None:
        _conditional_replace(
            tmp_path, destination, expected, precommit=precommit,
            postcommit=postcommit)
        return
    if os.name == "nt":
        _replace_for_open_path_windows(
            tmp_path, destination, handle, reopen_path, expected,
            precommit=precommit, postcommit=postcommit)
        return
    original_position = handle.tell()
    handle.flush()
    original_fd = os.dup(handle.raw.fileno())
    try:
        replacement = io.FileIO(tmp_path, "r+b")
    except BaseException:
        os.close(original_fd)
        raise
    rebound = False
    try:
        # Switch the existing FileIO descriptor before the rename. The
        # BufferedRandom object stays open and keeps its original .name, while
        # any failure can still restore the old descriptor before disk changes.
        os.dup2(replacement.fileno(), handle.raw.fileno())
        rebound = True
        handle.seek(0, os.SEEK_END)
        _conditional_replace(
            tmp_path, destination, expected, precommit=precommit,
            postcommit=postcommit)
    except BaseException:
        if rebound:
            os.dup2(original_fd, handle.raw.fileno())
            handle.seek(0, os.SEEK_END)
            handle.seek(original_position)
        raise
    finally:
        try:
            replacement.close()
        except OSError:
            pass
        try:
            os.close(original_fd)
        except OSError:
            pass


def _path_destination(target, expected_identity=None):
    """Return the atomic replacement path and any existing mode bits."""
    requested = os.path.abspath(os.fspath(target))
    destination = os.path.realpath(requested) if os.path.islink(requested) \
        else requested
    if expected_identity is not None:
        if requested != expected_identity.requested:
            raise ValueError("expected identity belongs to a different path")
        identity = expected_identity
    else:
        identity = path_identity(requested, allow_missing=True)
    try:
        destination_mode = os.stat(destination).st_mode
        mode = destination_mode & 0o7777
        if mode & 0o222 == 0:
            from openpyxl.errors import UnsupportedStructureError

            raise UnsupportedStructureError(
                "destination {0!r} is not writable; refusing to replace it. "
                "Nothing was written.".format(requested),
                kind="destination-not-writable", anchor=requested)
    except FileNotFoundError:
        mode = None
    return destination, mode, identity


def _apply_creation_mode(path):
    """Give a new destination the mode a normal ``open(..., 'wb')`` gets."""
    probe = path + ".mode"
    fd = None
    try:
        fd = os.open(probe, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o666)
        mode = stat.S_IMODE(os.fstat(fd).st_mode)
    finally:
        if fd is not None:
            os.close(fd)
        try:
            os.unlink(probe)
        except OSError:
            pass
    os.chmod(path, mode)


def deliver(data, target, *, expected_identity=None, validator=None,
            precommit=None, postcommit=None):
    """Deliver bytes to a path or supported transactional in-memory handle.

    Path targets: temp file in the same directory + ``os.replace`` — the
    original survives any mid-write crash (never in-place truncation).
    An exact ``io.BytesIO`` receives one built-in state replacement. A verified
    path-backed ``io.BufferedRandom`` is rebound around the same temp-file
    replacement used for paths. Arbitrary streams are refused because their
    write and rollback behavior cannot be proven atomic.
    """
    path_handle = None
    reopen_path = None
    if hasattr(target, "write"):
        validate_target(target)
        if type(target) is io.BytesIO:
            if validator is not None:
                validator(bytes(data))
            if precommit is not None:
                precommit()
            _payload, _position, attributes = io.BytesIO.__getstate__(target)
            io.BytesIO.__setstate__(target, (data, len(data), attributes))
            return None
        path_handle = target
        reopen_path = _path_backed_target(target)
        target = reopen_path

    requested = os.path.abspath(os.fspath(target))
    target, mode, identity = _path_destination(
        target, expected_identity=expected_identity)
    directory = os.path.dirname(os.path.abspath(target))
    fd, tmp_path = tempfile.mkstemp(prefix=".paper_save_", suffix=".tmp",
                                    dir=directory)
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())     # durability BEFORE the rename
        if validator is not None:
            validator(tmp_path)
        if mode is not None:
            os.chmod(tmp_path, mode)
        else:
            _apply_creation_mode(tmp_path)
        _replace_for_open_path(
            tmp_path, target, path_handle, reopen_path, identity,
            precommit=precommit, postcommit=postcommit)
        _fsync_directory(directory)  # ... and of the rename itself
        return path_identity(requested)
    except BaseException:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        raise


def _fsync_directory(directory):
    """Best-effort fsync of a directory (the rename's durability); not
    every platform allows opening directories."""
    try:
        dfd = os.open(directory, os.O_RDONLY)
    except OSError:
        return
    try:
        os.fsync(dfd)
    except OSError:
        pass
    finally:
        os.close(dfd)


def build_and_deliver(build_fn, target, *, expected_identity=None,
                      validator=None, precommit=None, postcommit=None):
    """Build the archive DIRECTLY into the delivery temp file for path
    targets (~1x file-size peak memory instead of
    a whole in-memory copy), atomically replaced and fsynced; exact
    ``io.BytesIO`` targets keep the in-memory build."""
    path_handle = None
    reopen_path = None
    if hasattr(target, "write"):
        validate_target(target)
        if type(target) is io.BytesIO:
            return deliver(build_archive_bytes(build_fn), target,
                           validator=validator, precommit=precommit,
                           postcommit=postcommit)
        path_handle = target
        reopen_path = _path_backed_target(target)
        target = reopen_path
    requested = os.path.abspath(os.fspath(target))
    target, mode, identity = _path_destination(
        target, expected_identity=expected_identity)
    directory = os.path.dirname(os.path.abspath(target))
    fd, tmp_path = tempfile.mkstemp(prefix=".paper_save_", suffix=".tmp",
                                    dir=directory)
    try:
        with os.fdopen(fd, "wb") as f:
            with zipfile.ZipFile(f, "w", zipfile.ZIP_DEFLATED) as zout:
                build_fn(zout)
            f.flush()
            os.fsync(f.fileno())
        if validator is not None:
            validator(tmp_path)
        if mode is not None:
            os.chmod(tmp_path, mode)
        else:
            _apply_creation_mode(tmp_path)
        _replace_for_open_path(
            tmp_path, target, path_handle, reopen_path, identity,
            precommit=precommit, postcommit=postcommit)
        _fsync_directory(directory)
        return path_identity(requested)
    except BaseException:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        raise
