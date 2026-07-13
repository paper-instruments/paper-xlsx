# paper-xlsx: bounded reads for agent-facing package inputs

import os

from openpyxl.errors import UnsupportedStructureError


MAX_SOURCE_BYTES = 512 * 1024 * 1024
_READ_CHUNK = 1024 * 1024


def _too_large(context, size, max_bytes):
    raise UnsupportedStructureError(
        "{0} is {1} bytes, past the {2}-byte source cap; refusing before "
        "reading it into memory.".format(context, size, max_bytes),
        kind="source-too-large",
    )


def _read_stream_bounded(stream, max_bytes, context):
    chunks = []
    total = 0
    while total <= max_bytes:
        chunk = stream.read(min(_READ_CHUNK, max_bytes + 1 - total))
        if not chunk:
            break
        if not isinstance(chunk, (bytes, bytearray, memoryview)):
            raise TypeError("{0} must be opened in binary mode".format(context))
        chunk = bytes(chunk)
        chunks.append(chunk)
        total += len(chunk)
    if total > max_bytes:
        _too_large(context, total, max_bytes)
    return b"".join(chunks)


def read_bounded(source, *, max_bytes=MAX_SOURCE_BYTES,
                 context="workbook package"):
    """Read bytes without any unbounded ``read()`` call.

    Paths and seekable streams are size-checked before allocation. Stream
    cursors are restored even when the size check or read refuses.
    """
    if isinstance(source, (bytes, bytearray, memoryview)):
        data = bytes(source)
        if len(data) > max_bytes:
            _too_large(context, len(data), max_bytes)
        return data

    if hasattr(source, "read"):
        if not hasattr(source, "seek") or not hasattr(source, "tell"):
            raise TypeError(
                "{0} file-like sources must be seekable".format(context))
        try:
            position = source.tell()
        except (OSError, ValueError) as exc:
            raise TypeError(
                "{0} file-like sources must be seekable".format(context)
            ) from exc
        try:
            source.seek(0, os.SEEK_END)
            size = source.tell()
            if size > max_bytes:
                _too_large(context, size, max_bytes)
            source.seek(0)
            return _read_stream_bounded(source, max_bytes, context)
        finally:
            source.seek(position)

    path = os.fspath(source)
    size = os.path.getsize(path)
    if size > max_bytes:
        _too_large(context, size, max_bytes)
    with open(path, "rb") as handle:
        return _read_stream_bounded(handle, max_bytes, context)
