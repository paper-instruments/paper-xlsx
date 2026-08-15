# paper-xlsx: streaming source reads

"""Read package sources without imposing workbook eligibility caps."""

import os


_READ_CHUNK = 1024 * 1024


def _read_stream(stream, context):
    chunks = []
    while True:
        chunk = stream.read(_READ_CHUNK)
        if not chunk:
            return b"".join(chunks)
        if not isinstance(chunk, (bytes, bytearray, memoryview)):
            raise TypeError("{0} must be opened in binary mode".format(context))
        chunks.append(bytes(chunk))


def read_source_bytes(source, *, context="workbook package"):
    """Read bytes in chunks while preserving a seekable stream's position."""
    if isinstance(source, (bytes, bytearray, memoryview)):
        return bytes(source)
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
            source.seek(0)
            return _read_stream(source, context)
        finally:
            source.seek(position)
    with open(os.fspath(source), "rb") as handle:
        return _read_stream(handle, context)
