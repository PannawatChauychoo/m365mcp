"""Check an untrusted Office (OOXML) zip before a parser opens it.

Word and Excel files are zip archives. A zip's directory declares each
member's size, but an attacker controls those numbers, so this module
actually inflates every member, counting real bytes against a size limit
and a compression-ratio limit, and caps the member count. It also refuses
any XML document type declaration: Office never writes one, and
entity-expansion attacks need one, so this guard holds whichever XML parser
the reader library ends up using. The declaration is matched in UTF-8 and
UTF-16, the encodings XML parsers detect.

Callers turn ``ZipGuardError`` into their own reason type.
"""

from __future__ import annotations

import io
import zipfile

_DOCTYPE_PATTERNS = tuple(
    "<!DOCTYPE".encode(encoding) for encoding in ("utf-8", "utf-16-le", "utf-16-be")
)
_TAIL_BYTES = max(len(pattern) for pattern in _DOCTYPE_PATTERNS) - 1
_CHUNK_BYTES = 1 << 20


class ZipGuardError(Exception):
    """The archive is unsafe or damaged; the message is the reason."""


class NotAZipError(ZipGuardError):
    """The bytes are not a zip archive at all."""


def check_zip_archive(
    content: bytes,
    *,
    max_inflated_bytes: int,
    max_ratio: int,
    max_members: int,
    label: str,
) -> None:
    """Raise ZipGuardError unless ``content`` is a zip that is safe to parse.

    The inflated size may not exceed ``max_inflated_bytes`` nor ``max_ratio``
    times the compressed size. ``label`` names the file kind in messages
    ("Word document", "Workbook")."""
    try:
        archive = zipfile.ZipFile(io.BytesIO(content))
    except zipfile.BadZipFile as error:
        raise NotAZipError(f"{label} is not a valid zip archive") from error
    except Exception as error:
        raise ZipGuardError(f"{label} archive is damaged: {error}") from error

    limit = min(max_inflated_bytes, max(len(content), 1) * max_ratio)
    too_large = ZipGuardError(f"{label} expands to more than {limit} bytes and was not opened")
    with archive:
        members = archive.infolist()
        if len(members) > max_members:
            raise ZipGuardError(
                f"{label} archive has more than {max_members} parts and was not opened"
            )
        # Declared sizes are only a cheap early refusal; the real check is
        # the inflated byte count below.
        if sum(info.file_size for info in members) > limit:
            raise too_large
        inflated = 0
        for info in members:
            if info.is_dir():
                continue
            try:
                with archive.open(info) as member:
                    tail = b""
                    while chunk := member.read(_CHUNK_BYTES):
                        inflated += len(chunk)
                        if inflated > limit:
                            raise too_large
                        window = tail + chunk
                        if any(pattern in window for pattern in _DOCTYPE_PATTERNS):
                            raise ZipGuardError(
                                f"{label} contains an XML document type "
                                "declaration, which Office never writes; it "
                                "was not opened"
                            )
                        tail = window[-_TAIL_BYTES:]
            except ZipGuardError:
                raise
            except Exception as error:
                raise ZipGuardError(f"{label} archive is damaged: {error}") from error
