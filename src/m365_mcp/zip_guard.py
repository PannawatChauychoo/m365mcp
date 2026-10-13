"""Check an untrusted Office (OOXML) zip before a parser opens it.

Word and Excel files are zip archives. A zip's directory declares each
member's size, but an attacker controls those numbers, so this module
actually inflates every member, counting real bytes against a size limit
and a compression-ratio limit, and caps the member count. As defence in
depth beside the parsers' own settings (python-docx and openpyxl parse with
entity resolution off), it also refuses an XML document type declaration in
the parts a reader may parse as XML: Office never writes one there, and
entity-expansion attacks need one. Readers choose a part's parser by its
content type, not its name, so "XML parts" means names ending ``.xml`` or
``.rels`` plus any part ``[Content_Types].xml`` gives an Office XML content
type. Image and HTML content types are not scanned, because SVG images and
altChunk HTML legitimately start with ``<!DOCTYPE`` and are never parsed as
XML. The declaration is matched in UTF-8 and UTF-16 (not UTF-32).

Callers turn ``ZipGuardError`` into their own reason type.
"""

from __future__ import annotations

import io
import posixpath
import re
import zipfile

_DOCTYPE_PATTERNS = tuple(
    "<!DOCTYPE".encode(encoding) for encoding in ("utf-8", "utf-16-le", "utf-16-be")
)
_TAIL_BYTES = max(len(pattern) for pattern in _DOCTYPE_PATTERNS) - 1
_XML_PART_SUFFIXES = (".xml", ".rels")
_CONTENT_TYPES_PART = "[Content_Types].xml"
_CONTENT_TYPE_ENTRY = re.compile(rb"<(Override|Default)\b([^>]*)>", re.IGNORECASE)
_ATTRIBUTE = re.compile(rb"""([A-Za-z]+)\s*=\s*(?:"([^"]*)"|'([^']*)')""")
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

    ratio_limit = max(len(content), 1) * max_ratio
    limit = min(max_inflated_bytes, ratio_limit)
    if ratio_limit < max_inflated_bytes:
        too_large = ZipGuardError(
            f"{label} expands to more than {max_ratio}x its compressed size "
            f"({limit} bytes), a compression ratio Office files do not reach; "
            "it was not opened"
        )
    else:
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
        xml_parts, xml_extensions = _declared_xml_parts(archive, limit)
        inflated = 0
        for info in members:
            if info.is_dir():
                continue
            name = info.filename.lower()
            scan_doctype = (
                name.endswith(_XML_PART_SUFFIXES)
                or name.lstrip("/") in xml_parts
                or posixpath.splitext(name)[1].lstrip(".") in xml_extensions
            )
            try:
                with archive.open(info) as member:
                    tail = b""
                    while chunk := member.read(_CHUNK_BYTES):
                        inflated += len(chunk)
                        if inflated > limit:
                            raise too_large
                        if not scan_doctype:
                            continue
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


def _is_office_xml(content_type: str) -> bool:
    content_type = content_type.strip().lower()
    if content_type.startswith(("image/", "text/html")) or "xhtml" in content_type:
        return False
    return content_type.endswith(("+xml", "/xml"))


def _declared_xml_parts(
    archive: zipfile.ZipFile, limit: int
) -> tuple[set[str], set[str]]:
    """Part names and extensions ``[Content_Types].xml`` gives an XML type.

    Read without an XML parser: the part itself is still scanned for a
    DOCTYPE in the main loop, and a damaged or missing part only means the
    name-based rule applies. Names are lower-case without a leading ``/``."""
    try:
        with archive.open(_CONTENT_TYPES_PART) as member:
            raw = member.read(limit + 1)
    except Exception:
        return set(), set()
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        try:
            raw = raw.decode("utf-16").encode("utf-8")
        except UnicodeDecodeError:
            return set(), set()
    parts: set[str] = set()
    extensions: set[str] = set()
    for kind, body in _CONTENT_TYPE_ENTRY.findall(raw):
        attributes = {
            key.decode("ascii").lower(): (double or single).decode("utf-8", "replace")
            for key, double, single in _ATTRIBUTE.findall(body)
        }
        if not _is_office_xml(attributes.get("contenttype", "")):
            continue
        if kind.lower() == b"override" and attributes.get("partname"):
            parts.add(attributes["partname"].lower().lstrip("/"))
        elif kind.lower() == b"default" and attributes.get("extension"):
            extensions.add(attributes["extension"].lower().lstrip("."))
    return parts, extensions
