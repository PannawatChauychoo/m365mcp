from __future__ import annotations

import io
import struct
import zipfile

import pytest

from m365_mcp.zip_guard import NotAZipError, ZipGuardError, check_zip_archive

LIMITS = {"max_inflated_bytes": 10_000_000, "max_ratio": 100, "max_members": 50}


def _zip(members: dict[str, bytes], compression: int = zipfile.ZIP_DEFLATED) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression) as archive:
        for name, data in members.items():
            archive.writestr(name, data)
    return buffer.getvalue()


def _check(content: bytes, **overrides: int) -> None:
    check_zip_archive(content, **{**LIMITS, **overrides}, label="Word document")


def test_accepts_an_ordinary_archive() -> None:
    _check(_zip({"[Content_Types].xml": b"<Types/>", "word/document.xml": b"<w:document/>" * 100}))


def test_refuses_bytes_that_are_not_a_zip() -> None:
    with pytest.raises(NotAZipError):
        _check(b"\xd0\xcf\x11\xe0 legacy compound file")


def test_counts_inflated_bytes_against_the_compression_ratio() -> None:
    # 2 MB of zeros deflates to a few KB: well under the absolute limit, far
    # over 100x the compressed size.
    content = _zip({"word/document.xml": b"\0" * 2_000_000})
    with pytest.raises(ZipGuardError, match="compression ratio"):
        _check(content)


def test_does_not_trust_declared_sizes() -> None:
    data = b"<w:t>" + b"x" * 200_000 + b"</w:t>"
    content = bytearray(_zip({"word/document.xml": data}, zipfile.ZIP_STORED))
    # Rewrite the declared uncompressed size (local header and central
    # directory) to 10 bytes; the real member is still 200 KB.
    real_size = struct.pack("<I", len(data))
    assert content.count(real_size) >= 2
    content = content.replace(real_size, struct.pack("<I", 10))
    with pytest.raises(ZipGuardError):
        _check(bytes(content))


def test_caps_member_count() -> None:
    content = _zip({f"part{index}.xml": b"<a/>" for index in range(6)})
    with pytest.raises(ZipGuardError, match="more than 5 parts"):
        _check(content, max_members=5)


@pytest.mark.parametrize("encoding", ["utf-8", "utf-16-le", "utf-16-be"])
def test_refuses_a_doctype_in_any_xml_encoding(encoding: str) -> None:
    xml = '<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "a">]><x>&a;</x>'.encode(encoding)
    with pytest.raises(ZipGuardError, match="document type declaration"):
        _check(_zip({"word/document.xml": xml}))


def test_finds_a_doctype_split_across_read_chunks() -> None:
    from m365_mcp import zip_guard

    prefix = b" " * (zip_guard._CHUNK_BYTES - 4)
    content = _zip({"word/document.xml": prefix + b"<!DOCTYPE x><x/>"})
    with pytest.raises(ZipGuardError, match="document type declaration"):
        _check(content, max_ratio=100_000)


SVG_WITH_DOCTYPE = (
    b'<?xml version="1.0"?>\n<!DOCTYPE svg PUBLIC "-//W3C//DTD SVG 1.1//EN" '
    b'"http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd">\n<svg/>'
)


@pytest.mark.parametrize(
    ("name", "data"),
    [
        ("word/media/image1.svg", SVG_WITH_DOCTYPE),
        ("word/afchunk.html", b"<!DOCTYPE html><html><body>Hi</body></html>"),
    ],
)
def test_ignores_a_doctype_in_media_parts(name: str, data: bytes) -> None:
    _check(_zip({"word/document.xml": b"<w:document/>", name: data}))


def test_refuses_a_doctype_in_a_relationships_part() -> None:
    content = _zip({"word/_rels/document.xml.rels": b"<!DOCTYPE x><Relationships/>"})
    with pytest.raises(ZipGuardError, match="document type declaration"):
        _check(content)


def test_counts_media_parts_against_the_size_limits() -> None:
    content = _zip({"word/document.xml": b"<w:document/>", "word/media/big.svg": b"\0" * 2_000_000})
    with pytest.raises(ZipGuardError, match="compression ratio"):
        _check(content)


CONTENT_TYPES = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    b'<Default Extension="xml" ContentType="application/xml"/>'
    b'<Default Extension="svg" ContentType="image/svg+xml"/>'
    b'<Default Extension="html" ContentType="text/html"/>'
    b"%s"
    b"</Types>"
)
MAIN_XML = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"


def test_refuses_a_doctype_in_a_part_declared_as_xml_under_another_name() -> None:
    # Readers pick a part's parser by content type, not name: a media-named
    # part overridden to the main document type is parsed as XML.
    override = f'<Override PartName="/word/media/evil.svg" ContentType="{MAIN_XML}"/>'
    content = _zip(
        {
            "[Content_Types].xml": CONTENT_TYPES % override.encode(),
            "word/media/evil.svg": b"<!DOCTYPE x [<!ENTITY a 'b'>]><w:document/>",
        }
    )
    with pytest.raises(ZipGuardError, match="document type declaration"):
        _check(content)


def test_refuses_a_doctype_in_an_extension_declared_as_xml() -> None:
    default = f"<Default Extension='bin' ContentType='{MAIN_XML}'/>"
    content = _zip(
        {
            "[Content_Types].xml": CONTENT_TYPES % default.encode(),
            "word/document.bin": b"<!DOCTYPE x><w:document/>",
        }
    )
    with pytest.raises(ZipGuardError, match="document type declaration"):
        _check(content)


def test_declared_image_and_html_parts_may_carry_a_doctype() -> None:
    _check(
        _zip(
            {
                "[Content_Types].xml": CONTENT_TYPES % b"",
                "word/document.xml": b"<w:document/>",
                "word/media/image1.svg": SVG_WITH_DOCTYPE,
                "word/afchunk.html": b"<!DOCTYPE html><html><body>Hi</body></html>",
            }
        )
    )
