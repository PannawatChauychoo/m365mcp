"""Turn downloaded file bytes into text the model can read.

Used by sharepoint_files.py to read Word, PDF, and plain-text files stored in
SharePoint/OneDrive. Everything is parsed in memory; nothing is written to disk.

- Plain text (.md, .txt, .csv, ...) is decoded as UTF-8 (UTF-16 when a BOM says so).
- Word (.docx, .docm, .dotx, .dotm) is rendered as Markdown in document order:
  headings become `#`, list paragraphs become `-` / `1.` items, and tables become
  Markdown tables. Macros in .docm/.dotm are never run; only the text is read.
- PDF returns the text layer per page via pypdf; a scan has no text layer.

Both parsers are optional imports, like PdfReader in microsoft_graph.py, so a
missing dependency becomes a readable reason instead of an import error.
Unreadable files raise DocumentTextError, which callers turn into an
unsupportedReason. Extraction stops once it has more than ``max_chars``
characters (and PDFs after MAX_PDF_PAGES pages), so a large or crafted file
cannot keep a worker thread busy for long.
"""

from __future__ import annotations

import codecs
import io
import re
import zipfile
from dataclasses import dataclass

from .zip_guard import NotAZipError, ZipGuardError, check_zip_archive

try:
    from pypdf import PdfReader
except Exception:  # pragma: no cover - pypdf is an optional import at module load
    PdfReader = None  # type: ignore[assignment]

try:
    from docx import Document as DocxDocument
    from docx.table import Table as DocxTable
except Exception:  # pragma: no cover - python-docx is an optional import at module load
    DocxDocument = None  # type: ignore[assignment]
    DocxTable = None  # type: ignore[assignment]

TEXT_EXTENSIONS = {
    "csv",
    "htm",
    "html",
    "ics",
    "json",
    "log",
    "markdown",
    "md",
    "tsv",
    "txt",
    "xml",
    "yaml",
    "yml",
}
WORD_EXTENSIONS = {"docx", "docm", "dotx", "dotm"}
PDF_EXTENSIONS = {"pdf"}
WORKBOOK_EXTENSIONS = {"xlsx", "xlsm", "xltx", "xltm"}
LEGACY_WORD_EXTENSIONS = {"doc", "dot", "rtf", "odt"}

# A .docx is a zip; refuse archives that inflate past this (counted by actually
# inflating every member, see zip_guard.py) before python-docx opens them.
MAX_DOCX_INFLATED_BYTES = 50_000_000
MAX_DOCX_COMPRESSION_RATIO = 100
MAX_DOCX_ARCHIVE_MEMBERS = 10_000
# Pages past this are not read, even when they hold no text (a scan).
MAX_PDF_PAGES = 500

# python-docx only opens a package whose main part has the .docx content type.
# Macro-enabled documents and templates use the same WordprocessingML under a
# different content type, so it is swapped for the .docx one before parsing.
_DOCX_MAIN_CONTENT_TYPE = (
    b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
)
_OTHER_WORD_MAIN_CONTENT_TYPES = (
    b"application/vnd.ms-word.document.macroEnabled.main+xml",
    b"application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
    b"application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
)
_CONTENT_TYPES_PART = "[Content_Types].xml"

_HEADING_STYLE = re.compile(r"^Heading (\d)$")


class DocumentTextError(Exception):
    """A file could not be turned into text; the message is the reason."""


@dataclass
class ExtractedText:
    """At most ``max_chars`` characters of a file's text. ``truncated`` is
    true when the file holds more than was returned."""

    content: str
    truncated: bool


def extract_text(kind: str, content: bytes, *, max_chars: int) -> ExtractedText:
    """Extract text for a ``file_kind`` result, bounded by ``max_chars``.

    Any parser failure becomes DocumentTextError."""
    extract = {
        "word": extract_docx_markdown,
        "pdf": extract_pdf_text,
        "text": lambda data, max_chars: (decode_text(data), False),
    }[kind]
    try:
        text, stopped_early = extract(content, max_chars=max_chars)
    except DocumentTextError:
        raise
    except Exception as error:  # pragma: no cover - parsers wrap their own
        raise DocumentTextError(f"Could not read file: {error}") from error
    return ExtractedText(
        content=text[:max_chars],
        truncated=stopped_early or len(text) > max_chars,
    )


def file_kind(extension: str | None) -> str | None:
    """Return 'text', 'word', or 'pdf' for a readable extension, else None."""
    ext = (extension or "").lower().lstrip(".")
    if ext in TEXT_EXTENSIONS:
        return "text"
    if ext in WORD_EXTENSIONS:
        return "word"
    if ext in PDF_EXTENSIONS:
        return "pdf"
    return None


def unreadable_reason(extension: str | None) -> str:
    """Explain why a file with this extension is not read, and what to use instead."""
    ext = (extension or "").lower().lstrip(".")
    if ext in WORKBOOK_EXTENSIONS:
        return (
            "Excel workbooks are read with the workbook tools: pass driveId and "
            "itemId to workbook_list_worksheets or workbook_get_range"
        )
    if ext in LEGACY_WORD_EXTENSIONS:
        return f".{ext} files are not supported; save the document as .docx to read it"
    if not ext:
        return "File has no extension, so its type is unknown and it was not read"
    return f".{ext} files are not readable as text by this MCP server"


def decode_text(content: bytes) -> str:
    if content.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
        return content.decode("utf-16", errors="replace")
    return content.decode("utf-8-sig", errors="replace")


def extract_pdf_text(content: bytes, *, max_chars: int) -> tuple[str, bool]:
    """Return the PDF's text layer and whether pages were left unread."""
    if PdfReader is None:
        raise DocumentTextError("PDF text extraction requires the pypdf package")
    try:
        reader = PdfReader(io.BytesIO(content))
        if reader.is_encrypted and not reader.decrypt(""):
            raise DocumentTextError("PDF is password-protected")
        page_count = len(reader.pages)
        pages: list[str] = []
        length = 0
        stopped_early = False
        for index in range(page_count):
            if index >= MAX_PDF_PAGES or length > max_chars:
                stopped_early = True
                break
            text = (reader.pages[index].extract_text() or "").strip()
            if text:
                pages.append(f"--- Page {index + 1} ---\n{text}")
                length += len(pages[-1]) + 2
    except DocumentTextError:
        raise
    except Exception as error:
        raise DocumentTextError(f"Could not read PDF: {error}") from error
    return "\n\n".join(pages), stopped_early


def extract_docx_markdown(content: bytes, *, max_chars: int) -> tuple[str, bool]:
    """Return the document as Markdown and whether blocks were left unread."""
    if DocxDocument is None:
        raise DocumentTextError("Word text extraction requires the python-docx package")
    _check_docx_archive(content)
    try:
        document = DocxDocument(io.BytesIO(_as_docx_package(content)))
        blocks: list[tuple[bool, str]] = []
        length = 0
        stopped_early = False
        for block in document.iter_inner_content():
            if length > max_chars:
                stopped_early = True
                break
            if isinstance(block, DocxTable):
                text, is_list_item = _table_markdown(block), False
            else:
                text, is_list_item = _paragraph_markdown(block)
            if text:
                blocks.append((is_list_item, text))
                length += len(text) + 2
    except Exception as error:
        raise DocumentTextError(f"Could not read Word document: {error}") from error

    # Keep consecutive list items on adjacent lines so they render as one list.
    parts: list[str] = []
    previous_was_list_item = False
    for is_list_item, text in blocks:
        if parts:
            parts.append("\n" if is_list_item and previous_was_list_item else "\n\n")
        parts.append(text)
        previous_was_list_item = is_list_item
    return "".join(parts), stopped_early


def _check_docx_archive(content: bytes) -> None:
    try:
        check_zip_archive(
            content,
            max_inflated_bytes=MAX_DOCX_INFLATED_BYTES,
            max_ratio=MAX_DOCX_COMPRESSION_RATIO,
            max_members=MAX_DOCX_ARCHIVE_MEMBERS,
            label="Word document",
        )
    except NotAZipError as error:
        raise DocumentTextError(
            "File is not a valid .docx; it may be truncated, "
            "password-protected, or a legacy .doc renamed to .docx"
        ) from error
    except ZipGuardError as error:
        raise DocumentTextError(str(error)) from error


def _as_docx_package(content: bytes) -> bytes:
    """Return the package with a .docm/.dotx/.dotm main part relabelled as .docx."""
    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        try:
            content_types = archive.read(_CONTENT_TYPES_PART)
        except KeyError:
            return content
        relabelled = content_types
        for content_type in _OTHER_WORD_MAIN_CONTENT_TYPES:
            relabelled = relabelled.replace(content_type, _DOCX_MAIN_CONTENT_TYPE)
        if relabelled == content_types:
            return content
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w", zipfile.ZIP_STORED) as copy:
            for info in archive.infolist():
                data = (
                    relabelled
                    if info.filename == _CONTENT_TYPES_PART
                    else archive.read(info)
                )
                copy.writestr(info.filename, data)
    return output.getvalue()


def _paragraph_markdown(paragraph) -> tuple[str, bool]:
    text = paragraph.text.strip()
    if not text:
        return "", False
    style = paragraph.style.name if paragraph.style is not None else ""
    style = style or ""
    if style == "Title":
        return f"# {text}", False
    heading = _HEADING_STYLE.match(style)
    if heading:
        return f"{'#' * min(int(heading.group(1)), 6)} {text}", False
    if style.startswith("List Number"):
        return f"1. {text}", True
    properties = paragraph._p.pPr
    if style.startswith("List") or (properties is not None and properties.numPr is not None):
        return f"- {text}", True
    return text, False


def _table_markdown(table) -> str:
    rows = [[_cell_markdown(cell.text) for cell in row.cells] for row in table.rows]
    rows = [row for row in rows if any(row)]
    if not rows:
        return ""
    width = max(len(row) for row in rows)
    rows = [row + [""] * (width - len(row)) for row in rows]
    lines = [
        "| " + " | ".join(rows[0]) + " |",
        "| " + " | ".join(["---"] * width) + " |",
    ]
    lines.extend("| " + " | ".join(row) + " |" for row in rows[1:])
    return "\n".join(lines)


def _cell_markdown(text: str) -> str:
    return " ".join(text.split()).replace("|", "\\|")
