"""Load research files from the project folder into Claude API content blocks."""

from __future__ import annotations

import base64
from pathlib import Path

TEXT_SUFFIXES = {".txt", ".md", ".markdown", ".csv", ".tsv", ".json", ".xml", ".html", ".htm", ".yaml", ".yml", ".log", ".rst"}
IMAGE_TYPES = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif", ".webp": "image/webp"}


def _read_text(path: Path) -> str:
    data = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            return data.decode(encoding)
        except UnicodeError:
            continue
    return data.decode("utf-8", errors="replace")


def _docx_text(path: Path) -> str:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    doc = Document(str(path))
    lines = []
    for el in doc.element.body.iterchildren():
        if el.tag == qn("w:p"):
            lines.append(Paragraph(el, doc).text)
        elif el.tag == qn("w:tbl"):
            for row in Table(el, doc).rows:
                lines.append(" | ".join(c.text.strip() for c in row.cells))
    return "\n".join(lines)


def _xlsx_text(path: Path) -> str:
    from openpyxl import load_workbook

    wb = load_workbook(str(path), data_only=True, read_only=True)
    out = []
    for ws in wb.worksheets:
        out.append(f"## Sheet: {ws.title}")
        for row in ws.iter_rows(values_only=True):
            if any(v is not None for v in row):
                out.append("\t".join("" if v is None else str(v) for v in row))
    return "\n".join(out)


def _pdf_text(path: Path) -> str:
    from pypdf import PdfReader

    pages = [(page.extract_text() or "").strip() for page in PdfReader(str(path)).pages]
    if not any(pages):
        raise ValueError("no extractable text; scanned PDFs need the Claude API provider")
    return "\n\n".join(f"[page {i}]\n{t}" for i, t in enumerate(pages, 1) if t)


def _b64(path: Path) -> str:
    return base64.standard_b64encode(path.read_bytes()).decode("ascii")


def list_files(folder: Path) -> list[Path]:
    if not folder.is_dir():
        return []
    return sorted(p for p in folder.rglob("*") if p.is_file() and not p.name.startswith(("~$", ".")))


def load(folder: Path, native_pdf: bool = True, images: bool = True) -> tuple[list[dict], list[str], list[str]]:
    """Return (content blocks, names of loaded files, names of skipped files).

    native_pdf: send PDFs as files (Claude API). Otherwise their text is extracted locally.
    """
    blocks: list[dict] = []
    loaded, skipped = [], []
    for path in list_files(folder):
        name = path.relative_to(folder).as_posix()
        suffix = path.suffix.lower()
        try:
            if suffix in TEXT_SUFFIXES:
                text = _read_text(path)
            elif suffix == ".docx":
                text = _docx_text(path)
            elif suffix in (".xlsx", ".xlsm"):
                text = _xlsx_text(path)
            elif suffix == ".pdf" and not native_pdf:
                text = _pdf_text(path)
            elif suffix == ".pdf":
                blocks.append({
                    "type": "document",
                    "title": name,
                    "source": {"type": "base64", "media_type": "application/pdf", "data": _b64(path)},
                })
                loaded.append(name)
                continue
            elif suffix in IMAGE_TYPES and not images:
                skipped.append(f"{name} (images are turned off)")
                continue
            elif suffix in IMAGE_TYPES:
                blocks.append({"type": "text", "text": f"Image research file: {name}"})
                blocks.append({
                    "type": "image",
                    "source": {"type": "base64", "media_type": IMAGE_TYPES[suffix], "data": _b64(path)},
                })
                loaded.append(name)
                continue
            else:
                skipped.append(f"{name} (unsupported file type)")
                continue
        except Exception as exc:  # a broken file should not stop the run
            skipped.append(f"{name} ({exc})")
            continue
        blocks.append({"type": "text", "text": f'<document name="{name}">\n{text}\n</document>'})
        loaded.append(name)
    return blocks, loaded, skipped
