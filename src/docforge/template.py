"""Read a Word template: find its sections, the guidance text in them, and placeholders."""

from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

TEMPLATE_SUFFIXES = (".docx", ".dotx")

# {{name}}, [Name], <Name>, «Name»  -- "[12]"-style citations are not placeholders.
PLACEHOLDER_RE = re.compile(
    r"\{\{[^{}\n]{1,80}\}\}|\[(?!\d+\])[^\[\]\n]{2,80}\]|<[^<>\n]{2,80}>|«[^«»\n]{1,80}»"
)
_HEADING_NAME_RE = re.compile(r"(?i)^(heading|rubrik|overskrift|otsikko|überschrift|titre)\s*(\d)$")


@dataclass
class Section:
    id: str
    level: int
    title: str
    heading: Paragraph
    body: list = field(default_factory=list)  # lxml elements (w:p / w:tbl) under the heading

    def guidance(self, doc) -> str:
        lines = []
        for el in self.body:
            if el.tag == qn("w:p"):
                text = Paragraph(el, doc).text.strip()
                if text:
                    lines.append(text)
            else:
                for row in Table(el, doc).rows:
                    lines.append(" | ".join(c.text.strip() for c in row.cells))
        return "\n".join(lines)


def open_template(path: Path):
    """Open a .docx or .dotx as a python-docx Document."""
    data = Path(path).read_bytes()
    if Path(path).suffix.lower() == ".dotx":
        # A .dotx only differs in its declared content type; rewrite it in memory.
        src = zipfile.ZipFile(io.BytesIO(data))
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
            for item in src.infolist():
                content = src.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    content = content.replace(
                        b"wordprocessingml.template.main+xml", b"wordprocessingml.document.main+xml"
                    )
                dst.writestr(item, content)
        data = out.getvalue()
    return Document(io.BytesIO(data))


def heading_level(p: Paragraph) -> int | None:
    """Heading level 1-9 of a paragraph, or None. Works for localized/custom heading styles."""
    style = p.style
    for _ in range(10):
        if style is None:
            return None
        m = _HEADING_NAME_RE.match(style.name or "")
        if m:
            return int(m.group(2))
        ppr = style.element.pPr
        if ppr is not None:
            ol = ppr.find(qn("w:outlineLvl"))
            if ol is not None and int(ol.get(qn("w:val"))) < 9:
                return int(ol.get(qn("w:val"))) + 1
        style = style.base_style
    return None


def parse_sections(doc) -> list[Section]:
    sections: list[Section] = []
    current: Section | None = None
    for el in doc.element.body.iterchildren():
        if el.tag == qn("w:p"):
            p = Paragraph(el, doc)
            level = heading_level(p)
            if level and p.text.strip():
                current = Section(f"s{len(sections) + 1}", level, p.text.strip(), p)
                sections.append(current)
                continue
        elif el.tag != qn("w:tbl"):
            continue  # sectPr, content controls, ... are left untouched
        if current is not None:
            current.body.append(el)
    return sections


def all_paragraphs(doc):
    """Every paragraph in body, tables, headers and footers."""

    def walk(container):
        for p in container.paragraphs:
            yield p
        for t in container.tables:
            for row in t.rows:
                for cell in row.cells:
                    yield from walk(cell)

    yield from walk(doc)
    seen = set()
    for s in doc.sections:
        for part in (s.header, s.footer, s.first_page_header, s.first_page_footer):
            if id(part._element) not in seen:
                seen.add(id(part._element))
                yield from walk(part)


def find_placeholders(doc) -> dict[str, str]:
    """Map each placeholder to the text of the first paragraph it appears in."""
    found: dict[str, str] = {}
    for p in all_paragraphs(doc):
        text = p.text
        for m in PLACEHOLDER_RE.finditer(text):
            found.setdefault(m.group(0), text.strip()[:200])
    return found


def describe(doc, name: str) -> str:
    """Text description of the template, as shown to the LLM (and by `docforge inspect`)."""
    out = [f'<template name="{name}">']
    for s in parse_sections(doc):
        out.append(f'<section id="{s.id}" level="{s.level}" title="{s.title}">')
        guidance = s.guidance(doc)
        out.append(f"<template_text>\n{guidance}\n</template_text>" if guidance else "<template_text/>")
        out.append("</section>")
    placeholders = find_placeholders(doc)
    if placeholders:
        out.append("<placeholders>")
        for ph, ctx in placeholders.items():
            out.append(f"- {ph!r} appears in: {ctx!r}")
        out.append("</placeholders>")
    out.append("</template>")
    return "\n".join(out)
