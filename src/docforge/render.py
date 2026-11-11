"""Write generated content into the template and save the result as .docx."""

from __future__ import annotations

import re
from collections import Counter
from pathlib import Path

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .template import all_paragraphs, heading_level, open_template, parse_sections

_INLINE_RE = re.compile(r"(\*\*.+?\*\*|\*[^*\s](?:.*?[^*\s])?\*)")
_INSTRUCTION_STYLE_RE = re.compile(r"(?i)instruc|guidance|guide|help|hint|prompt|comment|list|toc|caption|title")


def _style(doc, kind, *names):
    for name in names:
        try:
            style = doc.styles[name]
        except KeyError:
            continue
        if style.type == kind:
            return style
    return None


def pick_body_style(doc, override: str | None = None):
    """Paragraph style for generated body text."""
    if override:
        style = _style(doc, WD_STYLE_TYPE.PARAGRAPH, override)
        if style is None:
            raise SystemExit(f"Body style {override!r} does not exist in the template.")
        return style
    counts = Counter()
    for el in doc.element.body.iterchildren(qn("w:p")):
        p = Paragraph(el, doc)
        if p.text.strip() and not heading_level(p) and not _INSTRUCTION_STYLE_RE.search(p.style.name or ""):
            counts[p.style.name] += 1
    if counts:
        return doc.styles[counts.most_common(1)[0][0]]
    return _style(doc, WD_STYLE_TYPE.PARAGRAPH, "Body Text", "Normal")


def _add_inline(paragraph, text: str):
    """Add text with **bold** / *italic* markers as formatted runs."""
    for part in _INLINE_RE.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            paragraph.add_run(part[2:-2]).bold = True
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            paragraph.add_run(part[1:-1]).italic = True
        else:
            paragraph.add_run(part)


class _Writer:
    """Inserts new block elements one after another, starting after `anchor`."""

    def __init__(self, doc, section, body_style, table_style):
        self.doc, self.section, self.anchor = doc, section, section.heading._p
        self.body_style, self.table_style = body_style, table_style

    def _place(self, element):
        self.anchor.addnext(element)
        self.anchor = element

    def paragraph(self, text, style=None, prefix=""):
        p = self.doc.add_paragraph(style=style or self.body_style)
        if prefix:
            p.add_run(prefix)
        _add_inline(p, text)
        self._place(p._p)
        return p

    def block(self, block: dict):
        kind, text = block.get("type"), (block.get("text") or "").strip()
        P = WD_STYLE_TYPE.PARAGRAPH
        if kind == "paragraph":
            if text:
                self.paragraph(text)
        elif kind == "subheading":
            if not text:
                return
            style = _style(self.doc, P, f"Heading {min(self.section.level + 1, 9)}")
            p = self.paragraph(text if style else f"**{text}**", style)
            if style is None:
                p.paragraph_format.keep_with_next = True
        elif kind in ("bullets", "numbered"):
            style = _style(self.doc, P, "List Bullet" if kind == "bullets" else "List Number")
            for i, item in enumerate(block.get("items") or [], 1):
                prefix = "" if style else ("• " if kind == "bullets" else f"{i}. ")
                self.paragraph(item, style, prefix)
        elif kind == "table":
            rows = [r for r in (block.get("rows") or []) if r]
            if not rows:
                return
            if text:
                self.paragraph(f"**{text}**")
            table = self.doc.add_table(rows=len(rows), cols=max(len(r) for r in rows))
            if self.table_style is not None:
                table.style = self.table_style
            for r, row in enumerate(rows):
                for c, value in enumerate(row):
                    p = table.cell(r, c).paragraphs[0]
                    _add_inline(p, f"**{value}**" if r == 0 and "*" not in value else value)
            self._place(table._tbl)
            self.paragraph("")  # keeps adjacent tables apart


def _clear_section_body(section):
    for el in section.body:
        ppr = el.find(qn("w:pPr")) if el.tag == qn("w:p") else None
        if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
            # This paragraph carries a section break (page layout); keep it, drop its text.
            for child in list(el):
                if child is not ppr:
                    el.remove(child)
        else:
            el.getparent().remove(el)


def _replace_placeholders(paragraph, mapping: dict[str, str]):
    if not any(k in paragraph.text for k in mapping):
        return
    runs = paragraph.runs
    for run in runs:
        for k, v in mapping.items():
            if k in run.text:
                run.text = run.text.replace(k, v)
    joined = "".join(r.text for r in runs)
    if runs and any(k in joined for k in mapping):  # placeholder split across runs
        for k, v in mapping.items():
            joined = joined.replace(k, v)
        runs[0].text = joined
        for run in runs[1:]:
            run.text = ""


def _refresh_fields_on_open(doc):
    """If the template has a table of contents, ask Word to update it when the file is opened."""
    xml = doc.element.xml
    if "TOC " not in xml:
        return
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        el = OxmlElement("w:updateFields")
        el.set(qn("w:val"), "true")
        settings.append(el)


def render(template_path: Path, result: dict, out_path: Path, body_style: str | None = None) -> dict:
    """Fill `template_path` with `result` (the LLM's JSON) and save to `out_path`."""
    doc = open_template(template_path)
    sections = {s.id: s for s in parse_sections(doc)}
    style = pick_body_style(doc, body_style)
    T = WD_STYLE_TYPE.TABLE
    table_style = doc.tables[0].style if doc.tables else None
    table_style = table_style or _style(doc, T, "Table Grid")

    stats = {"replaced": 0, "kept": 0, "unknown": [], "untouched": []}
    done = set()
    for item in result.get("sections", []):
        section = sections.get(item.get("id"))
        if section is None:
            stats["unknown"].append(item.get("id"))
            continue
        done.add(section.id)
        if item.get("mode") == "keep":
            stats["kept"] += 1
            continue
        _clear_section_body(section)
        writer = _Writer(doc, section, style, table_style)
        for block in item.get("blocks", []):
            writer.block(block)
        stats["replaced"] += 1
    stats["untouched"] = [f"{s.id} {s.title}" for s in sections.values() if s.id not in done]

    mapping = {
        f["placeholder"]: f["value"]
        for f in result.get("fields", [])
        if f.get("placeholder") and f.get("value") and f["value"] != f["placeholder"]
    }
    if mapping:
        for p in all_paragraphs(doc):
            _replace_placeholders(p, mapping)

    _refresh_fields_on_open(doc)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return stats
