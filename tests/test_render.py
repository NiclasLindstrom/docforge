import io
import zipfile

from docx import Document

from docforge import template
from docforge.cli import main
from docforge.render import render

RESULT = {
    "sections": [
        {"id": "s1", "mode": "replace", "blocks": [
            {"type": "paragraph", "text": "On track, with **one** delay.", "items": [], "rows": []}]},
        {"id": "s2", "mode": "replace", "blocks": [
            {"type": "paragraph", "text": "Replace the old system.", "items": [], "rows": []}]},
        {"id": "s3", "mode": "replace", "blocks": []},
        {"id": "s4", "mode": "replace", "blocks": [
            {"type": "bullets", "text": "", "items": ["Rehearsal passed", "Pilot done"], "rows": []}]},
        {"id": "s5", "mode": "replace", "blocks": [
            {"type": "table", "text": "", "items": [], "rows": [["Risk", "Impact"], ["Key person", "High"]]}]},
        {"id": "s6", "mode": "replace", "blocks": [
            {"type": "numbered", "text": "", "items": ["Escalate vendor delay"], "rows": []},
            {"type": "subheading", "text": "Next steps", "items": [], "rows": []}]},
        {"id": "s7", "mode": "keep", "blocks": []},
    ],
    "fields": [
        {"placeholder": "[Project Name]", "value": "Aurora"},
        {"placeholder": "{{date}}", "value": "2026-11-11"},
        {"placeholder": "{{author}}", "value": ""},
    ],
    "notes": [],
}


def _project(tmp_path):
    assert main(["init", str(tmp_path), "--sample"]) == 0
    return tmp_path / "templates" / "example-report.docx"


def test_parse(tmp_path):
    doc = template.open_template(_project(tmp_path))
    sections = template.parse_sections(doc)
    assert [s.title for s in sections] == [
        "Summary", "Background", "Findings", "Progress", "Risks", "Recommendations", "Confidentiality"]
    assert sections[4].level == 2 and "table" in sections[4].guidance(doc)
    assert set(template.find_placeholders(doc)) == {"[Project Name]", "{{date}}", "{{author}}"}


def test_render(tmp_path):
    out = tmp_path / "out.docx"
    stats = render(_project(tmp_path), RESULT, out)
    assert stats["replaced"] == 6 and stats["kept"] == 1 and not stats["untouched"]
    doc = Document(str(out))
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "Aurora - Status Report" in text and "Date: 2026-11-11" in text
    assert "{{author}}" in text  # no value given: left alone
    assert "Summarize the status" not in text  # guidance removed
    assert "intended solely for the named recipients" in text  # boilerplate kept
    assert [p.text for p in doc.paragraphs if p.style.name == "List Bullet"] == ["Rehearsal passed", "Pilot done"]
    assert doc.tables[0].cell(1, 0).text == "Key person"
    # order: content sits under its own heading
    names = [p.text for p in doc.paragraphs]
    assert names.index("Summary") < names.index("On track, with one delay.") < names.index("Background")
    assert names.index("Recommendations") < names.index("Next steps") < names.index("Confidentiality")
    bold = [r.text for p in doc.paragraphs for r in p.runs if r.bold]
    assert "one" in bold


def test_dotx(tmp_path):
    src = _project(tmp_path)
    dotx = tmp_path / "t.dotx"
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dotx, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(b"document.main+xml", b"template.main+xml")
            zout.writestr(item, data)
    out = tmp_path / "out.docx"
    render(dotx, RESULT, out)
    assert "Aurora" in Document(str(out)).paragraphs[0].text
