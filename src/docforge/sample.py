"""Example template and research file for `docforge init --sample`."""

from pathlib import Path

from docx import Document


def write_sample(project: Path) -> None:
    doc = Document()
    doc.add_paragraph("[Project Name] - Status Report", style="Title")
    doc.add_paragraph("Date: {{date}}")
    doc.add_paragraph("Author: {{author}}")

    doc.add_heading("Summary", level=1)
    doc.add_paragraph("Summarize the status in at most 150 words for a reader who reads nothing else.")
    doc.add_heading("Background", level=1)
    doc.add_paragraph("Describe why the project exists and what it is meant to deliver.")
    doc.add_heading("Findings", level=1)
    doc.add_heading("Progress", level=2)
    doc.add_paragraph("What has been done since the last report. Use a bullet list.")
    doc.add_heading("Risks", level=2)
    doc.add_paragraph("List the main risks in a table with the columns Risk, Likelihood, Impact, Mitigation.")
    doc.add_heading("Recommendations", level=1)
    doc.add_paragraph("Give numbered, concrete recommendations with an owner for each.")
    doc.add_heading("Confidentiality", level=1)
    doc.add_paragraph("This document is confidential and intended solely for the named recipients.")
    doc.save(str(project / "templates" / "example-report.docx"))

    (project / "research" / "example-notes.md").write_text(
        "# Notes, project Aurora (warehouse system replacement)\n\n"
        "- Goal: replace the 2009 warehouse system before the support contract ends 2027-03-31.\n"
        "- Done in October: data migration rehearsal 2 passed (99.2 % of records), scanner pilot at the Malmo site.\n"
        "- Delayed: integration with the invoicing system, vendor is 3 weeks late.\n"
        "- Risk: only one person (Karin) knows the old pricing rules.\n"
        "- Budget: 4.1 of 6.0 MSEK spent.\n"
        "- Project manager: Jonas Ek.\n",
        encoding="utf-8",
    )
