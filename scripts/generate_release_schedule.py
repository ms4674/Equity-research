"""Generate the quarterly release schedule Word document for the Equity Tracker.

Usage: python3 scripts/generate_release_schedule.py
Output: docs/Equity_Tracker_Quarterly_Release_Schedule.docx
"""

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

ACCENT = RGBColor(0x1F, 0x4E, 0x79)      # dark blue
HEADER_BG = "1F4E79"
ALT_ROW_BG = "DCE6F1"
STATUS_COLORS = {
    "Released": RGBColor(0x2E, 0x7D, 0x32),
    "In Progress": RGBColor(0xE6, 0x7E, 0x22),
    "Planned": RGBColor(0x60, 0x60, 0x60),
}


def set_cell_background(cell, color_hex):
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:fill"), color_hex)
    cell._tc.get_or_add_tcPr().append(shading)


def style_header_row(table):
    for cell in table.rows[0].cells:
        set_cell_background(cell, HEADER_BG)
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.size = Pt(10)


def add_release_table(doc, rows):
    headers = ["Release", "Version", "Target Date", "Key Deliverables", "Status"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
    style_header_row(table)

    for row_index, row_data in enumerate(rows):
        row = table.add_row()
        for col_index, value in enumerate(row_data):
            cell = row.cells[col_index]
            cell.text = value
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)
                    if col_index == 4 and value in STATUS_COLORS:
                        run.font.color.rgb = STATUS_COLORS[value]
                        run.font.bold = True
        if row_index % 2 == 1:
            for cell in row.cells:
                set_cell_background(cell, ALT_ROW_BG)

    widths = [Inches(0.9), Inches(0.8), Inches(1.1), Inches(3.0), Inches(1.0)]
    for row in table.rows:
        for cell, width in zip(row.cells, widths):
            cell.width = width
    return table


def add_heading(doc, text, level):
    heading = doc.add_heading(text, level=level)
    for run in heading.runs:
        run.font.color.rgb = ACCENT
    return heading


def build_document():
    doc = Document()

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)

    # --- Title page ---
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Equity Tracker")
    run.font.size = Pt(36)
    run.font.bold = True
    run.font.color.rgb = ACCENT

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run("Quarterly Release Schedule — 2026")
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = meta.add_run(
        "\nDocument Owner: Product Management\n"
        "Version: 1.0\n"
        "Last Updated: September 15, 2026\n"
        "Status: Approved"
    )
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    doc.add_page_break()

    # --- Overview ---
    add_heading(doc, "1. Overview", 1)
    doc.add_paragraph(
        "This document defines the 2026 quarterly release schedule for the Equity "
        "Tracker, the platform used to monitor portfolios, watchlists, and market "
        "data for equity research. Each quarter delivers one major release, "
        "followed by a stabilization patch. Dates are targets and may shift by up "
        "to two weeks based on QA results and market-data vendor dependencies."
    )
    doc.add_paragraph(
        "Release cadence: one major release per quarter (weeks 10–11 of the "
        "quarter), with a patch release approximately three weeks later. Hotfixes "
        "are released as needed outside this cadence."
    )

    # --- Quarterly schedules ---
    add_heading(doc, "2. Release Schedule by Quarter", 1)

    quarters = [
        (
            "Q1 2026 — Foundation & Data Quality",
            "Theme: harden the core platform and improve the reliability of "
            "market data ingestion.",
            [
                ("R1.0", "v2.4.0", "Mar 12, 2026",
                 "Real-time price feed upgrade; ticker deduplication; "
                 "portfolio import from CSV/Excel; audit logging", "Released"),
                ("R1.1", "v2.4.1", "Apr 2, 2026",
                 "Stabilization patch: feed reconnection fixes, import "
                 "edge cases, performance tuning", "Released"),
            ],
        ),
        (
            "Q2 2026 — Analytics & Screening",
            "Theme: expand analytical tooling for research workflows.",
            [
                ("R2.0", "v2.5.0", "Jun 11, 2026",
                 "Custom screeners; factor exposure dashboard; earnings "
                 "calendar integration; export to PDF", "Released"),
                ("R2.1", "v2.5.1", "Jul 2, 2026",
                 "Stabilization patch: screener query optimizer, dashboard "
                 "rendering fixes", "Released"),
            ],
        ),
        (
            "Q3 2026 — Alerts & Collaboration",
            "Theme: keep analysts informed and enable shared research.",
            [
                ("R3.0", "v2.6.0", "Sep 10, 2026",
                 "Price and volume alerts (email/Slack); shared watchlists; "
                 "research note attachments; role-based access control", "Released"),
                ("R3.1", "v2.6.1", "Oct 1, 2026",
                 "Stabilization patch: alert delivery retries, permissions "
                 "edge cases", "In Progress"),
            ],
        ),
        (
            "Q4 2026 — Intelligence & Scale",
            "Theme: automated insights and platform scalability.",
            [
                ("R4.0", "v2.7.0", "Dec 10, 2026",
                 "AI-assisted research summaries; anomaly detection on "
                 "holdings; API v2 for external integrations; dark mode", "Planned"),
                ("R4.1", "v2.7.1", "Dec 22, 2026",
                 "Stabilization patch: year-end close support, API "
                 "hardening", "Planned"),
            ],
        ),
    ]

    for name, theme, rows in quarters:
        add_heading(doc, name, 2)
        paragraph = doc.add_paragraph(theme)
        paragraph.paragraph_format.space_after = Pt(6)
        add_release_table(doc, rows)
        doc.add_paragraph()

    # --- Milestones ---
    add_heading(doc, "3. Key Milestones", 1)
    milestones = [
        ("Feb 20, 2026", "Q1 feature freeze and start of regression testing"),
        ("May 22, 2026", "Q2 feature freeze; screener beta opens to pilot users"),
        ("Aug 21, 2026", "Q3 feature freeze; alerts load test at 10x volume"),
        ("Nov 20, 2026", "Q4 feature freeze; AI summary accuracy review"),
        ("Dec 31, 2026", "2027 roadmap finalized and published"),
    ]
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Date"
    table.rows[0].cells[1].text = "Milestone"
    style_header_row(table)
    for index, (date, description) in enumerate(milestones):
        row = table.add_row()
        row.cells[0].text = date
        row.cells[1].text = description
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)
        if index % 2 == 1:
            for cell in row.cells:
                set_cell_background(cell, ALT_ROW_BG)
    for row in table.rows:
        row.cells[0].width = Inches(1.4)
        row.cells[1].width = Inches(5.4)

    doc.add_paragraph()

    # --- Release process ---
    add_heading(doc, "4. Release Process", 1)
    steps = [
        "Feature freeze occurs three weeks before each major release date; only "
        "bug fixes are merged after freeze.",
        "Release candidates are deployed to staging for a two-week QA and UAT "
        "cycle, including market-data replay testing.",
        "Go/no-go review is held two business days before the target date with "
        "Product, Engineering, and Research stakeholders.",
        "Production deployments occur outside U.S. market hours (after 5:00 PM "
        "ET), with a rollback plan validated in staging.",
        "A stabilization patch is scheduled roughly three weeks after each major "
        "release to address issues found in production.",
    ]
    for step in steps:
        doc.add_paragraph(step, style="List Number")

    add_heading(doc, "5. Revision History", 1)
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    for i, header in enumerate(["Version", "Date", "Author", "Changes"]):
        table.rows[0].cells[i].text = header
    style_header_row(table)
    row = table.add_row()
    for i, value in enumerate(
        ["1.0", "Sep 15, 2026", "Product Management", "Initial approved schedule"]
    ):
        row.cells[i].text = value
        for paragraph in row.cells[i].paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(10)

    return doc


def main():
    output_dir = Path(__file__).resolve().parent.parent / "docs"
    output_dir.mkdir(exist_ok=True)
    output_path = output_dir / "Equity_Tracker_Quarterly_Release_Schedule.docx"
    build_document().save(output_path)
    print(f"Saved {output_path}")


if __name__ == "__main__":
    main()
