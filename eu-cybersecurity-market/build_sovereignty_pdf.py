#!/usr/bin/env python3
"""Render EU_Cybersecurity_Sovereignty_2035.md into a 2-page PDF.

Keeps the Markdown file as the single source of truth: this script parses its
headings, bold (**...**) and italic (*...*) inline spans and paragraphs, and
forces a page break at each top-level '---' separator so the brief lands on two
pages. Re-run with:  python3 build_sovereignty_pdf.py
"""

from __future__ import annotations

import os
import re

from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (PageBreak, Paragraph, SimpleDocTemplate,
                                Spacer)

HERE = os.path.dirname(os.path.abspath(__file__))
MD_PATH = os.path.join(HERE, "EU_Cybersecurity_Sovereignty_2035.md")
PDF_PATH = os.path.join(HERE, "EU_Cybersecurity_Sovereignty_2035.pdf")


def inline(text: str) -> str:
    """Convert a minimal subset of Markdown inline syntax to ReportLab markup."""
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    text = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", r"<i>\1</i>", text)
    return text


def build():
    with open(MD_PATH, encoding="utf-8") as fh:
        lines = fh.read().splitlines()

    styles = getSampleStyleSheet()
    title = ParagraphStyle("title", parent=styles["Title"], fontSize=15,
                           spaceAfter=5, textColor="#1F3864")
    h2 = ParagraphStyle("h2", parent=styles["Heading2"], fontSize=11.5,
                        spaceBefore=2, spaceAfter=5, textColor="#2E5496")
    body = ParagraphStyle("body", parent=styles["BodyText"], fontSize=8.4,
                          leading=11.0, alignment=TA_JUSTIFY, spaceAfter=3.5)
    note = ParagraphStyle("note", parent=body, fontSize=7.3, textColor="#555555",
                          spaceAfter=5)

    flow = []
    for raw in lines:
        line = raw.rstrip()
        if line.strip() == "---":
            # Horizontal rules are layout hints only; the real page break is
            # forced before the "Page 2" heading so the brief is exactly 2 pages.
            continue
        if not line.strip():
            continue
        if line.startswith("# "):
            flow.append(Paragraph(inline(line[2:]), title))
        elif line.startswith("## "):
            if "Page 2" in line:
                flow.append(PageBreak())
            flow.append(Paragraph(inline(line[3:]), h2))
        elif line.startswith("*") and line.endswith("*") and "**" not in line:
            flow.append(Paragraph(inline(line), note))
        else:
            flow.append(Paragraph(inline(line), body))

    doc = SimpleDocTemplate(
        PDF_PATH, pagesize=A4,
        leftMargin=1.6 * cm, rightMargin=1.6 * cm,
        topMargin=1.3 * cm, bottomMargin=1.3 * cm,
        title="European Cybersecurity Sovereignty by 2035",
    )
    doc.build(flow)
    print(f"Wrote PDF: {PDF_PATH}")


if __name__ == "__main__":
    build()
