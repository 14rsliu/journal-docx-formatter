#!/usr/bin/env python3
"""
Journal DOCX Formatter — Water Research / Elsevier style

Reformats a .docx file to match Water Research (Elsevier) submission format:
- Font: 12pt Times New Roman throughout
- Line spacing: double (2.0)
- Margins: 1 inch (2.54 cm) all sides
- Headings: bold, numbered (1. / 1.1 / 1.1.1), left-aligned
- Captions: bold, left-aligned ("Fig. 1." / "Table 1")
- Figures/tables: centered
- Body text: justified
- Continuous line numbering (not handled by python-docx)

Reference: https://www.sciencedirect.com/journal/water-research/publish/guide-for-authors

Usage:
    python format_journal.py input.docx [--output output.docx]
"""

import argparse
import re
import sys
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn


FONT_NAME = "Times New Roman"

# Patterns that identify caption paragraphs
CAPTION_PATTERNS = [
    re.compile(r'^(Figure|Fig\.?|Table|Tab\.?)\s*\d', re.IGNORECASE),
    re.compile(r'^(图|表)\s*\d'),
]


def is_caption(text):
    """Check if paragraph text looks like a figure/table caption."""
    text = text.strip()
    for pat in CAPTION_PATTERNS:
        if pat.match(text):
            return True
    return False


def is_heading(para):
    """Check if paragraph is a heading (H1-H3)."""
    style = para.style.name if para.style else ""
    if style.startswith("Heading"):
        try:
            level = int(style.replace("Heading", "").strip())
            return level if level <= 3 else 0
        except ValueError:
            return 0
    return 0


def is_image_paragraph(para):
    """Check if paragraph contains an image."""
    for run in para.runs:
        if run._element.findall(qn("w:drawing")) or run._element.findall(qn("w:pict")):
            return True
    # Also check inline shapes via XML
    if para._element.findall(".//" + qn("w:drawing")):
        return True
    if para._element.findall(".//" + qn("w:pict")):
        return True
    return False


def set_run_font(run, font_name=FONT_NAME, size=None, bold=None, color=None):
    """Set font properties on a run."""
    run.font.name = font_name
    r = run._element
    rPr = r.find(qn("w:rPr"))
    if rPr is None:
        rPr = r.makeelement(qn("w:rPr"), {})
        r.insert(0, rPr)
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("w:rFonts"), {})
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)
    rFonts.set(qn("w:cs"), font_name)

    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def set_paragraph_alignment(para, alignment):
    """Set paragraph alignment."""
    para.paragraph_format.alignment = alignment


def format_table_fonts(table, font_name=FONT_NAME, size=None):
    """Set font for all text in a table and center the table."""
    # Center table itself
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = tbl.makeelement(qn("w:tblPr"), {})
        tbl.insert(0, tblPr)
    jc = tblPr.find(qn("w:jc"))
    if jc is None:
        jc = tblPr.makeelement(qn("w:jc"), {})
        tblPr.append(jc)
    jc.set(qn("w:val"), "center")

    # Format cell text
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                set_paragraph_alignment(para, WD_ALIGN_PARAGRAPH.CENTER)
                for run in para.runs:
                    set_run_font(run, font_name, size=size)


def set_margins(doc, top=1.0, bottom=1.0, left=1.0, right=1.0):
    """Set page margins in inches for all sections."""
    for section in doc.sections:
        section.top_margin = Inches(top)
        section.bottom_margin = Inches(bottom)
        section.left_margin = Inches(left)
        section.right_margin = Inches(right)


def format_document(docx_path, output_path=None,
                    body_size=12, line_spacing=2.0, space_after=0,
                    margin=1.0):
    """
    Apply Water Research / Elsevier journal formatting to a .docx file.

    Defaults match Water Research submission requirements:
    - 12pt Times New Roman
    - Double-spaced (2.0)
    - 1-inch margins
    - Numbered headings, bold
    - Captions bold, left-aligned
    - Body justified
    """
    doc = Document(docx_path)

    # Set page margins
    set_margins(doc, top=margin, bottom=margin, left=margin, right=margin)

    for para in doc.paragraphs:
        heading_level = is_heading(para)
        text = para.text.strip()

        # Set line spacing for all paragraphs
        para.paragraph_format.line_spacing = line_spacing
        if space_after > 0:
            para.paragraph_format.space_after = Pt(body_size * space_after)
        else:
            para.paragraph_format.space_after = Pt(0)

        if heading_level:
            # Headings: bold, left-aligned, same size as body
            set_paragraph_alignment(para, WD_ALIGN_PARAGRAPH.LEFT)
            para.paragraph_format.space_before = Pt(12)
            for run in para.runs:
                set_run_font(run, size=body_size, bold=True)

        elif is_caption(text):
            # Captions: bold, left-aligned (Water Research style)
            set_paragraph_alignment(para, WD_ALIGN_PARAGRAPH.LEFT)
            for run in para.runs:
                set_run_font(run, size=body_size, bold=True)

        elif is_image_paragraph(para):
            # Image paragraphs: centered
            set_paragraph_alignment(para, WD_ALIGN_PARAGRAPH.CENTER)
            for run in para.runs:
                set_run_font(run, size=body_size)

        else:
            # Body text: justified
            set_paragraph_alignment(para, WD_ALIGN_PARAGRAPH.JUSTIFY)
            for run in para.runs:
                set_run_font(run, size=body_size)

    # Format tables
    for table in doc.tables:
        format_table_fonts(table, size=body_size)

    # Update default style fonts
    for style in doc.styles:
        if hasattr(style, "font") and style.font is not None:
            style.font.name = FONT_NAME
            if hasattr(style, "element"):
                rPr = style.element.find(qn("w:rPr"))
                if rPr is not None:
                    rFonts = rPr.find(qn("w:rFonts"))
                    if rFonts is None:
                        rFonts = rPr.makeelement(qn("w:rFonts"), {})
                        rPr.insert(0, rFonts)
                    rFonts.set(qn("w:ascii"), FONT_NAME)
                    rFonts.set(qn("w:hAnsi"), FONT_NAME)
                    rFonts.set(qn("w:eastAsia"), FONT_NAME)
                    rFonts.set(qn("w:cs"), FONT_NAME)

    save_path = output_path or docx_path
    doc.save(save_path)
    print(f"Formatted: {save_path}")
    print(f"  Style: Water Research / Elsevier")
    print(f"  Font: {FONT_NAME}, {body_size}pt")
    print(f"  Line spacing: {line_spacing}x")
    print(f"  Margins: {margin} inch")
    print(f"  Headings: bold, left-aligned")
    print(f"  Captions: bold, left-aligned")
    print(f"  Tables: centered")
    print(f"  Body: justified")


def main():
    parser = argparse.ArgumentParser(description="Format .docx for journal submission")
    parser.add_argument("input", help="Input .docx file")
    parser.add_argument("--output", help="Output path (default: overwrite input)")
    parser.add_argument("--body_size", type=float, default=12, help="Font size in pt (default: 12)")
    parser.add_argument("--line_spacing", type=float, default=2.0, help="Line spacing (default: 2.0 double)")
    parser.add_argument("--space_after", type=float, default=0, help="Space after paragraphs in lines (default: 0)")
    parser.add_argument("--margin", type=float, default=1.0, help="Page margins in inches (default: 1.0)")
    args = parser.parse_args()

    format_document(
        args.input,
        output_path=args.output,
        body_size=args.body_size,
        line_spacing=args.line_spacing,
        space_after=args.space_after,
        margin=args.margin,
    )


if __name__ == "__main__":
    main()
