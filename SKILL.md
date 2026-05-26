---
name: journal-docx-formatter
description: "Format .docx files to Water Research / Elsevier journal submission style: 12pt Times New Roman, double-spaced, 1-inch margins, bold numbered headings, bold captions, centered figures/tables, justified body text. Use when user asks to format a Word document for journal submission, apply journal formatting, or mentions 期刊格式."
allowed-tools: Bash, Read, Write, Glob
---

# Journal DOCX Formatter (Water Research / Elsevier)

Reformat an existing `.docx` file to match Water Research (Elsevier) submission requirements.

## Water Research Formatting Requirements

Based on the official Guide for Authors:

| Item | Requirement |
|------|-------------|
| **Font** | 12pt Times New Roman (all text) |
| **Line spacing** | Double-spaced (2.0) |
| **Margins** | 1 inch (2.54 cm) all sides |
| **Headings** | Numbered (1. / 1.1 / 1.1.1), bold, left-aligned |
| **Body text** | Justified |
| **Captions** | Bold, left-aligned ("Fig. 1." / "Table 1") |
| **Figures** | Centered |
| **Tables** | Centered |
| **Abstract** | Max 250 words, unstructured |
| **Keywords** | 1-7 keywords |
| **Highlights** | 3-5 bullet points, max 85 characters each |
| **Citations** | Author-year Harvard style, e.g., (Smith et al., 2020) |
| **References** | Alphabetical; Author, I., Year. Title. J. Abbrev. Vol, Pages. DOI |
| **Word limit** | ~7,000 words for original research |
| **Line numbering** | Continuous (must be added manually in Word) |

Reference: https://www.sciencedirect.com/journal/water-research/publish/guide-for-authors

Note: Resources, Conservation and Recycling (RCR) follows the same Elsevier formatting conventions.

## Requirements

```bash
pip install python-docx
```

## Instructions

1. **Run the formatter**:
   ```bash
   python ~/.claude/skills/journal-docx-formatter/format_journal.py input.docx
   ```
   This overwrites the file in-place. To save to a different path:
   ```bash
   python ~/.claude/skills/journal-docx-formatter/format_journal.py input.docx --output formatted.docx
   ```

2. **Options** (defaults match Water Research requirements):
   - `--body_size 12` — Font size in pt (default: 12)
   - `--line_spacing 2.0` — Line spacing (default: 2.0 double)
   - `--margin 1.0` — Page margins in inches (default: 1.0)
   - `--space_after 0` — Space after paragraphs in lines (default: 0)
   - `--output path.docx` — Save to different file instead of in-place

3. **Manual steps** (not automated by the script):
   - Add continuous line numbering: Layout > Line Numbers > Continuous
   - Verify heading numbering matches 1. / 1.1 / 1.1.1 format
   - Check abstract word count <= 250
   - Verify highlights format (3-5 bullets, <= 85 chars each)

## Examples

```bash
# Water Research default formatting
python ~/.claude/skills/journal-docx-formatter/format_journal.py paper.docx

# Save to new file
python ~/.claude/skills/journal-docx-formatter/format_journal.py paper.docx --output paper_wr.docx

# Custom settings (e.g., single-spaced for draft review)
python ~/.claude/skills/journal-docx-formatter/format_journal.py paper.docx --line_spacing 1.0 --body_size 11
```
