---
name: journal-docx-formatter
description: "Two-in-one document skill: (1) Convert Markdown to Word (.docx) via pandoc with optional journal formatting; (2) Reformat existing .docx to Water Research / Elsevier journal submission style (12pt Times New Roman, double-spaced, 1-inch margins). Use when converting .md to .docx, formatting Word documents for journal submission, or mentions journal format / 期刊格式."
allowed-tools: Bash, Read, Write, Glob
---

# Journal DOCX Toolkit

Two tools for academic paper preparation:
1. **`convert_md_to_docx.py`** — Markdown to DOCX via pandoc (with optional `--journal` flag)
2. **`format_journal.py`** — Reformat existing DOCX to Water Research / Elsevier style

## Requirements

```bash
pip install python-docx
# pandoc required for md-to-docx conversion
# macOS: brew install pandoc | Linux: sudo apt-get install pandoc | Windows: choco install pandoc
```

---

## Tool 1: Markdown to DOCX

```bash
# Basic conversion
python ~/.claude/skills/journal-docx-formatter/convert_md_to_docx.py paper.md paper.docx

# Convert AND apply journal formatting in one step
python ~/.claude/skills/journal-docx-formatter/convert_md_to_docx.py paper.md paper.docx --journal

# With TOC and metadata
python ~/.claude/skills/journal-docx-formatter/convert_md_to_docx.py paper.md paper.docx --journal \
  --toc --title "My Paper" --author "John Doe"

# Batch convert
python ~/.claude/skills/journal-docx-formatter/convert_md_to_docx.py --batch input_dir/ output_dir/ --journal
```

| Option | Description |
|--------|-------------|
| `--journal` | Apply Water Research formatting after conversion |
| `--toc` | Generate table of contents |
| `--reference-doc FILE` | Use custom Word template |
| `--title TEXT` | Document title metadata |
| `--author TEXT` | Document author metadata |
| `--date TEXT` | Document date metadata |
| `--batch` | Batch conversion mode |

---

## Tool 2: Journal Formatter

Reformat any existing `.docx` to Water Research / Elsevier submission style.

```bash
# Apply formatting (in-place)
python ~/.claude/skills/journal-docx-formatter/format_journal.py paper.docx

# Save to new file
python ~/.claude/skills/journal-docx-formatter/format_journal.py paper.docx --output paper_formatted.docx

# Custom settings
python ~/.claude/skills/journal-docx-formatter/format_journal.py paper.docx --line_spacing 1.5 --body_size 11
```

| Option | Default | Description |
|--------|---------|-------------|
| `--output` | in-place | Output file path |
| `--body_size` | 12 | Font size in pt |
| `--line_spacing` | 2.0 | Line spacing (2.0 = double) |
| `--margin` | 1.0 | Page margins in inches |
| `--space_after` | 0 | Space after paragraphs (in lines) |

---

## Water Research / Elsevier Formatting Specs

| Item | Requirement |
|------|-------------|
| **Font** | 12pt Times New Roman |
| **Line spacing** | Double-spaced (2.0) |
| **Margins** | 1 inch (2.54 cm) all sides |
| **Headings** | Numbered (1. / 1.1 / 1.1.1), bold, left-aligned |
| **Body text** | Justified |
| **Captions** | Bold, left-aligned ("Fig. 1." / "Table 1") |
| **Figures / Tables** | Centered |
| **Abstract** | Max 250 words, unstructured |
| **Keywords** | 1-7 keywords |
| **Highlights** | 3-5 bullet points, max 85 chars each |
| **Citations** | Author-year Harvard style, e.g., (Smith et al., 2020) |
| **References** | Alphabetical: Author, I., Year. Title. J. Abbrev. Vol, Pages. DOI |
| **Word limit** | ~7,000 words (original research) |
| **Line numbering** | Continuous (add manually: Layout > Line Numbers > Continuous) |

Reference: https://www.sciencedirect.com/journal/water-research/publish/guide-for-authors

Note: Resources, Conservation and Recycling (RCR) follows the same Elsevier conventions.

---

## Typical Workflow

```bash
# Step 1: Write paper in Markdown
# Step 2: Convert + format in one command
python ~/.claude/skills/journal-docx-formatter/convert_md_to_docx.py paper.md paper.docx --journal

# Or two-step for more control:
python ~/.claude/skills/journal-docx-formatter/convert_md_to_docx.py paper.md paper.docx
python ~/.claude/skills/journal-docx-formatter/format_journal.py paper.docx
```
