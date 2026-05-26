---
name: journal-docx-formatter
description: "Two-in-one document skill: (1) Convert Markdown to Word (.docx) with Chinese font support (FangSong), citation superscript, and proper formatting; (2) Reformat existing .docx to Water Research / Elsevier journal submission style (12pt Times New Roman, double-spaced, 1-inch margins). Use when converting .md to .docx, formatting Word documents for journal submission, or mentions 期刊格式 / Word / DOCX."
allowed-tools: Bash, Read, Write, Glob
---

# Journal DOCX Toolkit

Two tools in one skill:
1. **Markdown to DOCX** — Convert `.md` to `.docx` with Chinese template support
2. **Journal Formatter** — Reformat existing `.docx` to Water Research / Elsevier style

## Requirements

```bash
# pandoc (required for md-to-docx conversion)
# macOS: brew install pandoc | Linux: sudo apt-get install pandoc | Windows: choco install pandoc

# python-docx (required for both tools)
pip install python-docx
```

---

## Tool 1: Markdown to DOCX Conversion

Convert Markdown files to Word documents with proper formatting.

### Features
- Chinese font template: FangSong (仿宋) for all Chinese text, Times New Roman for English
- Black font color, 1.5x line spacing, first-line indent (24pt)
- Heading spacing after (1 line / 12pt), no italic headings
- Citation superscript: auto-converts citation numbers (e.g., `文献1。` → `文献¹。`)
- Supports TOC, metadata, batch conversion, custom templates

### Usage

```bash
# Basic conversion (default Chinese template)
python ~/.claude/skills/journal-docx-formatter/convert_md_to_docx.py input.md output.docx

# With table of contents
python ~/.claude/skills/journal-docx-formatter/convert_md_to_docx.py input.md output.docx --toc

# With metadata
python ~/.claude/skills/journal-docx-formatter/convert_md_to_docx.py paper.md paper.docx \
  --title "研究报告" --author "张三" --date "2025-11-20"

# Without Chinese template
python ~/.claude/skills/journal-docx-formatter/convert_md_to_docx.py input.md output.docx --no-chinese-template

# Disable citation superscript
python ~/.claude/skills/journal-docx-formatter/convert_md_to_docx.py input.md output.docx --no-superscript-citations

# Batch convert
python ~/.claude/skills/journal-docx-formatter/convert_md_to_docx.py --batch input_dir/ output_dir/
```

### Options

| Option | Description |
|--------|-------------|
| `--toc` | Generate table of contents |
| `--no-chinese-template` | Don't use FangSong font template |
| `--no-superscript-citations` | Don't convert citation numbers to superscript |
| `--reference-doc FILE` | Use custom Word template |
| `--title TEXT` | Document title metadata |
| `--author TEXT` | Document author metadata |
| `--date TEXT` | Document date metadata |
| `--batch` | Batch conversion mode |
| `--pattern PATTERN` | File pattern for batch mode (default: `*.md`) |

### Chinese Template Defaults
- Font: FangSong (仿宋) for Chinese, Times New Roman for English
- Font size: 12pt body
- Font color: Black
- Line spacing: 1.5x
- First-line indent: 24pt (2 Chinese characters)
- Heading spacing: 12pt after each heading
- Heading style: Bold, no italic

---

## Tool 2: Journal Formatter (Water Research / Elsevier)

Reformat an existing `.docx` to match journal submission requirements.

### Water Research Formatting Requirements

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

### Usage

```bash
# Water Research default formatting (in-place)
python ~/.claude/skills/journal-docx-formatter/format_journal.py paper.docx

# Save to new file
python ~/.claude/skills/journal-docx-formatter/format_journal.py paper.docx --output paper_wr.docx

# Custom settings (e.g., single-spaced draft)
python ~/.claude/skills/journal-docx-formatter/format_journal.py paper.docx --line_spacing 1.0 --body_size 11
```

### Options

| Option | Default | Description |
|--------|---------|-------------|
| `--output` | (in-place) | Output path |
| `--body_size` | 12 | Font size in pt |
| `--line_spacing` | 2.0 | Line spacing (2.0 = double) |
| `--margin` | 1.0 | Page margins in inches |
| `--space_after` | 0 | Space after paragraphs in lines |

### Manual Steps (not automated)
- Add continuous line numbering: Layout > Line Numbers > Continuous
- Verify heading numbering matches 1. / 1.1 / 1.1.1 format
- Check abstract word count <= 250
- Verify highlights format (3-5 bullets, <= 85 chars each)

---

## Typical Workflow

```bash
# Step 1: Convert Chinese markdown to docx
python ~/.claude/skills/journal-docx-formatter/convert_md_to_docx.py paper_chinese.md paper_chinese.docx

# Step 2: Format for journal submission
python ~/.claude/skills/journal-docx-formatter/format_journal.py paper_chinese.docx --output paper_submission.docx
```

## Supporting Files

- **convert_md_to_docx.py** — Markdown to DOCX converter with pandoc
- **format_journal.py** — Journal formatting script
- **create_chinese_template.py** — Utility to create custom Chinese templates
- **fix_heading_fonts.py** — Post-processing for font and spacing fixes
- **chinese_template.docx** — Default Chinese font template
