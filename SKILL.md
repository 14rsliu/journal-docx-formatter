---
name: journal-docx-formatter
description: "Two-in-one document skill: (1) Convert Markdown to Word (.docx) via pandoc with optional journal formatting; (2) Reformat existing .docx to Water Research / Elsevier journal submission style (11pt Times New Roman, single-spaced, 1-inch margins). Handles oMath formula styling, equation numbering, bullet normalization, and text cleanup. Use when converting .md to .docx, formatting Word documents for journal submission, or mentions journal format / 期刊格式."
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

# Custom settings (e.g. 12pt body for some journals)
python ~/.claude/skills/journal-docx-formatter/format_journal.py paper.docx --line_spacing 1.5 --body_size 12

# Double-spaced for some journals
python ~/.claude/skills/journal-docx-formatter/format_journal.py paper.docx --line_spacing 2.0 --space_after 0
```

| Option | Default | Description |
|--------|---------|-------------|
| `--output` | in-place | Output file path |
| `--body_size` | 11 | Body font size in pt |
| `--author_size` | 11 | Author/affiliation font size in pt |
| `--table_size` | 11 | Table font size in pt |
| `--line_spacing` | 1.0 | Line spacing (1.0 = single) |
| `--margin` | 1.0 | Page margins in inches |
| `--space_after` | 0.5 | Space after paragraphs (in lines) |

---

## Water Research / Elsevier Formatting Specs

| Item | Requirement |
|------|-------------|
| **Body font** | 11pt Times New Roman, black, no bold |
| **Title** | 14pt, **bold**, centered |
| **Author info** | 11pt Times New Roman, centered; corresponding author marked with `*` superscript after name; email → footnote |
| **Line spacing** | Single (1.0), 0.5 line space after |
| **Margins** | 1 inch (2.54 cm) all sides |
| **Headings** | Numbered (1. / 1.1 / 1.1.1), **bold**, left-aligned |
| **Body text** | Justified, no bold, no italic, dashes/tildes removed, number ranges converted to "X to Y" (e.g., `15–30%` → "15 to 30%"); references/acknowledgements preserve dashes |
| **Captions** | Label only bold ("Fig. 1." / "Table 1"), no italic, centered |
| **Tables** | Three-line (三线表), compact (zero top/bottom cell margin, minimal left/right, single line spacing, vertical center), 100% width, content-aware column widths, header bold, 11pt Times New Roman |
| **Figures** | Centered |
| **Inline formulas** | Times New Roman 11pt, **not italic** (plain/roman style), flattened from oMath to plain text when simple |
| **Display equations** | Times New Roman 11pt, **italic**, preserved as oMathPara, right-aligned equation numbers `(1)`–`(N)` |
| **Formula cleanup** | Orphan `±`, `~` symbols from oMath flattening auto-removed; field codes (`w:instrText`) also cleaned |
| **References** | Page break before section |
| **Abstract title** | 11pt, centered, bold ("Abstract") |
| **Abstract body** | 11pt, justified |
| **Keywords** | "Keywords:" bold, comma-space separated, 1-7 keywords |
| **Highlights** | Bullet "•" prefix, 3-5 points, max 85 chars each |
| **Citations** | Author-year Harvard style, e.g., (Smith et al., 2020) |
| **References** | Alphabetical: Author, I., Year. Title. J. Abbrev. Vol, Pages. DOI |
| **Word limit** | ~7,000 words (original research) |
| **Line numbering** | Continuous (add manually: Layout > Line Numbers > Continuous) |

Reference: https://www.sciencedirect.com/journal/water-research/publish/guide-for-authors

Note: Resources, Conservation and Recycling (RCR) follows the same Elsevier conventions.

---

## Writing & Content Guidelines

These rules apply when writing or editing paper content (not just formatting):

| Rule | Details |
|------|---------|
| **Section 1 = Introduction** | First numbered section must be "1. Introduction" — no extra subtitle or description |
| **Section 2 = Methods** | Section 2 covers all methodology: overall approach/system architecture → problem formulation → data & preprocessing → model architecture → FL strategies → DP mechanism → personalization → training config. Data processing belongs here unless the paper is specifically about data processing experiments |
| **Section 3 = Results only** | Section 3 contains only experiment results and analysis — no setup/config/data description (those go in Section 2) |
| **Minimize symbols in prose** | Body text should read as natural language. Avoid `~`, `≈`, `<`, `>`, `±`, excessive parentheses, and mathematical shorthand. Rewrite with words: `~50` → "approximately 50", `>90%` → "above 90%", `<0.5%` → "less than 0.5%", `≈29,102` → "approximately 29,102", `±0.88%` → "with a standard deviation of 0.88%". Parenthetical asides like `(ε ≤ 53)` should be woven into the sentence: "where ε does not exceed 53". The formatter auto-removes tildes and orphan `±`, but the writer should proactively use natural phrasing |
| **Bullet lists use "•"** | When itemizing in body text, use bullet "•" (not dashes or numbers); numbered lists (1. 2. 3.) and Word XML numbering (`w:numPr`) are auto-converted to "•" bullets |
| **Introduction: no results** | Introduction should NOT mention or preview experimental findings from later sections |
| **Tables: three-line (三线表)** | Only 3 horizontal rules: top of header (thick), bottom of header (thin), bottom of table (thick); no vertical lines |
| **Table font** | All table content in Times New Roman, same size as body |
| **Table width** | Tables fill 100% page width with evenly distributed columns |
| **Minimize cell wrapping** | Keep table cell content concise to avoid text wrapping across lines |
| **Contributions in prose** | Introduction's contributions section must use flowing paragraph text with transitions, not bullet lists |
| **Bullet style** | Bullet/list paragraphs: "•" prefix, no first-line indent, **justified** (same as body text) |
| **Section intro sentence** | Every (sub)section must open with a transitional sentence before the first table/figure; never place a table caption or figure immediately after a heading |
| **Minimize parentheses** | Avoid excessive parentheses in body text; integrate info into the sentence flow instead of (parenthetical asides) |
| **Reduce relative improvements** | In highlights and abstract, avoid "by X percentage points", "within X pp of", "X pp improvement" phrasing; state absolute accuracy values instead |
| **Avoid possessive 's** | Do not use apostrophe-s ('s) for possession. Rewrite using "of" or other phrasing: "the model's accuracy" → "the accuracy of the model", "Hong Kong's climate" → "the climate of Hong Kong", "FedAvg's performance" → "the performance of FedAvg". Contractions (it's, don't) are already banned in academic writing |
| **Prefer "this study"** | Use "this study" as the default subject when referring to the current work. Avoid "this paper", "this work", "we", "our approach", etc. Examples: "This study proposes…", "The results of this study indicate…", "This study adopts federated learning to…" |
| **Clean writing style** | Review body text for unnecessary symbols: dashes (auto-removed), tildes (auto-removed), orphan `±` (auto-removed), `≈`/`<`/`>` (rewrite as words), excessive parentheses (rewrite as clauses), redundant content. The goal is fluent academic prose, not shorthand notation |

---

## Processing Pipeline (Pass Order)

The formatter processes the document in a specific order — **pass ordering matters** because later passes depend on earlier ones:

| Pass | Name | What it does |
|------|------|-------------|
| 1.0 | Page setup | Margins, sections, page size |
| 1.1 | Heading detection | Identify and format section headings |
| 1.2 | Caption/figure | Format captions, center figures |
| 1.3 | Title/author/abstract | Format front matter, email → footnote |
| 1.4a | Corresponding author `*` | Add `*` superscript after corresponding author name |
| 1.4b | Affiliation superscripts | Add letter superscripts (a, b, c) for affiliations |
| 1.4c | Bullet normalization | Convert numbered/lettered/Word-XML lists to `•` bullets |
| 1.5 | Bullet merging | Merge Introduction bullet lists into prose paragraphs |
| **1.6** | **oMath processing** | Flatten simple inline/table oMath to plain text (TNR 11pt, not italic); set display oMath to italic; add right-aligned equation numbers `(1)`–`(N)`; clean field codes |
| 2.0 | Paragraph formatting | Body font, alignment, spacing, dash/tilde/orphan `±` removal, reduce parentheses, reduce relative improvements |
| 3.0 | Table formatting | Three-line style, TNR 11pt, full width, header bold |

**Key constraint:** Pass 1.6 (oMath) must run **before** Pass 2.0 (text cleanup), so that symbols flattened from formulas (like orphan `±`) get cleaned up by Pass 2.

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
