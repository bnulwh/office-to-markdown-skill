---
name: office-to-markdown
description: >-
  Convert Office files (.pptx, .docx, .xlsx) to Markdown format. Auto-detects
  file type by extension and uses the appropriate converter. Supports PowerPoint
  presentations (slide titles as headings, tables with merged cells, multiple
  text boxes sorted by visual position, inline formatting, bullet lists, speaker
  notes, heuristic title detection for custom templates), Word documents
  (heading styles, Chinese numbering detection, tables with merged cells,
  ordered/unordered lists, inline formatting, hyperlinks), and Excel
  spreadsheets (multiple sheets, merged cells with rowspan/colspan, basic
  formatting). Use this skill whenever the user wants to convert any Office
  file to Markdown, extract content from .pptx/.docx/.xlsx as .md, or transform
  Office documents into plain-text format. Also triggers when the user mentions
  "office to markdown", "pptx to md", "docx to md", "xlsx to md",
  "PPT转Markdown", "Word转Markdown", "Excel转Markdown", or "Office文件转Markdown".
version: 1.0.0
---

# Office to Markdown Converter

Unified converter for Office files (.pptx, .docx, .xlsx) to well-structured Markdown.

## Quick Start

```bash
# Unified entry point (auto-detects file type)
python scripts/office2md.py input.pptx -o output.md
python scripts/office2md.py input.docx -o output.md
python scripts/office2md.py input.xlsx -o output.md

# Or use individual converters directly
python scripts/pptx2md.py input.pptx -o output.md
python scripts/docx2md.py input.docx -o output.md
python scripts/xlsx2md.py input.xlsx -o output.md
```

Scripts are at `scripts/` relative to this skill's directory.

## Usage

### Installation

```bash
# Install via skills CLI (recommended)
npx skills add bnulwh/office-to-md

# Or clone the repository
git clone https://github.com/bnulwh/office-to-md.git

# Install Python dependencies
pip install python-pptx python-docx openpyxl
```

### Command-Line Options

All converters support the same interface:

```bash
python scripts/office2md.py <input> [-o <output>]
```

| Option | Description | Default |
|--------|-------------|---------|
| `input` | Input file path (.pptx, .docx, .xlsx) | Required |
| `-o, --output` | Output Markdown file path | `<input>.md` |

### Examples

**Convert a PowerPoint presentation:**

```bash
# Output to presentation.md (same directory as input)
python scripts/office2md.py slides.pptx

# Specify output path
python scripts/office2md.py slides.pptx -o output/slides.md
```

**Convert a Word document:**

```bash
python scripts/office2md.py report.docx -o report.md
```

**Convert an Excel spreadsheet:**

```bash
# All sheets are converted
python scripts/office2md.py data.xlsx -o data.md
```

**Use individual converters directly:**

```bash
# PowerPoint only
python scripts/pptx2md.py presentation.pptx -o output.md

# Word only
python scripts/docx2md.py document.docx -o output.md

# Excel only
python scripts/xlsx2md.py spreadsheet.xlsx -o output.md
```

### Python Module Usage

Import converters as Python modules for programmatic use:

```python
import sys
sys.path.insert(0, 'path/to/skill/scripts')

from office2md import convert

# Auto-detect format and convert
convert("input.pptx", "output.md")
convert("input.docx", "output.md")
convert("input.xlsx", "output.md")

# Or use individual converters
from pptx2md import convert as pptx_convert
from docx2md import convert as docx_convert
from xlsx2md import convert as xlsx_convert

pptx_convert("slides.pptx", "slides.md")
```

### Batch Conversion

Convert multiple files using a shell loop:

```bash
# Bash/Linux/macOS
for file in *.pptx; do
    python scripts/office2md.py "$file" -o "output/${file%.pptx}.md"
done

# Windows PowerShell
Get-ChildItem *.pptx | ForEach-Object {
    python scripts/office2md.py $_.FullName -o "output\$($_.BaseName).md"
}
```

## Supported Formats

| Format | Extension | Converter | Key Features |
|---|---|---|---|
| PowerPoint | .pptx, .pptm | pptx2md.py | Slide titles, tables, text boxes, bullets, speaker notes |
| Word | .docx, .docm | docx2md.py | Headings, lists, tables, Chinese numbering, hyperlinks |
| Excel | .xlsx, .xlsm | xlsx2md.py | Multiple sheets, merged cells, basic formatting |

## What It Converts

### PowerPoint (.pptx)

| PPT element | Markdown output |
|---|---|
| Slide title placeholder | `## Title` (one heading per slide) |
| Heuristic title detection | `## Title` (for PPTs without standard placeholders) |
| Subtitle placeholder | Plain text below the heading |
| Text boxes (multiple) | Plain text, sorted by visual position (top->bottom, left->right) |
| Tables (merged cells) | HTML `<table>` with `rowspan` / `colspan` |
| Bullet / numbered lists | `- item` / `1. item` |
| **Bold** / *Italic* / ~~Strike~~ / `Code` | Standard Markdown inline syntax |
| Speaker notes | `> **Notes:** text` (block quote) |
| Slide separators | `---` between slides |

### Word (.docx)

| Word element | Markdown output |
|---|---|
| Title / Heading 1-6 | `#` through `######` |
| 附件N / 一、二、/ （一）（二） | `#` / `##` / `###` (auto-detected) |
| Centered bold large text | `#` (document title) |
| Tables (with merged cells) | HTML `<table>` with `rowspan` / `colspan` |
| Bullet / numbered lists | `- item` / `1. item` (nested supported) |
| Body Text with numbering patterns | `- item` (auto-detected) |
| **Bold** / *Italic* / ~~Strike~~ / `Code` | Standard Markdown inline syntax |
| Hyperlinks | `[text](url)` |

### Excel (.xlsx)

| Excel element | Markdown output |
|---|---|
| Worksheet (multi-sheet) | `## SheetName` + HTML table per sheet |
| Merged cells | HTML `<table>` with `rowspan` / `colspan` |
| **Bold** / *Italic* cells | `<strong>` / `<em>` in HTML table cells |
| Multi-line cell content | `<br>` separated within cells |
| Sheet separators | `---` between sheets |

## How It Works

### Title / Heading Detection

- **PowerPoint**: First checks placeholder `idx` (Title=0, Center Title=2, Subtitle=1,3). Falls back to heuristic detection: largest font size in top half of slide, filtering out pure-numeric text. For cover pages, searches entire slide.
- **Word**: 8-level fallback: explicit heading styles -> outlineLevel -> Chinese appendix markers -> centered bold large text -> Chinese numbering patterns -> font-size heuristic.
- **Excel**: Sheet names become `##` headings.

### Table Conversion

All three converters produce HTML `<table>` elements to support merged cells:
- **PowerPoint**: Reads raw `<a:tr>`/`<a:tc>` XML for `gridSpan`/`rowSpan`/`hMerge`/`vMerge`
- **Word**: Uses `_tc` element identity to detect vertical merges + `gridSpan` for horizontal merges
- **Excel**: Reads `merged_cells.ranges` from openpyxl to compute rowspan/colspan

## Pitfalls

- **PowerPoint**: Images/charts not extracted. SmartArt text is extracted but visual structure is lost. Heuristic title detection may not work for all custom templates.
- **Word**: Nested tables are flattened to text. Headers/footers not included. Images not extracted.
- **Excel**: Only cell values and basic formatting (bold/italic) are converted. Formulas, charts, images, conditional formatting, and data validation are not preserved. Large spreads may produce very wide Markdown tables.
- **All formats**: Password-protected files are not supported.
