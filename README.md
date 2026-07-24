# Office to Markdown

[中文文档](README_CN.md)

Convert Office files (.pptx, .docx, .xlsx) to well-structured Markdown.

## Features

- **PowerPoint** (.pptx): Slide titles as headings, tables with merged cells, multiple text boxes sorted by visual position, inline formatting, bullet lists, speaker notes, heuristic title detection for custom templates
- **Word** (.docx): Heading styles with Chinese numbering detection, tables with merged cells, ordered/unordered lists, inline formatting, hyperlinks
- **Excel** (.xlsx): Multiple sheets, merged cells with rowspan/colspan, basic bold/italic formatting

## Installation

```bash
pip install python-pptx python-docx openpyxl
```

## Usage

### Unified Entry Point

Auto-detects file type by extension:

```bash
python scripts/office2md.py input.pptx -o output.md
python scripts/office2md.py input.docx -o output.md
python scripts/office2md.py input.xlsx -o output.md
```

### Individual Converters

```bash
python scripts/pptx2md.py presentation.pptx -o output.md
python scripts/docx2md.py document.docx -o output.md
python scripts/xlsx2md.py spreadsheet.xlsx -o output.md
```

### Python Module

```python
from scripts.office2md import convert

# Auto-detect format
convert("input.pptx", "output.md")
convert("input.docx", "output.md")
convert("input.xlsx", "output.md")
```

## Output Examples

### PowerPoint

Each slide becomes a `##` heading section, separated by `---`:

```markdown
## Project Overview

- Background analysis
- Key objectives

<table>
  <tr>
    <th>Phase</th>
    <th>Timeline</th>
  </tr>
  <tr>
    <td>Phase 1</td>
    <td>Q1 2026</td>
  </tr>
</table>

---

## Next Steps

1. Review and approve
2. Begin implementation
```

### Word

Headings are detected from Word styles, Chinese numbering, or font heuristics:

```markdown
# Document Title

## 一、Project Background

Some paragraph text with **bold** and *italic* formatting.

## 二、Implementation Plan

<table>
  <tbody>
    <tr>
      <th>Task</th>
      <th>Owner</th>
    </tr>
    <tr>
      <td>Design</td>
      <td>Team A</td>
    </tr>
  </tbody>
</table>
```

### Excel

Each worksheet becomes a `##` section with an HTML table:

```markdown
## Sales Data

<table>
  <tr>
    <th>Region</th>
    <th>Q1</th>
    <th>Q2</th>
  </tr>
  <tr>
    <td>North</td>
    <td><strong>150</strong></td>
    <td>180</td>
  </tr>
</table>

---

## Summary

<table>
  ...
</table>
```

## Requirements

- Python 3.10+
- `python-pptx` — for PowerPoint conversion
- `python-docx` — for Word conversion
- `openpyxl` — for Excel conversion

## Limitations

- Images, charts, and SmartArt are not extracted
- Password-protected files are not supported
- Excel formulas are not evaluated (only cached values)
- Complex nested tables may be flattened

## License

MIT
