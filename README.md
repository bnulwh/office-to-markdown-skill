# Office to Markdown

[![skills.sh](https://skills.sh/b/bnulwh/office-to-markdown-skill)](https://skills.sh/bnulwh/office-to-markdown-skill)
![Python](https://img.shields.io/badge/python-3.10+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

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

## Install

### Via skills CLI (Recommended)

The [skills CLI](https://github.com/vercel-labs/skills) supports 70+ AI coding agents. Install once, use everywhere:

```bash
# Install to all detected agents
npx skills add bnulwh/office-to-markdown-skill --all

# Or install interactively (choose agents)
npx skills add bnulwh/office-to-markdown-skill
```

### For Specific Agents

```bash
# Claude Code
npx skills add bnulwh/office-to-markdown-skill -a claude-code

# OpenCode
npx skills add bnulwh/office-to-markdown-skill -a opencode

# Qoder / Qoder CN
npx skills add bnulwh/office-to-markdown-skill -a qoder
npx skills add bnulwh/office-to-markdown-skill -a qoder-cn

# OpenClaw
npx skills add bnulwh/office-to-markdown-skill -a openclaw

# Cursor
npx skills add bnulwh/office-to-markdown-skill -a cursor

# Codex
npx skills add bnulwh/office-to-markdown-skill -a codex

# Cline
npx skills add bnulwh/office-to-markdown-skill -a cline

# Gemini CLI
npx skills add bnulwh/office-to-markdown-skill -a gemini-cli

# GitHub Copilot
npx skills add bnulwh/office-to-markdown-skill -a github-copilot

# Windsurf
npx skills add bnulwh/office-to-markdown-skill -a windsurf

# Trae / Trae CN
npx skills add bnulwh/office-to-markdown-skill -a trae
npx skills add bnulwh/office-to-markdown-skill -a trae-cn

# Qwen Code
npx skills add bnulwh/office-to-markdown-skill -a qwen-code
```

### Manual Installation

Clone the repository and copy/symlink the skill directory to your agent's skills folder:

| Agent | Project Path | Global Path |
|-------|--------------|-------------|
| Claude Code | `.claude/skills/` | `~/.claude/skills/` |
| OpenCode | `.agents/skills/` | `~/.config/opencode/skills/` |
| Qoder | `.qoder/skills/` | `~/.qoder/skills/` |
| OpenClaw | `skills/` | `~/.openclaw/skills/` |
| Cursor | `.agents/skills/` | `~/.cursor/skills/` |
| Codex | `.agents/skills/` | `~/.codex/skills/` |
| Cline | `.agents/skills/` | `~/.agents/skills/` |
| Gemini CLI | `.agents/skills/` | `~/.gemini/skills/` |
| GitHub Copilot | `.agents/skills/` | `~/.copilot/skills/` |
| Windsurf | `.windsurf/skills/` | `~/.codeium/windsurf/skills/` |
| Trae | `.trae/skills/` | `~/.trae/skills/` |
| Qwen Code | `.qwen/skills/` | `~/.qwen/skills/` |

```bash
# Example: install to Claude Code globally
git clone https://github.com/bnulwh/office-to-markdown-skill.git
ln -s $(pwd)/office-to-markdown-skill ~/.claude/skills/office-to-markdown

# Example: install to OpenCode project-level
mkdir -p .agents/skills
ln -s /path/to/office-to-markdown-skill .agents/skills/office-to-markdown
```

### Install Dependencies

```bash
pip install python-pptx python-docx openpyxl
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
