# Office 转 Markdown

[![skills.sh](https://skills.sh/b/bnulwh/office-to-markdown-skill)](https://skills.sh/bnulwh/office-to-markdown-skill)
![Python](https://img.shields.io/badge/python-3.10+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

将 Office 文件（.pptx、.docx、.xlsx）转换为结构清晰的 Markdown。

## 功能特性

- **PowerPoint**（.pptx）：幻灯片标题作为 Markdown 标题，表格支持合并单元格（rowspan/colspan），多文本框按视觉位置排序（从上到下、从左到右），行内格式（加粗、斜体、下划线、删除线、代码），项目符号列表，演讲者备注，自定义模板的启发式标题检测
- **Word**（.docx）：标题样式识别，中文编号检测（附件、一、二、三、（一）（二）），表格支持合并单元格，有序/无序列表，行内格式，超链接
- **Excel**（.xlsx）：多工作表转换，合并单元格支持，基础加粗/斜体格式

## 安装

```bash
pip install python-pptx python-docx openpyxl
```

## 使用方法

### 统一入口

根据文件扩展名自动识别格式：

```bash
python scripts/office2md.py input.pptx -o output.md
python scripts/office2md.py input.docx -o output.md
python scripts/office2md.py input.xlsx -o output.md
```

### 单独使用各转换器

```bash
python scripts/pptx2md.py 演示文稿.pptx -o output.md
python scripts/docx2md.py 文档.docx -o output.md
python scripts/xlsx2md.py 表格.xlsx -o output.md
```

### Python 模块调用

```python
from scripts.office2md import convert

# 自动识别格式
convert("input.pptx", "output.md")
convert("input.docx", "output.md")
convert("input.xlsx", "output.md")
```

## 转换效果示例

### PowerPoint

每张幻灯片生成一个 `##` 标题段落，幻灯片之间用 `---` 分隔：

```markdown
## 项目概述

- 背景分析
- 核心目标

<table>
  <tr>
    <th>阶段</th>
    <th>时间</th>
  </tr>
  <tr>
    <td>第一阶段</td>
    <td>2026年Q1</td>
  </tr>
</table>

---

## 后续步骤

1. 审核确认
2. 启动实施
```

### Word

标题通过 Word 样式、中文编号规则或字号启发式自动检测：

```markdown
# 文档标题

## 一、项目背景

段落文本，支持 **加粗** 和 *斜体* 格式。

## 二、实施方案

<table>
  <tbody>
    <tr>
      <th>任务</th>
      <th>负责人</th>
    </tr>
    <tr>
      <td>方案设计</td>
      <td>A组</td>
    </tr>
  </tbody>
</table>
```

### Excel

每个工作表生成一个 `##` 标题段落和 HTML 表格：

```markdown
## 销售数据

<table>
  <tr>
    <th>区域</th>
    <th>一季度</th>
    <th>二季度</th>
  </tr>
  <tr>
    <td>华北</td>
    <td><strong>150</strong></td>
    <td>180</td>
  </tr>
</table>

---

## 汇总

<table>
  ...
</table>
```

## 安装

### 通过 skills CLI 安装（推荐）

[skills CLI](https://github.com/vercel-labs/skills) 支持 70+ AI 编程助手。一次安装，处处可用：

```bash
# 安装到所有检测到的 agent
npx skills add bnulwh/office-to-markdown-skill --all

# 或交互式安装（选择 agent）
npx skills add bnulwh/office-to-markdown-skill
```

### 指定 Agent 安装

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

# Qwen Code / 通义灵码
npx skills add bnulwh/office-to-markdown-skill -a qwen-code
```

### 手动安装

克隆仓库并将技能目录复制或软链接到你的 agent 技能文件夹：

| Agent | 项目级路径 | 全局路径 |
|-------|-----------|---------|
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
# 示例：全局安装到 Claude Code
git clone https://github.com/bnulwh/office-to-markdown-skill.git
ln -s $(pwd)/office-to-markdown-skill ~/.claude/skills/office-to-markdown

# 示例：项目级安装到 OpenCode
mkdir -p .agents/skills
ln -s /path/to/office-to-markdown-skill .agents/skills/office-to-markdown
```

### 安装依赖

```bash
pip install python-pptx python-docx openpyxl
```

## 环境要求

- Python 3.10+
- `python-pptx` — PowerPoint 转换
- `python-docx` — Word 转换
- `openpyxl` — Excel 转换

## 已知限制

- 不提取图片、图表和 SmartArt
- 不支持加密文件
- Excel 公式不执行计算（仅读取缓存值）
- 复杂嵌套表格可能被扁平化处理

## 许可证

MIT
