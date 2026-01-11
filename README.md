# docx-format

使用 python-docx 精确读取、分析、修改 Word 文档（.docx）格式的 Claude Code Skill。

## 功能特性

- **格式分析**：分析 Word 文档的字体、字号、缩进、行距等格式
- **格式规范化**：批量修改文档格式，统一排版规范
- **Markdown 转 DOCX**：直接读取 Markdown 并按格式要求写入 DOCX，无需 pandoc
- **中英文混排**：正确处理中英文字体设置
- **预置标准**：内置中国公文格式（GB/T 9704-2012）和学术论文格式

## 安装

```bash
# 使用 Claude Code 安装
claude code skill install git@github.com:ninestep/docx-format-skill.git

# 或使用 HTTPS
claude code skill install https://github.com/ninestep/docx-format-skill.git
```

## 使用方法

### Markdown 转 DOCX

```bash
uv run --with python-docx python3 .claude/skills/docx-format/scripts/md_to_docx.py input.md output.docx
```

**支持的 Markdown 语法**：
- 标题（# ## ### ####）
- 段落
- 列表（有序、无序、任务列表）
- 表格
- 粗体（**text**）
- 代码块（自动跳过）

**格式应用**：
- 标题：黑体加粗
- 正文：宋体小四、1.5倍行距、首行缩进2字符
- 列表：自动项目符号/编号
- 表格：实线黑色0.5磅边框

### 格式分析

```bash
uv run --with python-docx python3 .claude/skills/docx-format/scripts/analyze.py input.docx
```

### 应用公文格式

```bash
uv run --with python-docx python3 .claude/skills/docx-format/scripts/format_official.py input.docx output.docx
```

### 应用学术论文格式

```bash
uv run --with python-docx python3 .claude/skills/docx-format/scripts/format_academic.py input.docx output.docx
```

## 格式标准

### 公文格式（GB/T 9704-2012）

| 元素 | 字体 | 字号 | 行距 |
|------|------|------|------|
| 文档标题 | 方正小标宋简体 | 二号(22pt) | 1.5倍 |
| 一级标题 | 黑体 | 三号(16pt) | 1.5倍 |
| 正文 | 仿宋_GB2312 | 三号(16pt) | 固定28pt |

### 学术论文格式

| 元素 | 字体 | 字号 | 行距 |
|------|------|------|------|
| 论文标题 | 黑体 | 小二(18pt) | 1.5倍 |
| 一级标题 | 黑体 | 小三(15pt) | 1.5倍 |
| 正文 | 宋体 | 小四(12pt) | 1.5倍 |

## 文档结构

```
docx-format/
├── SKILL.md          # Skill 主文档
├── EXAMPLES.md       # 详细示例代码
├── STANDARDS.md      # 格式标准参考
└── scripts/
    ├── md_to_docx.py        # Markdown 转 DOCX
    ├── analyze.py           # 格式分析
    ├── format_official.py   # 公文格式化
    └── format_academic.py   # 学术格式化
```

## 依赖

- python-docx

使用 `uv run --with python-docx` 自动安装依赖。

## 注意事项

1. **只支持 .docx**：不支持旧版 .doc 格式
2. **表格内容**：学术论文、报告等正文常在表格内，脚本会自动处理
3. **备份原文件**：修改前建议备份原文件
4. **中文字体**：使用 `qn('w:eastAsia')` 单独设置中文字体

## 示例

### 基础用法

```python
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

doc = Document('input.docx')

# 修改正文格式
for para in doc.paragraphs:
    if len(para.text.strip()) > 30:
        para.paragraph_format.first_line_indent = Pt(24)
        para.paragraph_format.line_spacing = 1.5
        for run in para.runs:
            run.font.name = 'Times New Roman'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            run.font.size = Pt(12)

doc.save('output.docx')
```

## 许可证

MIT

## 贡献

欢迎提交 Issue 和 Pull Request。
