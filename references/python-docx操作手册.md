# python-docx 操作手册

本文件记录 python-docx 在公文格式修订中的能力边界、常用操作和陷阱。核验日期为二〇二六年七月二十九日。

## 能力清单

| 操作 | 支持程度 | 说明 |
|---|---|---|
| 段落字体、字号 | 完全支持 | `run.font.name`, `run.font.size` |
| 段落对齐 | 完全支持 | `paragraph_format.alignment` |
| 行距 | 完全支持 | `paragraph_format.line_spacing` |
| 首行缩进 | 完全支持 | `paragraph_format.first_line_indent` |
| 段前段后间距 | 完全支持 | `paragraph_format.space_before/after` |
| 页面大小 | 完全支持 | `section.page_width/height` |
| 页边距 | 完全支持 | `section.top/bottom/left/right_margin` |
| 表格格式 | 完全支持 | 单元格字体、对齐、宽度 |
| 页眉页脚文字 | 完全支持 | `section.header/footer` |
| 首页不同 | 完全支持 | `section.different_first_page_header_footer` |
| 奇偶页不同 | 完全支持 | `section.different_odd_even` |
| 页码域代码 | 需 XML 操作 | 无原生 API，见 `页码与页眉页脚.md` |
| 段落边框（红线） | 需 XML 操作 | 操作 `w:pBdr` 元素 |
| 分节符精细控制 | 部分支持 | 可添加节，但类型控制有限 |
| 字体嵌入 | 不支持 | 只写字体名，不嵌入字体文件 |
| 图片格式修改 | 有限支持 | 可调整大小和位置，不能编辑图片内容 |
| 修订标记（Track Changes） | 不支持 | 不能以修订模式写入 |
| 域代码（Field） | 需 XML 操作 | 页码、日期等域需手动构建 |

## 字体设置（关键陷阱）

python-docx 的 `run.font.name` 只设置 `w:ascii` 和 `w:hAnsi` 属性（控制拉丁字母和阿拉伯数字）。中文字符使用 `w:eastAsia` 属性，必须单独设置。

公文中数字和拉丁字母应使用与中文相同的字体（如正文中的"2024年""OA平台"应使用仿宋_GB2312，不用 Times New Roman）。因此三个属性统一设置为同一字体：

```python
from docx.oxml.ns import qn

def set_font_all(run, font_name, font_size_pt=None, bold=None):
    """统一设置 run 的全部字体（中文、数字、拉丁）"""
    run.font.name = font_name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn('w:eastAsia'), font_name)
    rfonts.set(qn('w:ascii'), font_name)
    rfonts.set(qn('w:hAnsi'), font_name)
    if font_size_pt is not None:
        run.font.size = Pt(font_size_pt)
    if bold is not None:
        run.font.bold = bold
```

常见错误：
- 只设置 `run.font.name`：中文字符不变（缺少 eastAsia）。
- 将 ascii/hAnsi 设为 Times New Roman：数字和拉丁字母字体与中文不一致。

## 段落格式批量设置

```python
from docx.shared import Pt, Mm
from docx.enum.text import WD_LINE_SPACING, WD_ALIGN_PARAGRAPH

def set_body_format(paragraph):
    """将段落设置为公文正文格式"""
    pf = paragraph.paragraph_format
    pf.line_spacing = Pt(28.95)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.first_line_indent = Pt(32)  # 2字符 × 16pt
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    for run in paragraph.runs:
        set_font_all(run, '仿宋_GB2312', 16)
```

## 页面设置

```python
from docx.shared import Mm

def set_page_layout(section):
    """设置 A4 页面和公文页边距"""
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Mm(37)
    section.bottom_margin = Mm(35)
    section.left_margin = Mm(28)
    section.right_margin = Mm(26)
```

## 内容哈希验证

```python
import hashlib
from docx import Document

def content_hash(docx_path):
    """计算文档文字内容的 SHA-256 哈希"""
    doc = Document(docx_path)
    texts = []
    for para in doc.paragraphs:
        texts.append(para.text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                texts.append(cell.text)
    combined = '\n'.join(texts)
    return hashlib.sha256(combined.encode('utf-8')).hexdigest()
```

修订前后调用此函数，哈希必须一致。

## 常见陷阱

| 陷阱 | 说明 | 解决 |
|---|---|---|
| 中文字体不生效 | 只设了 font.name，未设 eastAsia | 用 `rFonts.set(qn('w:eastAsia'), ...)` |
| 行距不生效 | 未设 line_spacing_rule 为 EXACTLY | 同时设 `line_spacing` 和 `line_spacing_rule` |
| 修改了样式而非直接格式 | 通过 style 修改会影响所有使用该样式的段落 | 直接修改 paragraph_format 和 run.font |
| 段落无 runs | 空段落或只含换行符的段落没有 runs | 检查 `len(paragraph.runs) > 0` |
| 继承格式 | run 未显式设置的属性继承自样式 | 修订时需要显式设置所有目标属性 |
| 表格内段落 | 表格单元格内的段落不在 `doc.paragraphs` 中 | 需遍历 `table.rows[].cells[].paragraphs` |
| 节属性 | 多节文档每节有独立的页面设置 | 遍历 `doc.sections` 而非只改第一个 |

## 版本兼容

- 推荐 python-docx >= 0.8.11。
- Python >= 3.8。
- 安装：`pip install python-docx`。
- 不依赖 Microsoft Office 安装（纯 Python 操作 XML）。
- 但渲染验证需要 Office 或 LibreOffice。
