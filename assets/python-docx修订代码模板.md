# python-docx 修订代码模板

常用公文格式修订操作的代码片段集合。所有代码只操作格式属性，不修改文字内容。

## 前置导入

```python
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_LINE_SPACING, WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import hashlib
```

## 字体设置

公文中数字和拉丁字母应使用与中文相同的字体（如正文数字用仿宋_GB2312，不用 Times New Roman）。
因此 `w:ascii`、`w:hAnsi`、`w:eastAsia` 三个属性统一设置为同一字体：

```python
def set_font_all(run, font_name, size_pt=None, bold=None):
    """设置 run 的全部字体（中文、数字、拉丁统一）"""
    run.font.name = font_name  # 设置 w:ascii 和 w:hAnsi
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn('w:eastAsia'), font_name)  # 设置 w:eastAsia
    rfonts.set(qn('w:ascii'), font_name)     # 显式设置 w:ascii
    rfonts.set(qn('w:hAnsi'), font_name)     # 显式设置 w:hAnsi
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bold
```

## 页面设置

```python
def fix_page_layout(doc):
    """设置 A4 页面和公文标准页边距"""
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Mm(37)
        section.bottom_margin = Mm(35)
        section.left_margin = Mm(28)
        section.right_margin = Mm(26)
```

## 正文段落格式

```python
def fix_body_paragraph(paragraph):
    """将段落设置为公文正文格式"""
    pf = paragraph.paragraph_format
    pf.line_spacing = Pt(28.95)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.first_line_indent = Pt(32)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    for run in paragraph.runs:
        set_font_all(run, '仿宋_GB2312', size_pt=16)
```

## 标题格式

```python
def fix_title_paragraph(paragraph):
    """将段落设置为公文标题格式"""
    pf = paragraph.paragraph_format
    pf.line_spacing = Pt(28.95)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.first_line_indent = Pt(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    for run in paragraph.runs:
        set_font_all(run, '方正小标宋简体', size_pt=22)
        run.font.bold = False
```

## 一级标题格式

```python
def fix_heading1(paragraph):
    """一级标题：黑体，三号，不加粗"""
    pf = paragraph.paragraph_format
    pf.line_spacing = Pt(28.95)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.first_line_indent = Pt(32)
    for run in paragraph.runs:
        set_font_all(run, '黑体', size_pt=16)
        run.font.bold = False
```

## 二级标题格式（含接排正文）

二级标题常与正文接排在同一段落中（如"（一）事项上报。对网格员发现……"）。
原始文档中通常用 bold 区分标题部分和正文部分。修订时按 run 的原始 bold 状态分别设置字体：

```python
def fix_heading2(paragraph):
    """二级标题：标题部分楷体_GB2312，接排正文部分仿宋_GB2312"""
    pf = paragraph.paragraph_format
    pf.line_spacing = Pt(28.95)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.first_line_indent = Pt(32)
    for run in paragraph.runs:
        if run.font.bold:
            # 标题部分（原始文档中加粗的 run）
            set_font_all(run, '楷体_GB2312', size_pt=16)
            run.font.bold = False  # 楷体标题不加粗
        else:
            # 接排正文部分
            set_font_all(run, '仿宋_GB2312', size_pt=16)
            run.font.bold = False
```

如果二级标题独占一段（无接排正文），所有 run 均设为楷体_GB2312。

## 三级标题格式

```python
def fix_heading3(paragraph):
    """三级标题：仿宋_GB2312，三号，加粗"""
    pf = paragraph.paragraph_format
    pf.line_spacing = Pt(28.95)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.first_line_indent = Pt(32)
    for run in paragraph.runs:
        set_font_all(run, '仿宋_GB2312', size_pt=16)
        run.font.bold = True
```

## 内容哈希验证

```python
def content_hash(docx_path):
    """计算文档文字内容的 SHA-256 哈希（修订前后必须一致）"""
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

## 安全修订流程

修订采用"读取原文件→另存为新文件"模式，原文件始终不被修改，无需备份：

```python
def safe_revise(input_path, output_path, revise_func):
    """安全修订：读取→修订→另存→验证（原文件不被修改）"""
    # 修订前哈希（从原文件读取）
    hash_before = content_hash(input_path)

    # 执行修订（从原文件读取，保存到新文件）
    doc = Document(input_path)
    revise_func(doc)
    doc.save(output_path)

    # 修订后哈希（从新文件读取）
    hash_after = content_hash(output_path)

    # 验证内容不变
    assert hash_before == hash_after, f"内容哈希不一致！修订前={hash_before}，修订后={hash_after}"

    # 验证段落数
    doc_before = Document(input_path)
    doc_after = Document(output_path)
    assert len(doc_before.paragraphs) == len(doc_after.paragraphs), "段落数不一致"

    return True
```

## 三级标题 vs 编号列表区分

"1." 开头的段落可能是三级标题（如"1.事项接收。及时查看……"），也可能是编号列表项（如"1.市委社会工作部一处处长；"）。
区分方法：检查原始 run 是否加粗。原始加粗 → 三级标题；原始不加粗 → 编号列表项（按正文处理）。

```python
import re
h3_pat = re.compile(r'^\d+[\.\、]')

def fix_h3_or_list(paragraph):
    """三级标题或编号列表：按原始 bold 状态区分"""
    pf = paragraph.paragraph_format
    pf.line_spacing = Pt(28.95)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.first_line_indent = Pt(32)
    has_bold = any(r.font.bold for r in paragraph.runs if r.text.strip())
    if has_bold:
        for run in paragraph.runs:
            set_font_all(run, '仿宋_GB2312', size_pt=16, bold=bool(run.font.bold))
    else:
        for run in paragraph.runs:
            set_font_all(run, '仿宋_GB2312', size_pt=16, bold=False)
```

## 页脚重建（页码修订）

原文档页码可能在浮动文本框中（位置偏左或偏右）。修订时清除浮动对象，重建为居中段落：

```python
def rebuild_footer_centered(section):
    """重建页脚：居中、宋体四号、— X — 格式"""
    footer = section.footer
    footer.is_linked_to_previous = False

    # 清空现有页脚内容（含浮动文本框）
    for p in footer.paragraphs:
        p._element.getparent().remove(p._element)

    # 新建居中段落
    p = footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)

    # "— " 前缀
    r1 = p.add_run('\u2014 ')
    set_font_all(r1, '宋体', size_pt=14)

    # PAGE 域代码
    r2 = p.add_run()
    set_font_all(r2, '宋体', size_pt=14)
    fldBegin = OxmlElement('w:fldChar')
    fldBegin.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = ' PAGE '
    fldEnd = OxmlElement('w:fldChar')
    fldEnd.set(qn('w:fldCharType'), 'end')
    r2._element.append(fldBegin)
    r2._element.append(instrText)
    r2._element.append(fldEnd)

    # " —" 后缀
    r3 = p.add_run(' \u2014')
    set_font_all(r3, '宋体', size_pt=14)
```

## 表格单元格格式修订

表格内段落不在 `doc.paragraphs` 中，需单独遍历。修订原则：只统一字体为仿宋_GB2312，不改字号（避免撑破单元格），不改表格结构。

```python
def fix_table_fonts(doc):
    """统一表格单元格字体为仿宋_GB2312，不改字号和结构"""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        set_font_all(run, '仿宋_GB2312')
                        run.font.color.rgb = RGBColor(0, 0, 0)
```

注意：表格段前间距无法直接设置在表格上（OOXML 表格无段前属性）。如需表格上方留白，
可设置表格上方相邻段落的 `space_after`。该操作天然幂等，重复执行不会累加。

```python
def fix_table_spacing(doc, before_pt=Pt(14)):
    """设置表格上方相邻段落的段后间距（默认 0.5 行 ≈ 14pt）"""
    body = doc.element.body
    for table in doc.tables:
        tbl_elem = table._tbl
        prev = tbl_elem.getprevious()
        if prev is not None and prev.tag.endswith('}p'):
            from docx.text.paragraph import Paragraph
            para = Paragraph(prev, doc)
            para.paragraph_format.space_after = before_pt
```

## 颜色归一化

装饰性字体文档（华文彩云、方正舒体等）常伴随彩色文字。修订字体后颜色可能残留，
需统一归一为黑色。此操作只改颜色属性，不改文字内容。

```python
def normalize_colors(doc):
    """将全文文字颜色统一为黑色（含表格）"""
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.color.rgb = RGBColor(0, 0, 0)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.color.rgb = RGBColor(0, 0, 0)
```

## 段前段后间距（可配置）

GB/T 9704-2012 未明确规定段前段后间距，按本机关制度处理。
以下提供国企常见实践作为可选配置，默认不启用（全部为 0）：

```python
# 可选：一级标题段前 1 行（28.95pt）
def fix_heading1_with_spacing(paragraph):
    fix_heading1(paragraph)
    paragraph.paragraph_format.space_before = Pt(28.95)

# 可选：二级标题段前 0.5 行（14.475pt）
def fix_heading2_with_spacing(paragraph):
    fix_heading2(paragraph)
    paragraph.paragraph_format.space_before = Pt(14.475)

# 可选：大标题段后 1.5 行（43.425pt）
def fix_title_with_spacing(paragraph):
    fix_title_paragraph(paragraph)
    paragraph.paragraph_format.space_after = Pt(43.425)
```

启用前须确认本机关制度是否要求段前段后间距。若机关无此要求，保持默认全 0。

## 注意事项

- 所有 `set_font_all` 调用统一设置 ascii/hAnsi/eastAsia，确保数字和拉丁字母也使用公文标准字体。
- 行距必须同时设置 `line_spacing` 和 `line_spacing_rule = WD_LINE_SPACING.EXACTLY`。
- 修订代码中禁止出现 `paragraph.text = ...` 或 `run.text = ...`。
- 修订代码中禁止调用 `paragraph.clear()` 或删除段落。
- 图片段落（含 `inline_shapes` 的段落）跳过格式修订。
- 表格修订只改字体和颜色，不改字号、列宽、行高和单元格合并。
- 颜色归一化会将所有彩色文字变为黑色，修订前须确认用户无需保留的彩色强调。
