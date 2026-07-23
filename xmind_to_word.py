"""
XMind 转 Word 工具
解析 .xmind 文件（ZIP 格式），按原有父子节点层级关系生成 Word 文档，
并应用中文排版习惯的字体、段落格式。

用法:
    python xmind_to_word.py <输入.xmind> <输出.docx>
"""

import sys
import json
import re
import zipfile
import tempfile
import shutil
import os

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

MAX_HEADING = 9

# 中文排版字号（磅）
SIZE_TITLE = 22   # 二号（封面标题）
SIZE_H1 = 22       # 二号
SIZE_H2 = 15       # 小三
SIZE_H3 = 14       # 四号
SIZE_BODY = 12     # 小四

FONT_HEADING = '黑体'
FONT_BODY = '宋体'

HEADING_STYLES = {'Title', 'Heading 1', 'Heading 2', 'Heading 3'}


def extract_xmind(xmind_path, extract_dir):
    """解压 .xmind 文件到指定目录"""
    with zipfile.ZipFile(xmind_path, 'r') as zf:
        zf.extractall(extract_dir)


# 全角空格(U+3000)、不间断空格(U+00A0)、零宽空格(U+200B)、
# 零宽连字/断字(U+200C/200D)、BOM(U+FEFF) 等不可见排版污染字符
_INVISIBLE_CHARS = re.compile(
    "[\u3000\u00a0\u200b\u200c\u200d\ufeff]"
)
_CONTROL_CHARS = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")


def clean_text(text):
    """清除全角空格 / 不间断空格 / 零宽字符 / 控制字符等排版污染"""
    text = _INVISIBLE_CHARS.sub(' ', text)
    text = _CONTROL_CHARS.sub('', text)
    return re.sub(r' {2,}', ' ', text).strip()


def parse_topic(topic, level=0):
    """递归解析思维导图节点，保留父子层级关系"""
    result = []

    title = clean_text(topic.get('title', ''))
    if title:
        result.append({'level': level, 'text': title})

    image = topic.get('image')
    if image:
        src = image.get('src', '')
        if src.startswith('xap:'):
            src = src[4:]
        result.append({'level': level, 'image': src})

    children = topic.get('children', {})
    for child in children.get('attached', []):
        result.extend(parse_topic(child, level + 1))

    return result


def set_run_font(run, font_name, size_pt, bold=False):
    """设置字体（中英文字体、字号、颜色）"""
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = False
    run.font.color.rgb = RGBColor(0, 0, 0)

    rpr = run._element.get_or_add_rPr()
    rFonts = rpr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rpr.makeelement(qn('w:rFonts'), {})
        rpr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)


def set_first_line_indent_chars(paragraph, chars=2):
    """设置首行缩进（按字符数，中文排版惯用单位）"""
    pPr = paragraph._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = pPr.makeelement(qn('w:ind'), {})
        pPr.append(ind)
    ind.set(qn('w:firstLineChars'), str(chars * 100))
    if ind.get(qn('w:firstLine')) is not None:
        del ind.attrib[qn('w:firstLine')]


def reset_char_spacing(run):
    """将字符间距显式重置为标准（0），避免继承样式或外部来源写入的加宽值"""
    rpr = run._element.get_or_add_rPr()
    spacing = rpr.find(qn('w:spacing'))
    if spacing is None:
        spacing = rpr.makeelement(qn('w:spacing'), {})
        rpr.append(spacing)
    spacing.set(qn('w:val'), '0')


def create_word_doc(content_list, output_path, resource_dir, doc_title='文档'):
    """根据解析结果生成 Word 文档，并应用中文排版格式"""
    doc = Document()

    title_p = doc.add_heading(doc_title, 0)
    title_p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    for item in content_list:
        level = item['level']

        if 'image' in item:
            img_path = os.path.join(resource_dir, item['image'])
            try:
                doc.add_picture(img_path, width=Inches(5))
            except Exception:
                doc.add_paragraph(f'[图片加载失败: {item["image"]}]')
            continue

        text = item['text']
        heading_level = level + 1

        if heading_level <= MAX_HEADING:
            p = doc.add_heading(text, level=heading_level)
        else:
            indent = '    ' * (heading_level - MAX_HEADING - 1)
            p = doc.add_paragraph(style='List Bullet')
            p.add_run(indent + text)

        apply_style(p)

    doc.save(output_path)


def apply_style(paragraph):
    """对单个段落应用字体、颜色、对齐、间距、缩进等中文排版格式"""
    style_name = paragraph.style.name

    if style_name == 'Heading 1':
        font, size, bold = FONT_HEADING, SIZE_H1, True
    elif style_name == 'Heading 2':
        font, size, bold = FONT_HEADING, SIZE_H2, True
    elif style_name == 'Heading 3':
        font, size, bold = FONT_HEADING, SIZE_H3, True
    else:
        font, size, bold = FONT_BODY, SIZE_BODY, False

    for run in paragraph.runs:
        set_run_font(run, font, size, bold=bold)
        reset_char_spacing(run)

    if style_name not in HEADING_STYLES:
        pf = paragraph.paragraph_format
        # 短语类节点内容两端对齐会被强行拉伸字间距撑满整行，改为左对齐
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        pf.left_indent = 0
        set_first_line_indent_chars(paragraph, 2)


def convert(xmind_path, output_path):
    """主转换流程：解压 -> 解析 -> 生成 Word"""
    tmp_dir = tempfile.mkdtemp(prefix='xmind_')
    try:
        extract_xmind(xmind_path, tmp_dir)

        with open(os.path.join(tmp_dir, 'content.json'), 'r', encoding='utf-8') as f:
            data = json.load(f)

        all_items = []
        for sheet in data:
            root_topic = sheet.get('rootTopic', {})
            all_items.extend(parse_topic(root_topic, 0))

        doc_title = os.path.splitext(os.path.basename(xmind_path))[0]
        create_word_doc(all_items, output_path, tmp_dir, doc_title=doc_title)

        print(f'共解析 {len(all_items)} 个节点')
        print(f'Word 文档已保存到: {output_path}')
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def main():
    if len(sys.argv) != 3:
        print('用法: python xmind_to_word.py <输入.xmind> <输出.docx>')
        sys.exit(1)

    convert(sys.argv[1], sys.argv[2])


if __name__ == '__main__':
    main()
