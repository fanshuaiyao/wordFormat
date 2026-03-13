from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from lxml import etree
from src.core.format_spec import DocumentFormat, SectionFormat


ALIGNMENT_MAP = {
    "CENTER": WD_PARAGRAPH_ALIGNMENT.CENTER,
    "LEFT": WD_PARAGRAPH_ALIGNMENT.LEFT,
    "RIGHT": WD_PARAGRAPH_ALIGNMENT.RIGHT,
    "JUSTIFY": WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
}

# DocumentFormat 字段 → Word 样式名映射
STYLE_NAME_MAP = {
    "title": "Title",
    "abstract": "Abstract",
    "keywords": "Keywords",
    "heading1": "Heading 1",
    "heading2": "Heading 2",
    "body": "Normal",
    "references": "References",
}


class WordFormatter:
    def __init__(self, document, config_manager):
        self.document = document
        self.config_manager = config_manager
        self.format_spec: DocumentFormat = None
        self._styles = {}  # 缓存已设置的样式对象

    def set_format_spec(self, format_spec: DocumentFormat):
        """设置格式规范"""
        self.format_spec = format_spec

    def format(self):
        """应用格式到文档"""
        if not self.format_spec:
            print("未设置格式规范，跳过格式化")
            return False

        try:
            doc = self.document
            print("开始应用格式...")

            # 1. 修改 Word 样式定义
            self._setup_styles()

            # 2. 应用页边距
            self._apply_page_margins()

            # 3. 给段落赋样式
            # 标题
            if doc.title:
                self._assign_paragraph_style(doc.title, "title")

            # 摘要
            if doc.abstract:
                self._assign_paragraph_style(doc.abstract, "abstract")

            # 关键词
            if doc.keywords:
                self._assign_paragraph_style(doc.keywords, "keywords")

            # 各章节
            for section_name, paragraphs in doc.sections.items():
                # 章节标题段落
                heading_para = self._find_heading_paragraph(section_name)
                if heading_para:
                    if self._is_main_heading(section_name):
                        self._assign_paragraph_style(heading_para, "heading1")
                    else:
                        self._assign_paragraph_style(heading_para, "heading2")

                # 判断是否为参考文献章节
                is_references = '参考文献' in section_name or 'references' in section_name.lower()

                # 章节内段落
                for para in paragraphs:
                    if is_references:
                        self._assign_paragraph_style(para, "references")
                    else:
                        self._assign_paragraph_style(para, "body")

            print("格式应用完成")
            return True

        except Exception as e:
            print(f"格式化失败: {str(e)}")
            raise

    # ── 样式定义 ──

    def _setup_styles(self):
        """修改 Word 样式定义，使其与 format_spec 一致"""
        spec = self.format_spec
        for field_name, style_name in STYLE_NAME_MAP.items():
            section_spec = getattr(spec, field_name, None)
            if not section_spec:
                continue
            style = self._get_or_create_style(style_name)
            self._apply_style_format(style, section_spec)
            self._styles[field_name] = style

    def _get_or_create_style(self, style_name):
        """获取已有样式，不存在则创建"""
        doc = self.document.doc
        try:
            return doc.styles[style_name]
        except KeyError:
            style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
            # 自定义样式基于 Normal
            try:
                style.base_style = doc.styles['Normal']
            except KeyError:
                pass
            return style

    def _apply_style_format(self, style, spec: SectionFormat):
        """将 SectionFormat 应用到 Word 样式定义上"""
        # 字体
        style.font.name = spec.font_name
        style.font.size = Pt(spec.font_size)
        style.font.bold = spec.bold
        style.font.italic = spec.italic

        # 东亚字体（中文字体）和西文字体（ASCII）
        rPr = style.element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = etree.SubElement(rPr, qn('w:rFonts'))

        # 处理可能没有 cn 属性的向后兼容，防止出错
        font_name_cn = getattr(spec, "font_name_cn", spec.font_name)

        # 对应：西方文本字体 / 中文字体
        rFonts.set(qn('w:ascii'), spec.font_name)
        rFonts.set(qn('w:hAnsi'), spec.font_name)
        rFonts.set(qn('w:eastAsia'), font_name_cn)

        # 对齐方式
        style.paragraph_format.alignment = ALIGNMENT_MAP.get(
            spec.alignment, WD_PARAGRAPH_ALIGNMENT.LEFT
        )

        # 行距
        style.paragraph_format.line_spacing = spec.line_spacing

        # 段前段后间距
        style.paragraph_format.space_before = Pt(spec.space_before)
        style.paragraph_format.space_after = Pt(spec.space_after)

        # 首行缩进 / 悬挂缩进 (以中文字符数计算: 缩进值 * 字体大小)
        if spec.first_line_indent > 0:
            indent_pt = spec.first_line_indent * spec.font_size
            style.paragraph_format.first_line_indent = Pt(indent_pt)
            style.paragraph_format.left_indent = None
        elif spec.first_line_indent < 0:
            indent_pt = abs(spec.first_line_indent) * spec.font_size
            style.paragraph_format.first_line_indent = Pt(-indent_pt)
            style.paragraph_format.left_indent = Pt(indent_pt)
        else:
            style.paragraph_format.first_line_indent = None
            style.paragraph_format.left_indent = None

    # ── 段落赋样式 ──

    def _assign_paragraph_style(self, paragraph, format_field_name):
        """将段落绑定到对应的 Word 样式，并清除内联格式"""
        style = self._styles.get(format_field_name)
        if not style:
            return

        # 赋样式
        paragraph.style = style

        # 清除段落级内联格式覆盖，让样式定义生效
        pf = paragraph.paragraph_format
        pf.alignment = None
        pf.line_spacing = None
        pf.space_before = None
        pf.space_after = None
        pf.first_line_indent = None
        pf.left_indent = None

        # 清除 run 级内联格式覆盖
        for run in paragraph.runs:
            run.font.name = None
            run.font.size = None
            run.font.bold = None
            run.font.italic = None
            # 清除东亚字体的内联覆盖
            rPr = run._element.find(qn('w:rPr'))
            if rPr is not None:
                rFonts = rPr.find(qn('w:rFonts'))
                if rFonts is not None:
                    for attr in [qn('w:eastAsia'), qn('w:ascii'), qn('w:hAnsi')]:
                        if attr in rFonts.attrib:
                            del rFonts.attrib[attr]
                    # 如果 rFonts 没有属性了，删除整个元素
                    if not rFonts.attrib:
                        rPr.remove(rFonts)

    # ── 辅助方法 ──

    def _find_heading_paragraph(self, section_name):
        """在文档段落中查找与章节名匹配的标题段落"""
        for para in self.document.doc.paragraphs:
            if para.text.strip() == section_name:
                return para
        return None

    def _is_main_heading(self, text):
        """判断是否为一级标题"""
        text = text.strip()
        if any(text.startswith(f"{i}.") for i in range(1, 10)):
            return True
        chinese_numbers = ['一', '二', '三', '四', '五', '六', '七', '八', '九']
        if any(text.startswith(f"{num}、") for num in chinese_numbers):
            return True
        main_keywords = ['引言', '介绍', '研究背景', '研究方法', '实验', '结果', '讨论', '结论', '参考文献']
        return any(text.startswith(kw) for kw in main_keywords)

    def _apply_page_margins(self):
        """应用页边距"""
        margins = self.format_spec.page_margin
        if not margins:
            return
        section = self.document.doc.sections[0]
        section.top_margin = Inches(margins.get("top", 1.0))
        section.bottom_margin = Inches(margins.get("bottom", 1.0))
        section.left_margin = Inches(margins.get("left", 1.25))
        section.right_margin = Inches(margins.get("right", 1.25))
