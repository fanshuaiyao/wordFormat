from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from src.core.format_spec import DocumentFormat, SectionFormat


ALIGNMENT_MAP = {
    "CENTER": WD_PARAGRAPH_ALIGNMENT.CENTER,
    "LEFT": WD_PARAGRAPH_ALIGNMENT.LEFT,
    "RIGHT": WD_PARAGRAPH_ALIGNMENT.RIGHT,
    "JUSTIFY": WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
}


class WordFormatter:
    def __init__(self, document, config_manager):
        self.document = document
        self.config_manager = config_manager
        self.format_spec: DocumentFormat = None

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

            # 应用页边距
            self._apply_page_margins()

            # 格式化标题
            if doc.title:
                self._apply_section_format(doc.title, self.format_spec.title)

            # 格式化摘要
            if doc.abstract:
                self._apply_section_format(doc.abstract, self.format_spec.abstract)

            # 格式化关键词
            if doc.keywords:
                self._apply_section_format(doc.keywords, self.format_spec.keywords)

            # 格式化各章节（标题 + 正文段落）
            for section_name, paragraphs in doc.sections.items():
                # 查找并格式化章节标题段落
                heading_para = self._find_heading_paragraph(section_name)
                if heading_para:
                    if self._is_main_heading(section_name):
                        self._apply_section_format(heading_para, self.format_spec.heading1)
                    else:
                        self._apply_section_format(heading_para, self.format_spec.heading2)

                # 判断是否为参考文献章节
                is_references = '参考文献' in section_name or 'references' in section_name.lower()

                # 格式化章节内的段落
                for para in paragraphs:
                    if is_references:
                        self._apply_section_format(para, self.format_spec.references)
                    else:
                        self._apply_section_format(para, self.format_spec.body)

            print("格式应用完成")
            return True

        except Exception as e:
            print(f"格式化失败: {str(e)}")
            raise

    def _find_heading_paragraph(self, section_name):
        """在文档段落中查找与章节名匹配的标题段落"""
        for para in self.document.doc.paragraphs:
            if para.text.strip() == section_name:
                return para
        return None

    def _is_main_heading(self, text):
        """判断是否为一级标题"""
        text = text.strip()
        # 阿拉伯数字编号（如 "1." 或 "1. 引言"）
        if any(text.startswith(f"{i}.") for i in range(1, 10)):
            return True
        # 中文数字编号（如 "一、引言"）
        chinese_numbers = ['一', '二', '三', '四', '五', '六', '七', '八', '九']
        if any(text.startswith(f"{num}、") for num in chinese_numbers):
            return True
        # 一级标题关键词
        main_keywords = ['引言', '介绍', '研究背景', '研究方法', '实验', '结果', '讨论', '结论', '参考文献']
        return any(text.startswith(kw) for kw in main_keywords)

    def _apply_section_format(self, paragraph, spec: SectionFormat):
        """统一的段落格式化方法"""
        if not spec:
            return

        # 字体格式
        for run in paragraph.runs:
            run.font.name = spec.font_name
            run.font.size = Pt(spec.font_size)
            run.font.bold = spec.bold
            run.font.italic = spec.italic
            # 设置中文字体（东亚字体）
            rPr = run._element.get_or_add_rPr()
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                from lxml import etree
                rFonts = etree.SubElement(rPr, qn('w:rFonts'))
            rFonts.set(qn('w:eastAsia'), spec.font_name)

        # 对齐方式
        paragraph.alignment = ALIGNMENT_MAP.get(spec.alignment, WD_PARAGRAPH_ALIGNMENT.LEFT)

        # 段落格式
        pf = paragraph.paragraph_format
        pf.line_spacing = spec.line_spacing
        pf.space_before = Pt(spec.space_before)
        pf.space_after = Pt(spec.space_after)

        # 首行缩进 / 悬挂缩进
        if spec.first_line_indent > 0:
            pf.first_line_indent = Pt(spec.first_line_indent)
            pf.left_indent = None
        elif spec.first_line_indent < 0:
            # 负值表示悬挂缩进：首行回退，左缩进补偿
            pf.first_line_indent = Pt(spec.first_line_indent)
            pf.left_indent = Pt(abs(spec.first_line_indent))
        else:
            pf.first_line_indent = None
            pf.left_indent = None

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
