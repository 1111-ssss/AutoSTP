from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import re

def set_paragraph_format(para, font_size=14, first_line_indent=Cm(1.25), space_before=0, space_after=0, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY):
    """Универсальная функция для форматирования параграфа по СТП"""
    for run in para.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(font_size)
        run.font.bold = False
        run.font.italic = False
        run.font.underline = False
        run.font.color.rgb = RGBColor(0, 0, 0)
        rpr = run._element.get_or_add_rPr()
        hl = rpr.xpath('./w:highlight')
        for h in hl:
            rpr.remove(h)
    
    pf = para.paragraph_format
    pf.first_line_indent = first_line_indent
    pf.left_indent = Cm(0)
    pf.right_indent = Cm(0)
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = 1.0
    pf.alignment = alignment

def process_paragraphs(doc):
    """Обработка всех параграфов: заголовки, подразделы, задания и обычный текст"""
    total_paras = len(doc.paragraphs)

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            set_paragraph_format(para, first_line_indent=Cm(0))
            continue

        # Заголовок лабораторной
        if i == 0 and re.match(r'ЛАБОРАТОРНАЯ РАБОТА №\s*\d+', text, re.IGNORECASE):
            for run in para.runs:
                run.font.size = Pt(16)
                run.font.bold = True
                run.font.name = 'Times New Roman'
            pf = para.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.space_before = Pt(0)
            pf.space_after = Pt(12)
            pf.first_line_indent = Cm(0)
            continue

        # Подтемы: "1. ..."
        if re.match(r'^\d+\.\s', text):
            for run in para.runs:
                run.font.bold = True
                run.font.name = 'Times New Roman'
                run.font.size = Pt(14)
            pf = para.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(0)
            pf.space_after = Pt(6)
            continue

        # Задания
        if text.startswith("Задание №"):
            for run in para.runs:
                run.font.bold = True
                run.font.name = 'Times New Roman'
                run.font.size = Pt(14)
            pf = para.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(6)
            pf.space_after = Pt(0)
            continue

        # Обычный текст
        set_paragraph_format(para)