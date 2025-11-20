from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def process_images(doc):
    """Обработка изображений и подписей к ним"""
    paragraphs = doc.paragraphs
    for idx, para in enumerate(paragraphs):
        has_image = any(
            run._element.xpath('.//w:drawing') or run._element.xpath('.//w:pict')
            for run in para.runs
        )

        if has_image:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf = para.paragraph_format
            pf.space_before = Pt(6)
            pf.space_after = Pt(0)
            pf.first_line_indent = Cm(0)
            pf.left_indent = Cm(0)
            pf.right_indent = Cm(0)

            # Подпись (следующий параграф)
            if idx + 1 < len(paragraphs):
                next_para = paragraphs[idx + 1]
                if next_para.text.strip():
                    next_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    npf = next_para.paragraph_format
                    npf.space_before = Pt(0)
                    npf.space_after = Pt(6)
                    npf.first_line_indent = Cm(0)
                    for run in next_para.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(14)
                        run.font.color.rgb = RGBColor(0, 0, 0)
                        run.font.bold = False
                        run.font.italic = False
                        rpr = run._element.get_or_add_rPr()
                        hl = rpr.xpath('./w:highlight')
                        for h in hl:
                            rpr.remove(h)