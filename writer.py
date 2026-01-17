from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

def set_font(run, size=12, bold=False, underline=False):
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size)
    run.bold = bold
    run.underline = underline

def save(data, filepath):
    doc = Document()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("СОГЛАСИЕ\nна обработку персональных данных")
    set_font(run, size=14, bold=True)

    p = doc.add_paragraph()
    set_font(p.add_run("Я, "))
    set_font(p.add_run(f" {data['fio']} "), underline=True)
    p.add_run("\n")
    set_font(p.add_run("проживающий(ая) по адресу: "))
    set_font(p.add_run(f" {data['address']} "), underline=True)

    # Остальной юридический текст опущен для краткости на данном этапе
    doc.save(filepath)
