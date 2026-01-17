from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import datetime

MONTHS = {
    1: 'января', 2: 'февраля', 3: 'марта', 4: 'апреля',
    5: 'мая', 6: 'июня', 7: 'июля', 8: 'августа',
    9: 'сентября', 10: 'октября', 11: 'ноября', 12: 'декабря'
}

def format_russian_date(date_str):
    if not date_str: return "«__» _________ 20__ г."
    try:
        dt = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        return f"«{dt.day}» {MONTHS[dt.month]} {dt.year} г."
    except ValueError:
        return date_str

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

    p = doc.add_paragraph()
    set_font(p.add_run(f"Date: {format_russian_date(data['date'])}"))

    doc.save(filepath)
