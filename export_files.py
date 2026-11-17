from functools import lru_cache
from csv import writer
from io import BytesIO, StringIO
from math import isfinite
from pathlib import Path
import re
from xml.sax.saxutils import escape

MAX_BYTES = 12 * 1024 * 1024
EXCEL_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
GREEN = "#164C40"
PALE = "#EDF6F1"


def validate_export(data):
    def clean(value, maximum=800):
        if not isinstance(value, str) or len(value) > maximum:
            raise ValueError("Geçersiz çıktı alanı.")
        return "".join(char for char in value if char in "\n\t" or ord(char) >= 32)

    if not isinstance(data, dict) or not isinstance(data.get("sections"), list):
        raise ValueError("Çıktı verisi alınamadı.")
    title = clean(data.get("title"), 100)
    context = clean(data.get("context", ""), 1600)
    sections = []
    total_rows = 0
    if not 1 <= len(data["sections"]) <= 10:
        raise ValueError("Çıktı için bir liste veya rapor açın.")
    for section in data["sections"]:
        if not isinstance(section, dict):
            raise ValueError("Geçersiz çıktı bölümü.")
        columns = section.get("columns", [])
        rows = section.get("rows", [])
        if not isinstance(columns, list) or not 1 <= len(columns) <= 15 or not isinstance(rows, list):
            raise ValueError("Geçersiz çıktı tablosu.")
        total_rows += len(rows)
        if total_rows > 30000:
            raise ValueError("Çıktı 30.000 satırı aşıyor. Tarih aralığını daraltın.")
        checked = []
        for row in rows:
            if not isinstance(row, list) or len(row) != len(columns):
                raise ValueError("Çıktıda eksik sütun var.")
            cells = []
            for cell in row:
                if not isinstance(cell, dict):
                    raise ValueError("Geçersiz çıktı hücresi.")
                text = clean(cell.get("text", ""))
                value = cell.get("value")
                format_type = cell.get("format")
                unit = clean(cell.get("unit", ""), 250)
                if value is not None and (type(value) not in (int, float) or not isfinite(value)
                                          or format_type not in ("quantity", "money")):
                    raise ValueError("Geçersiz sayısal çıktı değeri.")
                cells.append((text, value, format_type, unit))
            checked.append(cells)
        sections.append((clean(section.get("title"), 100), [clean(col, 100) for col in columns],
                         checked, clean(section.get("empty", ""))))
    return title, context, sections


def write_csv(title, context, sections):
    stream = StringIO(newline="")
    output = writer(stream, delimiter=";")
    for index, (_, columns, rows, _) in enumerate(sections):
        if index:
            output.writerow([])
        output.writerow(columns)
        for row in rows:
            output.writerow([str(value).replace(".", ",") if value is not None else
                             ("'" + text if text.lstrip().startswith(("=", "+", "-", "@")) else text)
                             for text, value, _, _ in row])
    return stream.getvalue().encode("utf-8-sig")


def write_xlsx(title, context, sections):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    def put(sheet, row, index, value):
        cell = sheet.cell(row, index, value)
        if isinstance(value, str):
            cell.data_type = "s"
        return cell

    workbook = Workbook()
    workbook.remove(workbook.active)
    for section, columns, rows, empty in sections:
        sheet_name = re.sub(r"[:\\/?*\[\]]", " ", section).strip("' ")[:31] or "Liste"
        sheet = workbook.create_sheet(sheet_name)
        sheet.sheet_view.showGridLines = False
        sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columns))
        put(sheet, 1, 1, title + (" · " + section if section != title else "")).font = Font(
            name="Arial", size=14, bold=True, color=GREEN[1:])
        sheet.row_dimensions[1].height = 24
        if context:
            sheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(columns))
            put(sheet, 2, 1, context).alignment = Alignment(wrap_text=True, vertical="center")
            sheet.row_dimensions[2].height = 32
        widths = [min(50, max(15, len(heading) + 4)) for heading in columns]
        for index, heading in enumerate(columns, 1):
            cell = put(sheet, 4, index, heading)
            cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=GREEN[1:])
        for row_index, row in enumerate(rows, 5):
            height = 1
            for index, (text, value, format_type, unit) in enumerate(row, 1):
                cell = put(sheet, row_index, index, value if value is not None else text)
                if value is not None:
                    if format_type == "money":
                        cell.number_format = '#,##0.00 "TL"'
                    else:
                        suffix = unit.replace('"', "″").replace("\\", "/").replace("\n", " ").replace("\t", " ")[:40]
                        cell.number_format = '#,##0.###' + (f' " {suffix}"' if suffix else "")
                cell.alignment = Alignment(vertical="top", wrap_text=value is None)
                if row_index % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor=PALE[1:])
                text_width = max((len(line) for line in text.splitlines()), default=0)
                widths[index - 1] = min(50, max(widths[index - 1], text_width + 3))
                height = max(height, sum(max(1, (len(line) + 44) // 45) for line in text.splitlines()))
            if height > 1:
                sheet.row_dimensions[row_index].height = min(210, 15 * height)
        if rows:
            sheet.auto_filter.ref = f"A4:{get_column_letter(len(columns))}{4 + len(rows)}"
        elif empty:
            put(sheet, 5, 1, empty)
        for index, width in enumerate(widths, 1):
            sheet.column_dimensions[get_column_letter(index)].width = width
        sheet.freeze_panes = "A5"
    result = BytesIO()
    workbook.save(result)
    return result.getvalue()


@lru_cache(maxsize=1)
def pdf_fonts():
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    candidates = [
        (Path("C:/Windows/Fonts/arial.ttf"), Path("C:/Windows/Fonts/arialbd.ttf")),
        (Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
         Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")),
    ]
    for regular, bold in candidates:
        if regular.is_file() and bold.is_file():
            pdfmetrics.registerFont(TTFont("ReportBody", str(regular)))
            pdfmetrics.registerFont(TTFont("ReportBold", str(bold)))
            return "ReportBody", "ReportBold"
    raise RuntimeError("PDF için Türkçe karakterleri destekleyen Arial veya DejaVu Sans yazı tipi bulunamadı.")


def write_pdf(title, context, sections):
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_RIGHT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import LongTable, Paragraph, SimpleDocTemplate, Spacer, TableStyle

    regular, bold = pdf_fonts()
    output = BytesIO()
    page_width, page_height = landscape(A4)
    document = SimpleDocTemplate(output, pagesize=(page_width, page_height),
                                 leftMargin=26, rightMargin=26, topMargin=24, bottomMargin=34, title=title)
    body = ParagraphStyle("Body", fontName=regular, fontSize=8, leading=11, textColor=colors.HexColor("#19302A"))
    numeric = ParagraphStyle("Number", parent=body, alignment=TA_RIGHT)
    header = ParagraphStyle("Header", parent=body, fontName=bold, textColor=colors.white)
    section_style = ParagraphStyle("Section", parent=body, fontName=bold, fontSize=11, leading=16, spaceAfter=5)
    title_style = ParagraphStyle("Title", parent=section_style, fontSize=18, leading=24)

    def paragraph(text, style=body):
        return Paragraph(escape(text).replace("\n", "<br/>"), style)

    story = [paragraph(title, title_style)]
    if context:
        story.extend([paragraph(context), Spacer(1, 7)])
    for section, columns, rows, empty in sections:
        story.extend([Spacer(1, 8), paragraph(section, section_style)])
        if not rows:
            story.append(paragraph(empty or "Kayıt bulunamadı."))
            continue
        weights = [max(12, min(35, len(column) + 4)) for column in columns]
        for row in rows[:250]:
            for index, (text, value, _, _) in enumerate(row):
                maximum = 26 if value is not None else 43
                weights[index] = max(weights[index], min(maximum, len(text) + 2))
        available = document.width - 12
        widths = [available * weight / sum(weights) for weight in weights]
        cells = [[paragraph(column, header) for column in columns]]
        cells.extend([paragraph(text, numeric if value is not None else body)
                      for text, value, _, _ in row] for row in rows)
        table = LongTable(cells, colWidths=widths, repeatRows=1, hAlign="LEFT")
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(GREEN)),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor(PALE)]),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(table)

    def footer(canvas, doc):
        canvas.setFont(regular, 8)
        canvas.setFillColor(colors.HexColor("#48605A"))
        canvas.drawString(26, 18, "Sarf Malzeme Takip")
        canvas.drawRightString(page_width - 26, 18, f"Sayfa {doc.page}")

    document.build(story, onFirstPage=footer, onLaterPages=footer)
    return output.getvalue()
