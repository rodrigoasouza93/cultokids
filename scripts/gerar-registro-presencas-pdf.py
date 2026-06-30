from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "output" / "pdf" / "modelo-registro-presencas-culto-kids.pdf"

PAGE_W, PAGE_H = landscape(A4)
MARGIN = 12 * mm
BLUE = colors.HexColor("#0078D7")
DARK_BLUE = colors.HexColor("#005A9E")
LIGHT_BLUE = colors.HexColor("#EAF5FF")
SOFT_YELLOW = colors.HexColor("#FFF2B8")
CORAL = colors.HexColor("#FF8A7A")
GREEN = colors.HexColor("#8FD694")
TEXT = colors.HexColor("#1F2933")
LINE = colors.HexColor("#8FA7B8")


def draw_round_rect(c, x, y, w, h, fill, stroke=None, radius=4):
    c.setFillColor(fill)
    if stroke:
        c.setStrokeColor(stroke)
    else:
        c.setStrokeColor(fill)
    c.roundRect(x, y, w, h, radius, stroke=1, fill=1)


def centered_text(c, text, x, y, w, size=9, bold=False, color=TEXT):
    c.setFillColor(color)
    c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
    c.drawCentredString(x + w / 2, y, text)


def left_text(c, text, x, y, size=9, bold=False, color=TEXT):
    c.setFillColor(color)
    c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
    c.drawString(x, y, text)


def line_field(c, label, x, y, w):
    left_text(c, label, x, y + 2, 10, bold=True, color=DARK_BLUE)
    c.setStrokeColor(LINE)
    c.setLineWidth(0.8)
    c.line(x + 24 * mm, y, x + w, y)


def draw_attendance_table(c, x, y_top, w, row_h):
    num_w = 9 * mm
    name_w = 74 * mm
    date_w = 24 * mm
    obs_w = w - num_w - name_w - (date_w * 5)
    header_h = 11 * mm

    columns = [
        ("#", num_w),
        ("Nome do aluno", name_w),
        ("Aula 1\nData: ___/___", date_w),
        ("Aula 2\nData: ___/___", date_w),
        ("Aula 3\nData: ___/___", date_w),
        ("Aula 4\nData: ___/___", date_w),
        ("Aula 5\nData: ___/___", date_w),
        ("Observações", obs_w),
    ]

    draw_round_rect(c, x, y_top - header_h, w, header_h, BLUE, BLUE, radius=3)
    c.setStrokeColor(colors.white)
    c.setLineWidth(0.6)

    current_x = x
    for label, col_w in columns:
        lines = label.split("\n")
        centered_text(c, lines[0], current_x, y_top - 5 * mm, col_w, 9, bold=True, color=colors.white)
        if len(lines) > 1:
            centered_text(c, lines[1], current_x, y_top - 9 * mm, col_w, 7.5, color=colors.white)
        c.line(current_x, y_top - header_h, current_x, y_top)
        current_x += col_w
    c.line(x + w, y_top - header_h, x + w, y_top)

    table_h = header_h + (row_h * 15)
    for row in range(15):
        y = y_top - header_h - (row * row_h)
        if row % 2 == 0:
            c.setFillColor(colors.HexColor("#F8FCFF"))
            c.rect(x, y - row_h, w, row_h, stroke=0, fill=1)

    c.setStrokeColor(LINE)
    c.setLineWidth(0.7)
    c.rect(x, y_top - table_h, w, table_h, stroke=1, fill=0)

    current_x = x
    for _, col_w in columns:
        c.line(current_x, y_top - table_h, current_x, y_top)
        current_x += col_w
    c.line(x + w, y_top - table_h, x + w, y_top)

    for row in range(15):
        y = y_top - header_h - (row * row_h)
        c.setStrokeColor(LINE)
        c.line(x, y - row_h, x + w, y - row_h)
        centered_text(c, str(row + 1), x, y - 4.8 * mm, num_w, 8, color=TEXT)

    return y_top - table_h


def draw_totals_table(c, x, y_top, w):
    header_h = 8 * mm
    row_h = 7.5 * mm
    label_w = 40 * mm
    date_w = (w - label_w) / 5
    labels = ["Presentes", "Ausentes", "Visitantes", "Ofertas"]
    fills = [GREEN, SOFT_YELLOW, colors.HexColor("#D9C7FF"), CORAL]

    left_text(c, "Resumo quantitativo por aula", x, y_top + 4 * mm, 11, bold=True, color=DARK_BLUE)
    draw_round_rect(c, x, y_top - header_h, w, header_h, DARK_BLUE, DARK_BLUE, radius=3)

    centered_text(c, "Informação", x, y_top - 6 * mm, label_w, 8.5, bold=True, color=colors.white)
    for idx in range(5):
        centered_text(c, f"Aula {idx + 1}", x + label_w + idx * date_w, y_top - 6 * mm, date_w, 8.5, bold=True, color=colors.white)

    total_h = header_h + row_h * len(labels)
    c.setStrokeColor(LINE)
    c.setLineWidth(0.7)
    c.rect(x, y_top - total_h, w, total_h, stroke=1, fill=0)

    c.line(x + label_w, y_top - total_h, x + label_w, y_top)
    for idx in range(6):
        col_x = x + label_w + idx * date_w
        c.line(col_x, y_top - total_h, col_x, y_top)

    for row, label in enumerate(labels):
        y = y_top - header_h - row * row_h
        c.setFillColor(fills[row])
        c.rect(x, y - row_h, label_w, row_h, stroke=0, fill=1)
        left_text(c, label, x + 4 * mm, y - 5 * mm, 9.5, bold=True, color=TEXT)
        c.setStrokeColor(LINE)
        c.line(x, y - row_h, x + w, y - row_h)


def build_pdf():
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUTPUT), pagesize=landscape(A4))

    c.setTitle("Modelo de Registro de Presenças - Culto Kids")
    c.setAuthor("Culto Kids")

    draw_round_rect(c, MARGIN, PAGE_H - MARGIN - 18 * mm, PAGE_W - 2 * MARGIN, 18 * mm, BLUE, BLUE, radius=5)
    left_text(c, "Culto Kids", MARGIN + 8 * mm, PAGE_H - MARGIN - 7 * mm, 18, bold=True, color=colors.white)
    left_text(c, "Registro mensal de presenças", MARGIN + 8 * mm, PAGE_H - MARGIN - 13 * mm, 10.5, color=colors.white)
    field_y = PAGE_H - MARGIN - 28 * mm
    line_field(c, "Mês/Ano:", MARGIN, field_y, 70 * mm)
    line_field(c, "Tema do mês:", MARGIN + 85 * mm, field_y, PAGE_W - MARGIN - (MARGIN + 85 * mm))

    table_top = field_y - 8 * mm
    table_bottom = draw_attendance_table(c, MARGIN, table_top, PAGE_W - 2 * MARGIN, 6.2 * mm)
    draw_totals_table(c, MARGIN, table_bottom - 9 * mm, PAGE_W - 2 * MARGIN)

    c.showPage()
    c.save()


if __name__ == "__main__":
    build_pdf()
    print(OUTPUT)
