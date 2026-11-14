"""Gera o manual em PDF a partir de docs/MANUAL.md.

Uso (na pasta do projecto):  python docs/gerar_manual_pdf.py [saida.pdf]
Requer fpdf2 e matplotlib (as fontes DejaVu vêm com o matplotlib). Não altera o MANUAL.md.
"""
import os
import re
import sys

import matplotlib
from fpdf import FPDF
from fpdf.enums import TableBordersLayout

HERE = os.path.dirname(os.path.abspath(__file__))
SOURCE = os.path.join(HERE, "MANUAL.md")
DEFAULT_OUTPUT = os.path.join(os.path.dirname(HERE), "Manual do Utilizador — TermiFolha v{version}.pdf")
FONTS = os.path.join(matplotlib.get_data_path(), "fonts", "ttf")
INLINE = re.compile(r"(\*\*[^*]+\*\*|`[^`]+`|\*[^*\s][^*]*\*)")
BLUE = (16, 41, 160)


def clean(text):
    """Texto sem marcas de Markdown (para células de tabela)."""
    return re.sub(r"\*\*|`", "", re.sub(r"(?<!\*)\*([^*]+)\*(?!\*)", r"\1", text)).replace("\\|", "|")


class Manual(FPDF):
    def __init__(self, title):
        super().__init__(format="A4")
        self.doc_title = title
        self.set_margins(18, 18, 18)
        self.set_auto_page_break(True, margin=18)
        self.add_font("DejaVu", "", os.path.join(FONTS, "DejaVuSans.ttf"))
        self.add_font("DejaVu", "B", os.path.join(FONTS, "DejaVuSans-Bold.ttf"))
        self.add_font("DejaVu", "I", os.path.join(FONTS, "DejaVuSans-Oblique.ttf"))
        self.add_font("Mono", "", os.path.join(FONTS, "DejaVuSansMono.ttf"))

    def footer(self):
        self.set_y(-12)
        self.set_font("DejaVu", "", 8)
        self.set_text_color(120, 120, 120)
        self.cell(0, 6, f"{self.doc_title}  ·  página {self.page_no()}", align="C")

    def inline(self, text, size=10, height=5.2):
        """Escreve texto com **negrito**, *itálico* e `código` e muda de linha."""
        for part in INLINE.split(text):
            if not part:
                continue
            if part.startswith("**"):
                self.set_font("DejaVu", "B", size)
                part = part[2:-2]
            elif part.startswith("`"):
                self.set_font("Mono", "", size - 1)
                part = part[1:-1]
            elif part.startswith("*") and len(part) > 2:
                self.set_font("DejaVu", "I", size)
                part = part[1:-1]
            else:
                self.set_font("DejaVu", "", size)
            self.write(height, part)
        self.set_font("DejaVu", "", size)
        self.ln(height)

    def heading(self, level, text):
        sizes = {1: 20, 2: 14}
        self.ln(4 if level == 2 else 0)
        self.set_font("DejaVu", "B", sizes[level])
        self.set_text_color(*BLUE)
        self.multi_cell(0, 9 if level == 1 else 7, clean(text), new_x="LMARGIN", new_y="NEXT")
        self.set_text_color(0, 0, 0)
        self.ln(1)

    def code(self, lines):
        self.set_font("Mono", "", 8.5)
        self.set_fill_color(240, 240, 240)
        for line in lines:
            self.multi_cell(0, 4.8, line or " ", fill=True, new_x="LMARGIN", new_y="NEXT")
        self.ln(2)
        self.set_font("DejaVu", "", 10)

    def bullets(self, items):
        for item in items:
            self.set_x(self.l_margin + 4)
            self.write(5.2, "•  ")
            self.inline(item)
        self.ln(1)

    def table(self, rows):
        header, body = rows[0], rows[1:]
        self.set_font("DejaVu", "", 8.5)
        widths = _column_widths(rows, self.w - self.l_margin - self.r_margin)
        with super().table(
            col_widths=widths, borders_layout=TableBordersLayout.SINGLE_TOP_LINE, line_height=4.8, text_align="LEFT",
            headings_style=__import__("fpdf").FontFace(emphasis="BOLD", fill_color=(225, 232, 250)),
        ) as table:
            for cells in [header, *body]:
                row = table.row()
                for cell in cells:
                    row.cell(clean(cell))
        self.ln(3)
        self.set_font("DejaVu", "", 10)


def _column_widths(rows, total):
    """Larguras proporcionais ao texto de cada coluna (com mínimo e máximo)."""
    n = len(rows[0])
    sizes = [min(60, max(6, max(len(clean(r[i])) for r in rows if i < len(r)))) for i in range(n)]
    return tuple(total * s / sum(sizes) for s in sizes)


def split_row(line):
    return [c.strip() for c in re.split(r"(?<!\\)\|", line.strip().strip("|"))]


def render(markdown, pdf):
    lines = markdown.splitlines()
    i = 0
    paragraph = []

    def flush():
        if paragraph:
            pdf.inline(" ".join(paragraph))
            pdf.ln(1.5)
            paragraph.clear()

    pdf.add_page()
    pdf.set_font("DejaVu", "", 10)
    while i < len(lines):
        line = lines[i]
        if line.startswith("```"):
            flush()
            block = []
            i += 1
            while i < len(lines) and not lines[i].startswith("```"):
                block.append(lines[i])
                i += 1
            pdf.code(block)
        elif line.startswith("# ") or line.startswith("## "):
            flush()
            level = 1 if line.startswith("# ") else 2
            pdf.heading(level, line.lstrip("# "))
        elif line.startswith("|"):
            flush()
            rows = []
            while i < len(lines) and lines[i].startswith("|"):
                if not re.fullmatch(r"\|[\s:\-|]+\|", lines[i]):
                    rows.append(split_row(lines[i]))
                i += 1
            n = len(rows[0])  # um `|` dentro de código gera células a mais: juntam-se à última
            pdf.table([r[:n - 1] + [" | ".join(r[n - 1:])] if len(r) > n else r + [""] * (n - len(r)) for r in rows])
            continue
        elif line.startswith("- "):
            flush()
            items = []
            while i < len(lines) and lines[i].startswith("- "):
                items.append(lines[i][2:])
                i += 1
            pdf.bullets(items)
            continue
        elif not line.strip():
            flush()
        else:
            paragraph.append(line.strip())
        i += 1
    flush()


def main(argv):
    with open(SOURCE, encoding="utf-8") as f:
        markdown = f.read()
    title = markdown.splitlines()[0].lstrip("# ").strip()
    version = re.search(r"v(\d+\.\d+)", title).group(1)
    output = argv[1] if len(argv) > 1 else DEFAULT_OUTPUT.format(version=version)
    pdf = Manual(title)
    pdf.set_title(title)
    pdf.set_author("Gabriel Gomes Alves")
    render(markdown, pdf)
    pdf.output(output)
    print(f"Gerado: {output} ({pdf.page_no()} páginas)")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
