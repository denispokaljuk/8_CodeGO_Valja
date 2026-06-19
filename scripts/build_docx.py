# -*- coding: utf-8 -*-
"""
Генератор .docx из .md документов проекта.

Содержимое документов живёт в .md (единый источник). Этот скрипт собирает из них
готовые к печати/подписи .docx в _build/docx/, повторяя структуру папок.

Запуск:  python scripts/build_docx.py
Зависимость: python-docx  (pip install python-docx)

Поддерживаемый markdown:
  # ## ###            заголовки
  >> текст            абзац с выравниванием вправо (УТВЕРЖДАЮ, блок подписи)
  - текст             маркированный список
  > текст             цитата (курсив)
  | a | b |           таблица (строка-разделитель |---| отделяет шапку)
  **жирный**          жирный текст внутри строки
  <!-- ... -->        комментарий (игнорируется)
  ---                 разделитель (игнорируется)
  строки "1. ..."     остаются как есть (нумерация пунктов документа)
"""
import os
import re
import sys

try:
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
except ImportError:
    print("Не установлен python-docx. Выполните: python -m pip install python-docx")
    sys.exit(1)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_ROOT = os.path.join(ROOT, "_build", "docx")

# Какие папки конвертировать (документы для печати). Служебные файлы пропускаем.
SRC_DIRS = [
    "01_Кадровые_документы",
    "02_Положение_об_отделе",
    "03_Внедрение_и_коллектив",
    "04_Шаблоны_управления_отделом",
]
SKIP_FILES = {"README.md"}

FONT = "Times New Roman"


def set_base_style(doc):
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(12)
    p = st.paragraph_format
    p.space_after = Pt(4)
    p.line_spacing = 1.15


def add_runs_with_bold(paragraph, text):
    """Разбор **жирного** внутри строки."""
    for i, part in enumerate(re.split(r"\*\*", text)):
        if part == "":
            continue
        run = paragraph.add_run(part)
        run.font.name = FONT
        if i % 2 == 1:  # части между ** — жирные
            run.bold = True


def is_table_row(line):
    return line.strip().startswith("|") and line.strip().endswith("|")


def is_separator_row(line):
    cells = [c.strip() for c in line.strip().strip("|").split("|")]
    return all(re.fullmatch(r":?-{2,}:?", c or "-") for c in cells)


def parse_cells(line):
    return [c.strip() for c in line.strip().strip("|").split("|")]


def add_table(doc, rows):
    sep_idx = next((i for i, r in enumerate(rows) if is_separator_row(r)), None)
    header = parse_cells(rows[0]) if sep_idx else None
    body_start = sep_idx + 1 if sep_idx is not None else 0
    body = [parse_cells(r) for r in rows[body_start:]]
    ncols = max([len(header)] if header else [0] + [len(r) for r in body])
    if ncols == 0:
        return
    table = doc.add_table(rows=0, cols=ncols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    def fill(cells_text, bold=False):
        cells = table.add_row().cells
        for j in range(ncols):
            txt = cells_text[j] if j < len(cells_text) else ""
            para = cells[j].paragraphs[0]
            if bold:
                run = para.add_run(txt)
                run.font.name = FONT
                run.bold = True
                run.font.size = Pt(11)
            else:
                add_runs_with_bold(para, txt)
                for r in para.runs:
                    r.font.size = Pt(11)

    if header:
        fill(header, bold=True)
    for row in body:
        fill(row)
    doc.add_paragraph()


def render(md_path, docx_path):
    with open(md_path, encoding="utf-8") as f:
        lines = f.read().splitlines()

    doc = Document()
    set_base_style(doc)
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(1.5)

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped or stripped == "---" or stripped.startswith("<!--"):
            i += 1
            continue

        # Таблица
        if is_table_row(line):
            block = []
            while i < len(lines) and is_table_row(lines[i]):
                block.append(lines[i])
                i += 1
            add_table(doc, block)
            continue

        # Выравнивание вправо
        if stripped.startswith(">>"):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            add_runs_with_bold(p, stripped[2:].strip())
            i += 1
            continue

        # Заголовки
        if stripped.startswith("### "):
            p = doc.add_paragraph()
            run = p.add_run(stripped[4:].strip())
            run.bold = True
            run.font.name = FONT
            run.font.size = Pt(12)
            i += 1
            continue
        if stripped.startswith("## "):
            p = doc.add_paragraph()
            run = p.add_run(stripped[3:].strip())
            run.bold = True
            run.font.name = FONT
            run.font.size = Pt(13)
            i += 1
            continue
        if stripped.startswith("# "):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(stripped[2:].strip())
            run.bold = True
            run.font.name = FONT
            run.font.size = Pt(14)
            i += 1
            continue

        # Цитата
        if stripped.startswith(">"):
            p = doc.add_paragraph()
            run = p.add_run(stripped.lstrip(">").strip())
            run.italic = True
            run.font.name = FONT
            i += 1
            continue

        # Маркированный список
        if stripped.startswith("- ") or stripped.startswith("* "):
            p = doc.add_paragraph(style="List Bullet")
            add_runs_with_bold(p, stripped[2:].strip())
            i += 1
            continue

        # Обычный абзац (включая пункты "1. ...")
        p = doc.add_paragraph()
        add_runs_with_bold(p, stripped)
        i += 1

    os.makedirs(os.path.dirname(docx_path), exist_ok=True)
    doc.save(docx_path)


def main():
    count = 0
    for d in SRC_DIRS:
        src_dir = os.path.join(ROOT, d)
        if not os.path.isdir(src_dir):
            continue
        for name in sorted(os.listdir(src_dir)):
            if not name.endswith(".md") or name in SKIP_FILES:
                continue
            src = os.path.join(src_dir, name)
            dst = os.path.join(OUT_ROOT, d, name[:-3] + ".docx")
            render(src, dst)
            count += 1
            print("  +", os.path.relpath(dst, ROOT))
    print(f"Готово: собрано {count} .docx в _build/docx/")


if __name__ == "__main__":
    main()
