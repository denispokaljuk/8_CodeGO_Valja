# -*- coding: utf-8 -*-
"""
Сборщик PDF-пакета для назначения Иванюк В.С.

Скрипт:
1. Генерирует .docx из .md (используя build_docx.py)
2. Конвертирует нужные .docx в .pdf в отдельную папку

Запуск:  python scripts/build_pdf.py
Зависимости: python-docx, docx2pdf (pip install python-docx docx2pdf)
"""
import os
import sys
import shutil

try:
    from docx2pdf import convert
except ImportError:
    print("Не установлен docx2pdf. Выполните: python -m pip install docx2pdf")
    sys.exit(1)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Папка для PDF-пакета назначения Иванюк В.С.
PDF_DIR = os.path.join(ROOT, "_build", "pdf", "назначение_Иванюк")

# Список документов для включения в пакет (относительные пути от ROOT)
DOCS = [
    # Кадровые документы
    ("01_Кадровые_документы", "0_Заявление_работника_о_переводе.md"),
    ("01_Кадровые_документы", "1_Представление_о_назначении.md"),
    ("01_Кадровые_документы", "2_Приказ_о_переводе.md"),
    ("01_Кадровые_документы", "3_Приказ_о_руководстве_отделом.md"),
    ("01_Кадровые_документы", "4_Должностная_инструкция.md"),
    ("01_Кадровые_документы", "5_Доп_соглашение_к_трудовому_договору.md"),
    ("01_Кадровые_документы", "6_Лист_ознакомления.md"),
    ("01_Кадровые_документы", "7_Штатное_расписание.md"),
    ("01_Кадровые_документы", "8_Приказ_об_утверждении_штатного_расписания.md"),
    ("01_Кадровые_документы", "9_Приказ_об_установлении_надбавки.md"),
    ("01_Кадровые_документы", "10_Положение_об_аттестации_ИТР.md"),
    ("01_Кадровые_документы", "11_Информационное_письмо_руководителю_отдела.md"),
    ("01_Кадровые_документы", "12_Приказ_о_проведении_аттестации.md"),
    ("01_Кадровые_документы", "13_Характеристика-отзыв.md"),
    ("01_Кадровые_документы", "14_Приказ_по_итогам_аттестации.md"),
    # Положение об отделе
    ("02_Положение_об_отделе", "Положение_о_проектно-конструкторском_отделе.md"),
]


def main():
    # Шаг 1: Генерируем .docx (используя существующий скрипт)
    print("Шаг 1: Генерация .docx из .md...")
    import build_docx
    build_docx.main()
    print()

    # Шаг 2: Создаём папку для PDF
    os.makedirs(PDF_DIR, exist_ok=True)
    print(f"Шаг 2: Создание PDF в {os.path.relpath(PDF_DIR, ROOT)}")

    # Шаг 3: Конвертируем нужные .docx в .pdf
    count = 0
    for folder, md_name in DOCS:
        docx_name = md_name[:-3] + ".docx"
        docx_path = os.path.join(ROOT, "_build", "docx", folder, docx_name)
        pdf_name = md_name[:-3] + ".pdf"
        pdf_path = os.path.join(PDF_DIR, pdf_name)

        if not os.path.exists(docx_path):
            print(f"  ! Не найден: {os.path.relpath(docx_path, ROOT)}")
            continue

        print(f"  + {pdf_name}")
        convert(docx_path, pdf_path)
        count += 1

    print(f"\nГотово: собрано {count} PDF в {os.path.relpath(PDF_DIR, ROOT)}")
    print(f"Полный путь: {PDF_DIR}")


if __name__ == "__main__":
    main()
