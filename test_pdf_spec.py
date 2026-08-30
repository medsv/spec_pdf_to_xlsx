"""Скрипт для прямого вызова функции pdf_spec_to_row_list без запуска streamlit.

Пример использования:
    python test_pdf_spec.py path/to/spec.pdf

Результат выводится в терминал в виде строк спецификации.
"""

import sys
import pprint
from pathlib import Path

from libs.spec_pdf_to_xlsx import pdf_spec_to_row_list
from libs.utils import row_list_to_md


def main(pdf_path: str) -> None:
    path = Path(pdf_path)
    if not path.exists():
        print(f"Ошибка: файл не найден: {path}")
        sys.exit(1)

    print(f"Обработка файла: {path.name}")
    try:
        spec, template_cols_count, restored_rows = pdf_spec_to_row_list(str(path))
        print(row_list_to_md(spec))
    except Exception as e:
        print(f"Ошибка при обработке: {e}")
        sys.exit(1)

    print(f"\nНайдено строк спецификации: {len(spec)}\n")
    #pprint.pprint(spec, width=200)




if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Использование: python test_pdf_spec.py <путь_к_pdf>")
        sys.exit(1)
    main(sys.argv[1])