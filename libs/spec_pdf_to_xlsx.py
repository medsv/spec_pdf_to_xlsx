#from openpyxl.cell import _CellOrMergedCell
import pdfplumber
from copy import copy
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.cell.cell import MergedCell
from libs.utils import normalize_row, gost_spec_title, correct_row, extract_gosts, encoding_correction

from openpyxl.styles import PatternFill
import unicodedata

# Ключевые настройки табличного распознавания pdfplumber.
# Используются по умолчанию, если пользователь не изменил их в интерфейсе.
DEFAULT_TABLE_SETTINGS = {
    "snap_tolerance": 3,
    "join_tolerance": 3,
    "intersection_tolerance": 3,
    "edge_min_length": 3,
}

def pdf_spec_to_row_list(pdf_path, table_settings=None):
    """Извлекает из pdf-файла спецификации список строк спецификации (pdfplumber).

    table_settings — словарь настроек извлечения таблиц для pdfplumber.
    Если не передан (None), используются DEFAULT_TABLE_SETTINGS.
    """
    if table_settings is None:
        table_settings = DEFAULT_TABLE_SETTINGS
    with pdfplumber.open(pdf_path) as pdf:
        return parse_spec(pdf.pages, table_settings)

def parse_spec(pages, table_settings):
    spec = []
    # Ключевые настройки для таблиц из узких прямоугольников

    spec_row_count = 0
    restored_rows: dict[int, list[int]] = {}
    prev_spec_line = None
    first_table = True
    template_cols_count = 9  
    for page in pages:
        tabs = page.extract_tables(table_settings)
        #tabs = page.find_tables()  # locate and extract any tables on page
        if not tabs:
            continue
        for tab in tabs:
            in_spec = False  # не дошёл до спецификации
            #lines = tab.extract()
            #lines = [[_decode_cell(c) for c in line] for line in lines]
            for line in tab:
                line = _decode_cell(line)  # GOST-шрифты, ломающие Excel
                if len(line) < template_cols_count: continue
                #print(line)
                if not in_spec:
                    encoding_correction(line)
                    #if "Примечание" in line or "Код продукции" in line:  # шапка таблицы спецификации
                    if not any("опросного" in str(s).lower().strip() for s in line): # шапка таблицы спецификации
                        if first_table:
                            continue
                        else:
                            if not all(str(dig) in line for dig in range(1,9)):
                                continue
                    else:
                        if any("kks" in str(s).lower().strip() for s in line):
                            template_cols_count = 10  # таблица с KKS
                        else: 
                            template_cols_count = 9  # таблица по ГОСТ
                    if first_table:
                        spec.append(gost_spec_title(template_cols_count))  # добавляем в спецификацию шапку по ГОСТ 21.110-2013
                        spec_row_count += 1
                        first_table = False
                    spec_line_cols_count = len(line)  # количество столбцов в pdf-таблице представления спецификации
                    pattern = detect_pattern(line, template_cols_count)  # шаблон таблицы спецификации
                    #print(pattern)
                    spec_col_count = len(pattern)
                    if spec_col_count != template_cols_count:
                        #continue
                        raise ValueError(f"Идентифицированы {spec_col_count} столбца(ов) вместо {template_cols_count}.")
                    in_spec = True # внутри спецификации
                else:
                    # if None in line[first_index: last_index+1]: break
                    if len(line) != spec_line_cols_count or None in [line[i] for i in pattern]: continue  # игнорируем строки, набор столбцов которых не соответствует ранее зафиксированному набору для cпецификации
                    restored_cols = correct_row(line, pattern)
                    spec_line = [line[i] for i in pattern]
                    if all(cell == '' for cell in spec_line): continue  # Все столбцы содержат ''
                    #if spec_line[:9] == ['1', '2', '3', '4', '5', '6', '7', '8', '9']: continue  # ['1', '2', '3', '4', '5', '6', '7', '8', '9'] игнорируем
                    # range(1,9) потому что 9-й столбец может быть испорчен номером страницы
                    if all(str(dig) in line for dig in range(1,9)): continue  
                    
                    normalize_row(spec_line, prev_spec_line, template_cols_count)
                    encoding_correction(spec_line)
                    spec.append(spec_line)
                    spec_row_count += 1
                    if restored_cols:
                        restored_rows[spec_row_count] = restored_cols
                    prev_spec_line = spec_line
                    # spec.append(normalize_row(line[first_index: last_index+1]))
    if spec == []:
        raise ValueError("В файле спецификация не найдена")
    return spec, template_cols_count, restored_rows


import re as _re

_CID_RE = _re.compile(r'\(cid:(\d+)\)')

def _decode_cell(cell):
    """Восстанавливает текст ячейки из артефактов pdfplumber/pdfminer.

    При работе с нестандартными шрифтами (напр. GOST-шрифты без корректного
    ToUnicode CMap) pdfminer может выдать токены вида `(cid:N)` вместо символов.
    Здесь они заменяются на соответствующие символы по их коду (chr(N)), после
    чего управляющие символы заменяются пробелом, чтобы не попадали в XLSX.
    """
    if not isinstance(cell, str):
        return cell
    text = _CID_RE.sub(lambda m: chr(int(m.group(1))), cell)
    # Управляющие символы (кроме табуляции, переноса строки и CR) -> пробел
    text = "".join(ch if ch.isprintable() or ch in "\t\n\r" else " " for ch in text)
    return text


def detect_pattern(line, template_cols_count):
    """Определяет шаблон таблицы спецификации по шапке спецификации (в которой все поля не пустые).
    input: line - список столбцов pdf-таблицы спецификации
    output: pattern - список индексов столбцов, содерщащих данные спецификации.
    """
    pattern = []
    for i, col in enumerate(line):
        if col == '' or col is None or "формат" in col.lower():
            continue
        pattern.append(i)
        if len(pattern) == template_cols_count + 1:  # учёт таблиц с квадратиком в шапке в правом верхнем углу
            pattern = pattern[:template_cols_count]
    return pattern
    # return [i for i, col in enumerate(line) if col not in ('', None)] # быстрее на 10–30%

from openpyxl.utils.exceptions import IllegalCharacterError

def spec_to_unique_gosts(spec, template_cols_count) -> list[str]:
    """Собирает уникальный перечень всех ГОСТов, упомянутых в спецификации.

    ГОСТы извлекаются из содержательных колонок «Наименование…» и
    «Тип, марка, обозначение…»:
      - для таблицы на 9 столбцов — индексы [1, 2];
      - для таблицы с кодом KKS (10 столбцов) — индексы [2, 3].
    Сохраняется порядок первого вхождения каждого ГОСТа.
    """
    if template_cols_count == 9:
        gost_cols = [1, 2]
    elif template_cols_count == 10:
        gost_cols = [2, 3]  # пропускаем «Код KKS» (индекс 1)
    else:
        raise ValueError(f"В спецификации должно быть 9 или 10 столбцов, а не {template_cols_count}")

    unique: dict[str, None] = {}
    for row in spec[1:]:  # первая строка — заголовок, её пропускаем
        for idx in gost_cols:
            if idx >= len(row):
                continue
            for gost in extract_gosts(row[idx]):
                if gost not in unique:
                    unique[gost] = None
    return list(unique.keys())


def fill_gosts_sheet(wb, gost_list: list[str]) -> None:
    """Записывает перечень ГОСТов во вкладку «ГОСТы».

    ГОСТы записываются в столбец A начиная со строки 2. К каждой ячейке
    применяется формат эталонной ячейки A2 (границы, шрифт, выравнивание и т.д.).
    Если ГОСТы не найдены, вкладка не изменяется (форматированная строка 2
    сохраняется как образец).
    """
    if "ГОСТы" not in wb.sheetnames:
        raise ValueError("В шаблоне отсутствует вкладка «ГОСТы»")
    ws = wb["ГОСТы"]

    # Эталонный формат берём из ячейки A2 (уже отформатированный образец)
    ref_cell = ws.cell(row=2, column=1)
    ref_style = {
        'font': copy(ref_cell.font),
        'border': copy(ref_cell.border),
        'fill': copy(ref_cell.fill),
        'alignment': copy(ref_cell.alignment),
        'number_format': ref_cell.number_format,
    }

    for offset, gost in enumerate(gost_list):
        cell = ws.cell(row=2 + offset, column=1)
        cell.value = gost
        cell.font = ref_style['font']
        cell.border = ref_style['border']
        cell.fill = ref_style['fill']
        cell.alignment = ref_style['alignment']
        cell.number_format = ref_style['number_format']


def row_list_to_xlsx_bytes(spec, template_cols_count, restored_rows, pdf_stem, template_path):
    """Формирует xlsx-файл спецификации на базе шаблона и возвращает его в виде байтов."""
    if not template_path.exists():
        raise FileNotFoundError(f"Шаблон спецификации не найден: {template_path}")

    wb = load_workbook(template_path)
    ws = wb["Спецификация"]
    # Записываем имя PDF в ячейку A1
    ws["A1"] = pdf_stem
    
    # 1. Сохраняем референсные стили из строки 3 (A3:I3)
    # Именно здесь в шаблоне настроены перенос текста, выравнивание и границы
    ref_styles = {}
    for col_idx in range(1, template_cols_count+1): # Столбцы A-I(J) (индексы 1-9(10))
        ref_cell = ws.cell(row=3, column=col_idx)
        ref_styles[col_idx] = {
            'font': copy(ref_cell.font),
            'border': copy(ref_cell.border),
            'fill': copy(ref_cell.fill),
            'alignment': copy(ref_cell.alignment),
            'number_format': ref_cell.number_format
        }

    # 2. Пропускаем первую строку spec (так как это заголовок) 
    # и начинаем вывод данных с 3-й строки Excel
    data_rows = spec[1:] if len(spec) > 1 else []
    
    # 3. Заполняем данные и применяем сохраненные стили
    for row_offset, row_values in enumerate(data_rows):
        row_idx = 3 + row_offset
        for col_idx, value in enumerate(row_values, start=1):
            #if col_idx > 9:
            #    break
                
            cell = ws.cell(row=row_idx, column=col_idx)
            # ws.cell() всегда возвращает Cell (не MergedCell),
            # но защищаемся от случая объединённой ячейки
            if isinstance(cell, MergedCell):
                continue
            try:
                cell.value = value
            except IllegalCharacterError:
                cleaned_value = "".join(ch for ch in value if not unicodedata.category(ch).startswith("C"))
                cell.value = cleaned_value

            
            # Применяем скопированные стили из эталонной строки
            styles = ref_styles.get(col_idx)
            if styles:
                cell.font = styles['font']
                cell.border = styles['border']
                cell.fill = styles['fill']
                cell.alignment = styles['alignment']
                cell.number_format = styles['number_format']

    # 4. Жестко фиксируем автофильтр на строке заголовков (строка 2)
    # Указываем диапазон до последней заполненной строки, чтобы фильтр не "уехал"
    last_row = 2 + len(data_rows)
    
    if template_cols_count == 9:
        ws.auto_filter.ref = f"A2:I{last_row}"
    elif template_cols_count == 10:
        ws.auto_filter.ref = f"A2:J{last_row}"
    else:
        raise ValueError("Для таблицы с {template_cols_count} столбцом(ами) шаблон отсутствует")

    # 5. Закрашиваем красным пустые ячейки в столбце "Наименование"
    for row_idx in range(3, last_row + 1):
        cell = ws.cell(row=row_idx, column=template_cols_count-7)
        if not cell.value:
            cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    # 6. Закрашиваем жёлтым ячейки, которые были восстановлены из "битой" табличной строки
    for row_idx, row_values in restored_rows.items():
        for col_idx in row_values:
            cell = ws.cell(row=row_idx+1, column=col_idx+1)
            cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # 7. Формируем перечень уникальных ГОСТов и заполняем вкладку «ГОСТы»
    gost_list = spec_to_unique_gosts(spec, template_cols_count)
    fill_gosts_sheet(wb, gost_list)

    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
