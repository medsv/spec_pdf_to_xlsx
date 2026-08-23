import re

def gost_spec_title(template_cols_count):
     """Заголовок таблицы спецификации по ГОСТ 21.110-2013"""
     spec_title = ["Поз.",
                   "Наименование и техническая характеристика",
                   "Тип, марка, обозначение документа, опросного листа",
                   "Код продукции",
                   "Поставщик",
                   "Ед. измерения",
                   "Количество",
                   "Масса 1 ед., кг",
                   "Примечание"]
     if template_cols_count == 10: spec_title.insert(1, "Код KKS")


     return spec_title

def normalize_row(spec_line, prev_spec_line, template_cols_count):
    """Очищает строку: заменяет удаляет лишние пробелы."""
    for col in spec_line:
        if not isinstance(col, str):
            continue
        # Замена символов перевода строки на пробел
        # text = text.replace('\n', ' ')
        # Удаление множественных пробелов (один или более пробелов заменяем на один)
        col = re.sub(r' +', ' ', col)
        # Удаление пробелов в начале и конце (опционально)
        col = col.strip()
        # if col == '': col = "—"  # замена '' на длинное тире (для нейросетей)
    # ставим пробел при ГОСТ12345-67
    spec_line[template_cols_count-7] = re.sub(r'ГОСТ(\d)', r'ГОСТ \1', spec_line[template_cols_count-7])
    # Приводим единицу измерения к стандартному обозначению
    spec_line[template_cols_count-4] = normalize_unit(spec_line[template_cols_count-4])
    # Преобразовываем str в int или float
    spec_line[template_cols_count-3] = string_to_number(spec_line[template_cols_count-3])
    spec_line[template_cols_count-2] = string_to_number(spec_line[template_cols_count-2])
    # Меняем точку на запятую в Кол-во и Ед. масса
    #spec_line[6] = spec_line[6].replace('.', ',')
    #spec_line[7] = spec_line[7].replace('.', ',')
    if spec_line[template_cols_count-8].lower().strip() in ["то же", "тоже"]:
         spec_line[template_cols_count-8] = prev_spec_line[template_cols_count-8]

def correct_row(line, pattern):
    restored_cols: list[int] = []
    for i, idx in enumerate(pattern[:-1]):  # до предпоследнего
        if not line[idx]:
            if idx+1 == pattern[i+1]: continue
            if line[idx+1]: 
                line[idx] = line[idx+1]
                restored_cols.append(i)  # индексы восстановленных ячеек
    return restored_cols

def string_to_number(col: str | None):
    if col is None:
        return None
    s = str(col).strip().replace(",", ".")
    if not s:
        return s
    try:
        return int(s)
    except ValueError:
        try:
            return float(s)
        except ValueError:
            return s

def normalize_unit(unit: str) -> str:
    """
    Приводит единицу измерения к стандартному обозначению.
    Все варианты (с точками, пробелами, разным регистром) нормализуются
    путём удаления разделителей и приведения к нижнему регистру.
    """
    # Приводим к нижнему регистру и убираем лишние пробелы
    unit = unit.lower().strip()
    # Заменяем Unicode-индексы на обычные цифры
    unit = unit.replace('²', '2').replace('³', '3')
    # Удаляем все точки и пробелы (оставляем только буквы и цифры)
    cleaned = re.sub(r'[.\s]+', '', unit)
    
    # Словарь: нормализованный ключ -> стандартное обозначение
    MAPPING = {
        # Метры
        'м': 'м',
        'метр': 'м',
        'метры': 'м',
        'метров': 'м',
        'm': 'м',
        'metr': 'м',
        # Квадратные метры
        'м2': 'м2',
        'м^2': 'м2',
        'квм': 'м2',          # кв.м, кв м, кв.м. – все сводятся к "квм"
        'квадратныйметр': 'м2',
        'квадратныеметры': 'м2',
        'квадратныхметров': 'м2',
        'sqm': 'м2',
        # Кубические метры
        'м3': 'м3',
        'м^3': 'м3',
        'кубм': 'м3',
        'кубическийметр': 'м3',
        'кубическиеметры': 'м3',
        'кубическихметров': 'м3',
        'cum': 'м3',
        # Литры
        'л': 'л',
        'литр': 'л',
        'литры': 'л',
        'литров': 'л',
        'l': 'л',
        # Килограммы
        'кг': 'кг',
        'килограмм': 'кг',
        'kg': 'кг',
        # Тонны
        'т': 'т',
        'тн': 'т',
        'тонн': 'т',
        'тонны': 'т',
        'тонна': 'т',
        't': 'т',
        # Штуки
        'шт': 'шт.',
        'штука': 'шт.',
        'штук': 'шт.',
        # Упаковки
        'уп': 'уп.',
        'упак': 'уп.',
        'упаковка': 'уп.',
        'упаковки': 'уп.',
        'упаковок': 'уп.',
        # Вёдра
        'вед': 'вед.',
        'ведро': 'вед.',
        'ведра': 'вед.',
        'ведер': 'вед.',
        'вёдер': 'вед.',
        # Баллоны
        'бал': 'бал.',
        'баллон': 'бал.',
        'баллоны': 'бал.',
        'баллонов': 'бал.',
        # Погонные метры
        'погм': 'пог. м',
        'пм': 'пог. м',
        'погонныйметр': 'пог. м',
        'погонныеметры': 'пог. м',
        'погонныхметров': 'пог. м',
        'мпог': 'пог. м',
        'мп': 'пог. м',
        'метрпогонный': 'пог. м',
        'метрыпогонные': 'пог. м',
        'метровпогонных': 'пог. м',
    }
    
    # Ищем в словаре по очищенному ключу
    if cleaned in MAPPING:
        return MAPPING[cleaned]
    
    # Если не найдено – возвращаем исходную строку (или можно вернуть cleaned)
    return unit

def row_list_to_md(lines):
    """Преобразует список строк в Markdown-таблицу."""
    md = ""
    for line in lines:
        md += "| " + " | ".join(line) + " |\n"
    return md