
import os
import re
import math
from datetime import date, datetime
from pathlib import Path
from typing import List, Tuple, Optional, Set

import openpyxl

try:
    import xlrd  # for .xls reading
except ImportError:  # optional
    xlrd = None

try:
    import xlwings as xw  # optional, for .xls -> .xlsx conversion when needed
except ImportError:
    xw = None


class ColumnParseError(ValueError):
    pass


def col_to_index0(col: str) -> int:
    s = col.strip()
    if not s:
        raise ColumnParseError("Колонка не указана")
    if s.isdigit():
        n = int(s)
        if n <= 0:
            raise ColumnParseError("Номер колонки должен быть > 0")
        return n - 1
    acc = 0
    for ch in s.upper():
        if not ('A' <= ch <= 'Z'):
            raise ColumnParseError("Некорректная колонка. Пример: A, B, 1, 2, AA")
        acc = acc * 26 + (ord(ch) - ord('A') + 1)
    return acc - 1


def index0_to_col(idx: int) -> str:
    n = idx + 1
    out = []
    while n > 0:
        rem = (n - 1) % 26
        out.append(chr(ord('A') + rem))
        n = (n - 1) // 26
    return ''.join(reversed(out))


def normalize(s: Optional[str], ignore_case: bool) -> str:
    if not s:
        return ''
    s = str(s).strip()
    s = re.sub(r"\s+", " ", s)
    if ignore_case:
        s = s.lower()
    return s


def normalize_dish(s: Optional[str], ignore_case: bool) -> str:
    s = normalize(s, ignore_case)
    if not s:
        return ''
    s = re.sub(r"\((?:[^)]*?\d\s*(?:к?кал|ккал|г|гр|грамм(?:а|ов)?|мл|л|кг|руб\.?|р\.?|₽)[^)]*?)\)", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\b\d+[\.,]?\d*\s*(?:к?кал|ккал|г|гр|грамм(?:а|ов)?|мл|л|кг)\b\.?", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\b\d+[\.,]?\d*\s*(?:руб\.?|р\.?|₽)\b\.?", "", s, flags=re.IGNORECASE)
    s = re.sub(r"(?:₽|руб\.?|р\.?)\s*\d+[\.,]?\d*", "", s, flags=re.IGNORECASE)
    s = re.sub(r"[\s,;:.-]+$", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def levenshtein(a: str, b: str) -> int:
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)
    dp = [list(range(len(b) + 1))]
    dp += [[i] + [0] * len(b) for i in range(1, len(a) + 1)]
    for i in range(1, len(a) + 1):
        for j in range(1, len(b) + 1):
            cost = 0 if a[i - 1] == b[j - 1] else 1
            dp[i][j] = min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + cost)
    return dp[-1][-1]


def sim_percent(a: str, b: str) -> int:
    if not a and not b:
        return 100
    dist = levenshtein(a, b)
    m = max(len(a), len(b))
    if m == 0:
        return 100
    return int(round((1.0 - dist / m) * 100.0))


def get_sheet_names(path: str) -> List[str]:
    ext = Path(path).suffix.lower()
    if ext in ('.xlsx', '.xlsm'):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            return wb.sheetnames
        finally:
            wb.close()
    elif ext == '.xls':
        if xlrd is None:
            raise RuntimeError("Для .xls установите xlrd==1.2.0")
        book = xlrd.open_workbook(path, on_demand=True)
        try:
            return book.sheet_names()
        finally:
            book.release_resources()
    else:
        raise RuntimeError("Неподдерживаемый формат файла. Используйте .xls/.xlsx/.xlsm")


def read_cell_values(path: str, sheet_name: str) -> List[List[Optional[str]]]:
    """Читает лист как сетку значений (строки -> ячейки -> строка)."""
    ext = Path(path).suffix.lower()
    values: List[List[Optional[str]]] = []
    if ext in ('.xlsx', '.xlsm'):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            sh = wb[sheet_name]
            for row in sh.iter_rows(values_only=True):
                values.append([None if v is None else str(v) for v in row])
        finally:
            wb.close()
    elif ext == '.xls':
        if xlrd is None:
            raise RuntimeError("Для .xls установите xlrd==1.2.0")
        book = xlrd.open_workbook(path)
        try:
            sh = book.sheet_by_name(sheet_name)
            for r in range(sh.nrows):
                row = []
                for c in range(sh.ncols):
                    v = sh.cell_value(r, c)
                    row.append(None if v == '' else str(v))
                values.append(row)
        finally:
            book.release_resources()
    else:
        raise RuntimeError("Неподдерживаемый формат")
    return values


def auto_detect_dish_column(path: str, sheet_name: str) -> Tuple[str, int]:
    vals = read_cell_values(path, sheet_name)
    keys = ["блюд", "наимен", "назван", "позици", "меню"]
    max_r = min(10, len(vals))
    for r in range(max_r):
        row = vals[r] if r < len(vals) else []
        for c, v in enumerate(row):
            s = normalize(v or '', True)
            if not s:
                continue
            for k in keys:
                if k in s:
                    return (index0_to_col(c), r + 1)  # column letter, header row 1-based
    # default
    return ("A", 1)


def ensure_xlsx(path: str) -> str:
    ext = Path(path).suffix.lower()
    if ext in ('.xlsx', '.xlsm'):
        return path
    if ext == '.xls':
        if xw is None:
            raise RuntimeError("Для .xls нужен xlwings (и установленный Excel), либо преобразуйте файл в .xlsx вручную.")
        # convert via Excel automation
        src = Path(path)
        dst = src.with_suffix('.xlsx')
        app = xw.App(visible=False, add_book=False)
        try:
            wb = app.books.open(str(src))
            wb.save(str(dst))
            wb.close()
        finally:
            app.quit()
        return str(dst)
    raise RuntimeError("Неподдерживаемый формат")


def _extract_dates_from_text(s: str) -> List[date]:
    """Ищет в строке даты в популярных форматах и возвращает список дат.
    Поддерживает: dd.mm.yyyy, dd-mm-yyyy, yyyy-mm-dd, dd.mm, dd-mm (год подставляется текущий),
    а также русские месяцы вида "5 сентября 2025" или "05 сен 2025".
    """
    if not s:
        return []
    s_norm = str(s).strip()
    out: List[date] = []
    today = date.today()

    # dd.mm.yyyy, dd-mm-yyyy, dd/mm/yyyy
    for m in re.finditer(r"\b(\d{1,2})[./\-](\d{1,2})[./\-](\d{2,4})\b", s_norm):
        d, mth, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        try:
            out.append(date(y, mth, d))
        except ValueError:
            pass

    # yyyy-mm-dd, yyyy.mm.dd
    for m in re.finditer(r"\b(\d{4})[./\-](\d{1,2})[./\-](\d{1,2})\b", s_norm):
        y, mth, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            out.append(date(y, mth, d))
        except ValueError:
            pass

    # dd.mm or dd-mm (assume current year)
    for m in re.finditer(r"\b(\d{1,2})[./\-](\d{1,2})(?![./\-]\d)\b", s_norm):
        d, mth = int(m.group(1)), int(m.group(2))
        try:
            out.append(date(today.year, mth, d))
        except ValueError:
            pass

    # Russian month names (both nominative and genitive, short forms too)
    months = {
        'янв': 1, 'январь': 1, 'января': 1,
        'фев': 2, 'февраль': 2, 'февраля': 2,
        'мар': 3, 'март': 3, 'марта': 3,
        'апр': 4, 'апрель': 4, 'апреля': 4,
        'май': 5, 'мая': 5,
        'июн': 6, 'июнь': 6, 'июня': 6,
        'июл': 7, 'июль': 7, 'июля': 7,
        'авг': 8, 'август': 8, 'августа': 8,
        'сен': 9, 'сентябрь': 9, 'сентября': 9,
        'oct': 10, 'окт': 10, 'октябрь': 10, 'октября': 10,
        'ноя': 11, 'ноябрь': 11, 'ноября': 11,
        'дек': 12, 'декабрь': 12, 'декабря': 12,
    }
    # 5 сентября 2025 / 5 сен 2025 / 5 сентября
    for m in re.finditer(r"\b(\d{1,2})\s+([A-Za-zА-Яа-яёЁ]+)\s*(\d{4})?\b", s_norm):
        d = int(m.group(1))
        mon_str = m.group(2).lower()
        y_str = m.group(3)
        mon_str = mon_str.replace('ё', 'е')
        if mon_str in months:
            mth = months[mon_str]
            y = int(y_str) if y_str else today.year
            try:
                out.append(date(y, mth, d))
            except ValueError:
                pass

    # De-duplicate
    uniq = []
    seen = set()
    for dt in out:
        if dt not in seen:
            seen.add(dt)
            uniq.append(dt)
    return uniq


def _extract_best_date_from_file(path: str, sheet_name: Optional[str]) -> Optional[date]:
    """Пытается извлечь дату из имени файла и содержимого листа.
    Возвращает дату, максимально близкую к сегодня, отдавая приоритет последней дате не позже сегодня.
    """
    candidates: List[date] = []
    # from filename
    candidates += _extract_dates_from_text(Path(path).name)

    # from sheet content (scan small top area)
    try:
        if sheet_name:
            vals = read_cell_values(path, sheet_name)
            max_r = min(20, len(vals))
            max_c = 15
            for r in range(max_r):
                row = vals[r] if r < len(vals) else []
                for c in range(min(max_c, len(row))):
                    v = row[c]
                    if v is None:
                        continue
                    # if openpyxl returned datetime/date objects, handle them
                    if isinstance(v, (datetime, date)):
                        d = v.date() if isinstance(v, datetime) else v
                        candidates.append(d)
                    else:
                        candidates += _extract_dates_from_text(str(v))
    except Exception:
        # ignore any parsing errors, fallback to filename-based only
        pass

    if not candidates:
        return None

    today = date.today()
    not_future = [d for d in candidates if d <= today]
    if not_future:
        return max(not_future)
    # если все даты в будущем — возьмём ближайшую будущую
    return min(candidates)


def _find_category_ranges(values: List[List[Optional[str]]], synonyms_map: dict) -> dict:
    """Возвращает {canonical_key: (start_row_inclusive, end_row_inclusive)} в 1-базисе.
    Ищет по синонимам (ключи словаря), возвращает диапазоны по каноническим ключам (значения словаря).
    Пример: 'салаты', 'холодные закуски' -> 'салаты и холодные закуски'.
    """
    marker_rows = []  # list of (row1based, canonical_key)
    for r, row in enumerate(values, start=1):
        row_join = ' '.join([str(c) for c in row if c not in (None, '')])
        s = normalize(row_join, True)
        found = None
        for syn, canon in synonyms_map.items():
            if syn in s:
                found = canon
                break
        if found:
            marker_rows.append((r, found))
    # Сортировка по строкам
    marker_rows.sort(key=lambda x: x[0])
    ranges = {}
    for i, (r, mk) in enumerate(marker_rows):
        start = r + 1  # после заголовка
        end = (marker_rows[i + 1][0] - 1) if i + 1 < len(marker_rows) else len(values)
        if start <= end and mk not in ranges:
            ranges[mk] = (start, end)
    return ranges


def _choose_column_for_block(values: List[List[Optional[str]]], start: int, end: int) -> str:
    """Эвристика выбора колонки для блока: сравнивает заполненность A и E, берёт более заполненную.
    start/end 1-базисные.
    """
    def non_empty_in_col(idx0: int) -> int:
        cnt = 0
        for r in range(start, end + 1):
            row = values[r - 1] if r - 1 < len(values) else []
            v = row[idx0] if idx0 < len(row) else None
            if normalize_dish(v, True):
                cnt += 1
        return cnt
    a_cnt = non_empty_in_col(col_to_index0('A'))
    e_cnt = non_empty_in_col(col_to_index0('E'))
    return 'A' if a_cnt >= e_cnt else 'E'


def compare_and_highlight(
    path1: str,
    sheet1: str,
    path2: str,
    sheet2: str,
    col1: str,
    col2: str,
    header_row1: int,
    header_row2: int,
    ignore_case: bool,
    use_fuzzy: bool,
    fuzzy_threshold: int,
    final_choice: int,  # 0=auto by date (implemented), 1=first, 2=second
) -> Tuple[str, int]:
    """
    Сравнение по категориям 'ЗАВТРАКИ' (столбец A) и 'ПЕРВЫЕ БЛЮДА' (столбец E).
    Внутри каждой категории сравниваем значения в соответствующем столбце построчно от заголовка до следующей категории/конца листа.
    При отсутствии маркеров — используем прежнюю логику с одиночной колонкой.
    Возвращает (out_path, matches).
    """
    # determine final/ref based on choice and extract dates
    d1 = _extract_best_date_from_file(path1, sheet1)
    d2 = _extract_best_date_from_file(path2, sheet2)
    
    # Определяем самую позднюю дату из двух файлов для использования в выходном файле
    latest_date = None
    if d1 and d2:
        latest_date = max(d1, d2)
    elif d1:
        latest_date = d1
    elif d2:
        latest_date = d2
    
    # ВАЖНО: Независимо от выбора final_choice, ВСЕГДА используем самую позднюю дату
    # для названия выходного файла, даже если финальным выбран файл с более ранней датой
    
    if final_choice == 1:
        final_path, final_sheet = path1, sheet1
        ref_path, ref_sheet = path2, sheet2
    elif final_choice == 2:
        final_path, final_sheet = path2, sheet2
        ref_path, ref_sheet = path1, sheet1
    else:
        # Автовыбор: берем файл с более поздней датой как финальный
        if d1 and d2:
            if d2 >= d1:
                final_path, final_sheet = path2, sheet2
                ref_path, ref_sheet = path1, sheet1
            else:
                final_path, final_sheet = path1, sheet1
                ref_path, ref_sheet = path2, sheet2
        elif d1 and not d2:
            final_path, final_sheet = path1, sheet1
            ref_path, ref_sheet = path2, sheet2
        elif d2 and not d1:
            final_path, final_sheet = path2, sheet2
            ref_path, ref_sheet = path1, sheet1
        else:
            # Если дат нет ни в одном файле, берем второй файл по умолчанию
            final_path, final_sheet = path2, sheet2
            ref_path, ref_sheet = path1, sheet1

    # Синонимы -> канонические категории
    synonyms_map = {
        "завтраки": "завтраки",
        "первые блюда": "первые блюда",
        "салаты": "салаты и холодные закуски",
        "холодные закуски": "салаты и холодные закуски",
        "салаты и холодные закуски": "салаты и холодные закуски",
        "блюда из мяса": "блюда из мяса",
        "блюда из птицы": "блюда из птицы",
        "блюда из рыбы": "блюда из рыбы",
        "гарниры": "гарниры",
    }

    # Считаем диапазоны категорий в эталонном и итоговом файлах
    ref_vals = read_cell_values(ref_path, ref_sheet)
    final_xlsx = ensure_xlsx(final_path)
    wb = openpyxl.load_workbook(final_xlsx)
    sh = wb[final_sheet]

    final_vals = []
    for row in sh.iter_rows(values_only=True):
        final_vals.append([None if v is None else str(v) for v in row])

    ref_ranges = _find_category_ranges(ref_vals, synonyms_map)
    final_ranges = _find_category_ranges(final_vals, synonyms_map)

    # Если категории не обнаружены — fallback к старой логике одной колонки
    if not ref_ranges or not final_ranges:
        # load reference set по прежней колонке (col2 как ref, кол1 как final), используя входные col1/col2
        ref_idx = col_to_index0(col2)
        ref_set: Set[str] = set()
        for r in range(max(0, header_row2), len(ref_vals)):
            row = ref_vals[r] if r < len(ref_vals) else []
            v = row[ref_idx] if ref_idx < len(row) else None
            name = normalize_dish(v, ignore_case)
            if name:
                ref_set.add(name)

        def is_match_fallback(dish: str) -> bool:
            if not use_fuzzy:
                return dish in ref_set
            best = 0
            for s in ref_set:
                sim = sim_percent(dish, s)
                if sim > best:
                    best = sim
                if best >= fuzzy_threshold:
                    return True
            return best >= fuzzy_threshold

        from openpyxl.styles import Font
        red_font = Font(color="FF0000")
        idx_final = col_to_index0(col1)
        matches = 0
        for r in range(1, sh.max_row + 1):
            if r <= max(1, header_row1):
                continue
            cell = sh.cell(row=r, column=idx_final + 1)
            text = normalize_dish(str(cell.value) if cell.value is not None else '', ignore_case)
            if text and is_match_fallback(text):
                cell.font = red_font
                matches += 1
        # Дата используется только в названии файла, не добавляем её в содержимое
        out_path = make_final_output_path(final_xlsx, latest_date)
        wb.save(out_path)
        wb.close()
        return out_path, matches

    # Построим множества по категориям из ref, выбирая колонку A или E по заполненности блока
    ref_sets: dict = {}
    ref_cols_used: dict = {}
    for cat, (start, end) in ref_ranges.items():
        col_letter = _choose_column_for_block(ref_vals, start, end)
        ref_cols_used[cat] = col_letter
        idx = col_to_index0(col_letter)
        items: Set[str] = set()
        for r in range(start, min(end, len(ref_vals)) + 1):
            row = ref_vals[r - 1] if r - 1 < len(ref_vals) else []
            v = row[idx] if idx < len(row) else None
            name = normalize_dish(v, ignore_case)
            if name:
                items.add(name)
        ref_sets[cat] = items

    def is_match_cat(cat: str, dish: str) -> bool:
        sset = ref_sets.get(cat, set())
        if not use_fuzzy:
            return dish in sset
        best = 0
        for s in sset:
            sim = sim_percent(dish, s)
            if sim > best:
                best = sim
            if best >= fuzzy_threshold:
                return True
        return best >= fuzzy_threshold

    from openpyxl.styles import Font
    red_font = Font(color="FF0000")

    matches = 0
    for cat, (start, end) in final_ranges.items():
        # Используем ту же колонку, что и в ref, если она была определена, иначе выберем по заполненности в final
        col_letter = ref_cols_used.get(cat) or _choose_column_for_block(final_vals, start, end)
        idx = col_to_index0(col_letter)
        for r in range(start, min(end, sh.max_row) + 1):
            cell = sh.cell(row=r, column=idx + 1)
            text = normalize_dish(str(cell.value) if cell.value is not None else '', ignore_case)
            if text and is_match_cat(cat, text):
                cell.font = red_font
                matches += 1

    # Дата используется только в названии файла, не добавляем её в содержимое
    out_path = make_final_output_path(final_xlsx, latest_date)
    wb.save(out_path)
    wb.close()
    return out_path, matches


def make_final_output_path(original_path: str, latest_date: Optional[date] = None) -> str:
    p = Path(original_path)
    if latest_date:
        date_str = latest_date.strftime("%d.%m.%Y")
        return str(p.with_name(p.stem + f"_сравнение_{date_str}" + p.suffix))
    else:
        return str(p.with_name(p.stem + "_сравнение" + p.suffix))


def _add_date_info_to_worksheet(worksheet, latest_date: date) -> None:
    """Добавляет информацию о дате сравнения в верхнюю часть листа."""
    try:
        from openpyxl.styles import Font, Alignment
        
        # Сдвигаем содержимое вниз на одну строку
        worksheet.insert_rows(1)
        
        # Добавляем информацию о дате в первую строку
        date_str = latest_date.strftime("%d.%m.%Y")
        cell = worksheet.cell(row=1, column=1)
        cell.value = f"Результат сравнения меню на {date_str}"
        cell.font = Font(bold=True, size=14)
        cell.alignment = Alignment(horizontal='left')
        
        # Объединяем ячейки для красивого отображения
        worksheet.merge_cells('A1:E1')
        
    except Exception:
        # Если что-то пошло не так, просто игнорируем добавление даты
        pass


def auto_header_row_by_markers(path: str, sheet_name: str) -> int:
    """
    Ищет строки-маркеры (например, 'ЗАВТРАКИ', 'ПЕРВЫЕ БЛЮДА') и возвращает номер строки заголовка,
    после которой начинается сравнение (то есть саму строку маркера).
    Если найдено несколько таких строк, берём первую сверху. Если не найдено — 1.
    """
    markers = ["завтраки", "первые блюда"]
    try:
        vals = read_cell_values(path, sheet_name)
    except Exception:
        return 1
    best = None
    max_scan = min(len(vals), 500)
    for r in range(max_scan):
        row = vals[r]
        for cell in row:
            if cell is None:
                continue
            s = normalize(str(cell), ignore_case=True)
            if any(m in s for m in markers):
                best = r + 1  # 1-based
                break
        if best is not None:
            break
    return best if best is not None else 1

