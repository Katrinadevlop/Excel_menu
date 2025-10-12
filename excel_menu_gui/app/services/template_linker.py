from pathlib import Path
from typing import List, Tuple, Optional
import re
import sys

try:
    import xlwings as xw
except ImportError:
    xw = None

CATEGORIES = [
    "Завтраки", "Салаты и холодные закуски", "Первые блюда", "Блюда из мяса",
    "Блюда из птицы", "Блюда из рыбы", "Гарниры"
]


def default_template_path() -> str:
    """
    Возвращает путь к шаблону меню 'Шаблон меню пример.xlsx', учитывая запуск:
    - из dev-окружения (структура репозитория)
    - из PyInstaller onefile (sys._MEIPASS)
    - из любых рабочих директорий (поиск относительно cwd)
    """
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
    repo_root = Path(__file__).resolve().parents[2]  # excel_menu_gui/
    cwd = Path.cwd()

    candidates = [
        # PyInstaller раскладка
        base / "excel_menu_gui" / "templates" / "Шаблон меню пример.xlsx",
        base / "templates" / "Шаблон меню пример.xlsx",
        # Dev-режим: путь от корня репозитория
        repo_root / "templates" / "Шаблон меню пример.xlsx",
        # На случай запуска из корня проекта как CWD
        cwd / "templates" / "Шаблон меню пример.xlsx",
        # Оставляем старые пути для совместимости (.xls)
        base / "excel_menu_gui" / "templates" / "menu_template.xls",
        base / "templates" / "menu_template.xls",
        repo_root / "templates" / "menu_template.xls",
        cwd / "templates" / "menu_template.xls",
    ]
    for p in candidates:
        if p.exists():
            return str(p)
    # Последний fallback — путь относительно текущего пакета (может не существовать)
    return str(repo_root / "templates" / "Шаблон меню пример.xlsx")


def find_headers(ws) -> dict:
    """Возвращает {row_index: (col_index, text)} для найденных заголовков категорий.
    row_index/col_index 1-based.
    """
    headers = {}
    used = ws.used_range
    rows = used.last_cell.row
    cols = used.last_cell.column
    cats_low = {c.lower(): c for c in CATEGORIES}
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            val = ws.cells(r, c).value
            if isinstance(val, str) and val.strip():
                s = val.strip().lower()
                if s in cats_low:
                    headers[r] = (c, cats_low[s])
    return headers


def block_bounds(ws, start_row: int, start_col: int, next_header_row: Optional[int]) -> Tuple[int, int, int, int]:
    """Определяет прямоугольник блока для категории.
    Возвращает (top_row, left_col, height, width) в 1-based.
    height идёт до следующего заголовка - 1, либо до первых 2 подряд пустых строк, либо границ used_range.
    width берём по максимальной занятой ширине в первых 3 строках блока.
    """
    used = ws.used_range
    last_row = used.last_cell.row
    last_col = used.last_cell.column

    top = start_row + 1
    bottom_limit = (next_header_row - 1) if next_header_row else last_row

    # Найдём высоту: до двух подряд пустых строк или до bottom_limit
    empty_streak = 0
    r = top
    while r <= bottom_limit:
        row_empty = True
        for c in range(start_col, last_col + 1):
            if ws.cells(r, c).value not in (None, ""):
                row_empty = False
                break
        if row_empty:
            empty_streak += 1
            if empty_streak >= 2:
                r -= 1  # предыдущая была последней значимой
                break
        else:
            empty_streak = 0
        r += 1
    bottom = min(r, bottom_limit)
    if bottom < top:
        bottom = top

    # Определим ширину: возьмём максимальную занятую ширину по первым 3 строкам блока
    width = 1
    check_rows = range(top, min(bottom, top + 2) + 1)
    for rr in check_rows:
        last_used = start_col
        for cc in range(start_col, last_col + 1):
            if ws.cells(rr, cc).value not in (None, ""):
                last_used = cc
        width = max(width, last_used - start_col + 1)

    height = bottom - top + 1
    return top, start_col, max(1, height), max(1, width)


def link_template_categories(template_path: str) -> str:
    if xw is None:
        raise RuntimeError("Для обновления .xls шаблона требуется xlwings и установленный Excel.")

    src_path = Path(template_path)
    out_path = src_path.with_name(src_path.stem + "_linked" + src_path.suffix)

    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(str(src_path))
        try:
            ws1 = wb.sheets[0]
            ws2 = wb.sheets[1]

            h1 = find_headers(ws1)
            h2 = find_headers(ws2)

            # Отсортируем заголовки по возрастанию строки
            sorted_h1 = sorted(h1.items(), key=lambda x: x[0])
            sorted_h2 = sorted(h2.items(), key=lambda x: x[0])

            # Для каждой категории, найденной на обоих листах, проставим формулы блок-в-блок
            for r2_idx, (c2, text2) in h2.items():
                # ищем такую же категорию на листе 1
                src_row = None
                src_col = None
                for r1_idx, (c1, text1) in h1.items():
                    if text1 == text2:
                        src_row, src_col = r1_idx, c1
                        break
                if not src_row:
                    continue

                # Определим границы блока на обоих листах
                # next header rows
                next_r1 = next((rr for rr, _ in sorted_h1 if rr > src_row), None)
                next_r2 = next((rr for rr, _ in sorted_h2 if rr > r2_idx), None)

                top1, left1, h1h, w1w = block_bounds(ws1, src_row, src_col, next_r1)
                top2, left2, h2h, w2w = block_bounds(ws2, r2_idx, c2, next_r2)

                h = min(h1h, h2h)
                w = min(w1w, w2w)

                # Проставляем формулы в ws2, ссылающиеся на ws1
                for dr in range(h):
                    for dc in range(w):
                        r_dst = top2 + dr
                        c_dst = left2 + dc
                        r_src = top1 + dr
                        c_src = left1 + dc
                        # Пример формулы: ='Лист1'!A1
                        src_addr = xw.utils.col_name(c_src) + str(r_src)
                        ws2.cells(r_dst, c_dst).formula = f"='{ws1.name}'!{src_addr}"

            wb.save(str(out_path))
        finally:
            wb.close()
    finally:
        app.quit()

    return str(out_path)

