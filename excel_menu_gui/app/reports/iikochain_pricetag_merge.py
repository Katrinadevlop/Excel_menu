#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""iikoChain price tag export merger.

The user workflow:
- In iikoChain: select dish(es) -> Actions -> Print price tags -> choose "Большой ценник"
- Save as Excel (.xls)
- When multiple exports need to be combined, the desired output is one file with
  price tags stacked vertically and one blank row between tags.

Requirement update (черные ценники):
- The output must use the black iikoChain price tag style ("Черные ЦЕННИК.xls").
- We keep one tag as a stored Excel template and only fill:
  name+weight, price, and composition.

Why Excel COM:
- iikoChain exports are often .xls with many merged cells and formatting.
- openpyxl cannot read .xls, and xlrd reads but cannot easily write formatting.
- On Windows with Microsoft Excel installed, COM automation can copy the exact
  ranges (including merged cells, fonts, borders, fill, row heights).

This module keeps imports Windows/Excel-specific inside the public function.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, List

import os
import shutil
import sys
import tempfile
import uuid


_TEMPLATE_FILENAME = "black_pricetag_template.xls"

# iikoChain "Большой ценник" layout (rows are 1-indexed for Excel COM)
_HEADER_ROWS = 3
_TAG_START_ROW = 4
_TAG_HEIGHT_ROWS = 8  # rows per tag block (4..11)
_SEPARATOR_ROWS = 1   # one separator row (12)
_TAG_BLOCK_ROWS = _TAG_HEIGHT_ROWS + _SEPARATOR_ROWS  # 9
_TAG_WIDTH_COLS = 5   # columns A..E

# Within a tag block (relative to tag start row / col)
_NAME_OFFSET_ROW = 1
_NAME_OFFSET_COL = 1
_COMPOSITION_OFFSET_ROW = 4
_COMPOSITION_OFFSET_COL = 1
_PRICE_OFFSET_ROW = 6
_PRICE_OFFSET_COL = 3

# Left tag is A..E (start_col=1). Sometimes iikoChain puts a second tag on the
# same rows in G..K (start_col=7).
_POSSIBLE_START_COLS = (1, 7)


@dataclass(frozen=True)
class TagData:
    name: str
    weight: str = ""
    composition: str = ""
    price: str = ""


def _resolve_template_path() -> Path:
    """Находит файл шаблона (поддерживает запуск из исходников и PyInstaller _MEIPASS)."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[2]))
    candidates = [
        # PyInstaller иногда кладёт данные в _MEIPASS/excel_menu_gui/templates
        base / "excel_menu_gui" / "templates" / _TEMPLATE_FILENAME,
        base / "templates" / _TEMPLATE_FILENAME,
        # dev/run from repo
        Path(__file__).resolve().parents[2] / "templates" / _TEMPLATE_FILENAME,
        Path.cwd() / "templates" / _TEMPLATE_FILENAME,
    ]
    for p in candidates:
        if p.exists():
            return p
    raise FileNotFoundError(
        f"Не найден шаблон черного ценника: {_TEMPLATE_FILENAME}. "
        f"Пробовали: {', '.join(str(c) for c in candidates)}"
    )


def _to_text(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _format_name(name: Any, weight: Any = "") -> str:
    n = _to_text(name)
    w = _to_text(weight)
    if not n:
        return ""
    if w:
        # avoid duplicating weight if it's already part of the name
        if w.lower() not in n.lower():
            return f"{n} {w}".strip()
    return n


def _format_price(price: Any) -> str:
    """Formats price text for the template (usually like "90 р.")."""
    s = _to_text(price)
    if not s:
        return ""

    low = s.lower()
    if ("₽" in s) or (" руб" in low) or ("руб." in low) or ("р." in low) or (" р" in low):
        return s

    # numeric string -> add "р."
    try:
        v = float(s.replace(",", "."))
        if v.is_integer():
            return f"{int(v)} р."
        # keep without scientific notation
        v_txt = ("%f" % v).rstrip("0").rstrip(".")
        return f"{v_txt} р."
    except Exception:
        return s


def _is_nonempty(v) -> bool:
    if v is None:
        return False
    if isinstance(v, str):
        return bool(v.strip())
    return True


def _extract_tags_from_sheet(ws) -> List[TagData]:
    """Extract TagData from an iikoChain export sheet.

    We don't rely on a marker text; instead we scan tag blocks starting at row 4
    with step = 9 (8 rows tag + 1 separator).
    """

    ur = ws.UsedRange
    max_row = int(ur.Row) + int(ur.Rows.Count) - 1

    tags: List[TagData] = []
    empty_blocks = 0

    row = _TAG_START_ROW
    while row <= max_row:
        found_in_row = 0

        for start_col in _POSSIBLE_START_COLS:
            name_cell = ws.Cells(row + _NAME_OFFSET_ROW, start_col + _NAME_OFFSET_COL)
            name_v = None
            try:
                name_v = name_cell.Value
            except Exception:
                name_v = None
            if not _is_nonempty(name_v):
                continue

            comp_v = None
            price_v = None
            try:
                comp_v = ws.Cells(row + _COMPOSITION_OFFSET_ROW, start_col + _COMPOSITION_OFFSET_COL).Value
            except Exception:
                comp_v = None
            try:
                price_cell = ws.Cells(row + _PRICE_OFFSET_ROW, start_col + _PRICE_OFFSET_COL)
                price_v = price_cell.Value
            except Exception:
                price_v = None

            tags.append(
                TagData(
                    name=_to_text(name_v),
                    weight="",
                    composition=_to_text(comp_v),
                    price=_to_text(price_v),
                )
            )
            found_in_row += 1

        if found_in_row == 0:
            empty_blocks += 1
            if empty_blocks >= 3:
                break
        else:
            empty_blocks = 0

        row += _TAG_BLOCK_ROWS

    return tags


def export_black_pricetags(
    tags: List[TagData],
    output_path: str,
) -> None:
    """Создаёт чёрные ценники по шаблону (Excel COM).

    Заполняет:
    - название (+вес, если передан отдельно)
    - цену (если есть)
    - состав/описание (если есть)

    Args:
        tags: Список тегов для заполнения.
        output_path: Путь выходного файла (.xls или .xlsx).
    """

    items: List[TagData] = []
    for t in (tags or []):
        if not isinstance(t, TagData):
            continue
        nm = _to_text(t.name)
        if not nm:
            continue
        items.append(
            TagData(
                name=nm,
                weight=_to_text(t.weight),
                composition=_to_text(t.composition),
                price=_to_text(t.price),
            )
        )

    if not items:
        raise ValueError("Не выбраны блюда для ценников.")

    template_path = _resolve_template_path()

    out_path = Path(output_path)
    if not out_path.suffix:
        out_path = out_path.with_suffix(".xls")

    out_ext = out_path.suffix.lower()
    if out_ext not in (".xls", ".xlsx"):
        raise ValueError("Выберите выходной файл с расширением .xls или .xlsx")

    file_format = 56 if out_ext == ".xls" else 51  # 56 = xlExcel8, 51 = xlOpenXMLWorkbook

    tmp_name = f"excel_menu_gui_pricetags_{uuid.uuid4().hex}{out_ext}"
    tmp_path = Path(tempfile.gettempdir()) / tmp_name

    try:
        import win32com.client as win32  # type: ignore
    except Exception as e:
        raise RuntimeError(
            "Для выгрузки ценников нужен установленный Microsoft Excel (Windows) и pywin32."
        ) from e

    xl = None
    wb_out = None
    wb_tpl = None

    try:
        xl = win32.Dispatch("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        try:
            xl.ScreenUpdating = False
        except Exception:
            pass
        try:
            xl.Calculation = -4135  # xlCalculationManual
        except Exception:
            pass

        wb_tpl = xl.Workbooks.Open(str(template_path), ReadOnly=True)
        ws_tpl = wb_tpl.Worksheets(1)

        wb_out = xl.Workbooks.Add()
        ws_out = wb_out.Worksheets(1)
        try:
            ws_out.Name = "Page 1"
        except Exception:
            pass

        # Важно для .xls: копируем палитру цветов, иначе черный/белый может "переехать".
        try:
            wb_out.Colors = wb_tpl.Colors
        except Exception:
            pass

        # delete extra default sheets
        try:
            while wb_out.Worksheets.Count > 1:
                wb_out.Worksheets(wb_out.Worksheets.Count).Delete()
        except Exception:
            pass

        # header + base formatting
        ws_tpl.Range(ws_tpl.Cells(1, 1), ws_tpl.Cells(_HEADER_ROWS, _TAG_WIDTH_COLS)).Copy(ws_out.Cells(1, 1))

        for c in range(1, _TAG_WIDTH_COLS + 1):
            try:
                ws_out.Columns(c).ColumnWidth = ws_tpl.Columns(c).ColumnWidth
            except Exception:
                pass

        for r in range(1, _HEADER_ROWS + 1):
            try:
                ws_out.Rows(r).RowHeight = ws_tpl.Rows(r).RowHeight
            except Exception:
                pass

        try:
            ps_src = ws_tpl.PageSetup
            ps_out = ws_out.PageSetup
            ps_out.Orientation = ps_src.Orientation
            ps_out.PaperSize = ps_src.PaperSize
            ps_out.Zoom = ps_src.Zoom
            ps_out.FitToPagesWide = ps_src.FitToPagesWide
            ps_out.FitToPagesTall = ps_src.FitToPagesTall
        except Exception:
            pass

        dest_row = _TAG_START_ROW
        for t in items:
            # copy tag block 4..11 and separator row 12 from template
            ws_tpl.Range(
                ws_tpl.Cells(_TAG_START_ROW, 1),
                ws_tpl.Cells(_TAG_START_ROW + _TAG_HEIGHT_ROWS - 1, _TAG_WIDTH_COLS),
            ).Copy(ws_out.Cells(dest_row, 1))

            ws_tpl.Range(
                ws_tpl.Cells(_TAG_START_ROW + _TAG_HEIGHT_ROWS, 1),
                ws_tpl.Cells(_TAG_START_ROW + _TAG_HEIGHT_ROWS, _TAG_WIDTH_COLS),
            ).Copy(ws_out.Cells(dest_row + _TAG_HEIGHT_ROWS, 1))

            # row heights for tag+separator
            for off in range(_TAG_BLOCK_ROWS):
                try:
                    ws_out.Rows(dest_row + off).RowHeight = ws_tpl.Rows(_TAG_START_ROW + off).RowHeight
                except Exception:
                    pass

            # заполняем поля (и очищаем примеры из шаблона)
            ws_out.Cells(dest_row + _NAME_OFFSET_ROW, 2).Value = _format_name(t.name, t.weight)

            comp = _to_text(t.composition)
            ws_out.Cells(dest_row + _COMPOSITION_OFFSET_ROW, 2).Value = comp if comp else ""

            pr = _format_price(t.price)
            ws_out.Cells(dest_row + _PRICE_OFFSET_ROW, 4).Value = pr if pr else ""

            dest_row += _TAG_BLOCK_ROWS

        try:
            wb_out.SaveAs(str(tmp_path), FileFormat=file_format)
        finally:
            try:
                wb_out.Close(False)
            except Exception:
                pass

    finally:
        try:
            if wb_tpl is not None:
                wb_tpl.Close(False)
        except Exception:
            pass
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass

    out_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        if out_path.exists():
            out_path.unlink()
        os.replace(str(tmp_path), str(out_path))
    except Exception:
        try:
            shutil.move(str(tmp_path), str(out_path))
        except Exception:
            raise


def export_black_pricetags_from_dish_names(
    dish_names: List[str],
    output_path: str,
) -> None:
    """Совместимость: список названий -> чёрные ценники."""
    tags: List[TagData] = []
    for nm in (dish_names or []):
        s = _to_text(nm)
        if not s:
            continue
        tags.append(TagData(name=s))
    export_black_pricetags(tags, output_path)


# Backward-compatible name (older UI used merge_iikochain_big_pricetags)
def merge_iikochain_big_pricetags(
    input_paths: List[str],
    output_path: str,
) -> None:
    """Deprecated: kept for compatibility.

    Previously merged iikoChain exports by markers. Now we generate from a list of
    dish names passed in input_paths.
    """
    export_black_pricetags_from_dish_names(input_paths, output_path)
