#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""iikoChain price tag export merger.

The user workflow:
- In iikoChain: select dish(es) -> Actions -> Print price tags -> choose "Большой ценник"
- Save as Excel (.xls)
- When multiple exports need to be combined, the desired output is one file with
  price tags stacked vertically and one blank row between tags.

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
from typing import Callable, Iterable, List, Optional, Tuple

import os
import shutil
import tempfile
import uuid


_MARKER_TEXT = "Цена за порц."  # marker that exists once per price tag
_TAG_HEIGHT_ROWS = 8           # rows per tag block (start..start+7)
_TAG_WIDTH_COLS = 5            # columns per tag block (start..start+4)
_MARKER_OFFSET_ROW = 7         # marker row = start_row + 7
_MARKER_OFFSET_COL = 2         # marker col = start_col + 2 (3rd column inside tag)


@dataclass(frozen=True)
class TagRange:
    start_row: int
    start_col: int
    end_row: int
    end_col: int


def merge_iikochain_big_pricetags(
    input_paths: List[str],
    output_path: str,
    *,
    marker_text: str = _MARKER_TEXT,
) -> None:
    """Merge iikoChain exports into a single Excel file.

    Args:
        input_paths: List of source .xls/.xlsx files exported from iikoChain.
        output_path: Destination .xls or .xlsx path.
        marker_text: Text that marks a price tag block.

    Notes:
        - Requires Windows + installed Microsoft Excel.
        - Saves to a temp file first (Excel COM may be unable to SaveAs directly
          to some folders), then moves to the requested output path.
    """

    if not input_paths:
        raise ValueError("Не выбраны входные файлы ценников.")

    out_path = Path(output_path)
    if not out_path.suffix:
        # default to .xls for best compatibility with iikoChain exports
        out_path = out_path.with_suffix(".xls")

    out_ext = out_path.suffix.lower()
    if out_ext not in (".xls", ".xlsx"):
        raise ValueError("Выберите выходной файл с расширением .xls или .xlsx")

    # Excel file format constants
    file_format = 56 if out_ext == ".xls" else 51  # 56 = xlExcel8, 51 = xlOpenXMLWorkbook

    # temp target path (same extension)
    tmp_name = f"excel_menu_gui_pricetags_{uuid.uuid4().hex}{out_ext}"
    tmp_path = Path(tempfile.gettempdir()) / tmp_name

    # Import COM deps lazily to keep module importable on non-Windows.
    try:
        import win32com.client as win32  # type: ignore
    except Exception as e:
        raise RuntimeError(
            "Для объединения ценников нужен установленный Microsoft Excel (Windows) и pywin32."
        ) from e

    # Helper: iterate marker cells via Find/FindNext (fast, avoids per-cell reads).
    def _iter_marker_cells(ws) -> List[Tuple[int, int]]:
        used = ws.UsedRange
        # xlByRows=1, xlNext=1, xlPart=2, xlValues=-4163
        first = used.Find(
            What=marker_text,
            LookIn=-4163,
            LookAt=2,
            SearchOrder=1,
            SearchDirection=1,
            MatchCase=False,
        )
        if not first:
            return []

        hits: List[Tuple[int, int]] = []
        first_addr = first.Address
        cur = first
        while True:
            try:
                v = cur.Value
            except Exception:
                v = None
            if isinstance(v, str) and v.strip() == marker_text:
                hits.append((int(cur.Row), int(cur.Column)))

            cur = used.FindNext(cur)
            if (not cur) or (cur.Address == first_addr):
                break

        return hits

    def _sheet_max_bounds(ws) -> Tuple[int, int]:
        ur = ws.UsedRange
        max_row = int(ur.Row) + int(ur.Rows.Count) - 1
        max_col = int(ur.Column) + int(ur.Columns.Count) - 1
        return max_row, max_col

    def _build_tag_ranges(ws) -> List[TagRange]:
        max_row, max_col = _sheet_max_bounds(ws)
        markers = _iter_marker_cells(ws)
        ranges: set[TagRange] = set()

        for (mr, mc) in markers:
            start_row = mr - _MARKER_OFFSET_ROW
            start_col = mc - _MARKER_OFFSET_COL
            if start_row < 1 or start_col < 1:
                continue
            end_row = start_row + (_TAG_HEIGHT_ROWS - 1)
            end_col = start_col + (_TAG_WIDTH_COLS - 1)
            if end_row > max_row or end_col > max_col:
                continue
            ranges.add(TagRange(start_row=start_row, start_col=start_col, end_row=end_row, end_col=end_col))

        out = list(ranges)
        out.sort(key=lambda t: (t.start_row, t.start_col))
        return out

    xl = None
    wb_out = None
    try:
        xl = win32.Dispatch("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        try:
            xl.ScreenUpdating = False
        except Exception:
            pass
        try:
            # xlCalculationManual = -4135
            xl.Calculation = -4135
        except Exception:
            pass

        wb_out = xl.Workbooks.Add()
        ws_out = wb_out.Worksheets(1)
        try:
            ws_out.Name = "Page 1"
        except Exception:
            pass

        # delete any extra default sheets
        try:
            while wb_out.Worksheets.Count > 1:
                wb_out.Worksheets(wb_out.Worksheets.Count).Delete()
        except Exception:
            pass

        header_copied = False
        dest_row = 4

        for idx, in_path in enumerate(input_paths):
            src_path = str(Path(in_path))
            if not Path(src_path).exists():
                raise FileNotFoundError(f"Файл не найден: {src_path}")

            wb_src = xl.Workbooks.Open(src_path, ReadOnly=True)
            try:
                ws_src = wb_src.Worksheets(1)

                tag_ranges = _build_tag_ranges(ws_src)
                if not tag_ranges:
                    # no markers; skip
                    continue

                # Copy header + column widths from the first usable file.
                # iikoChain template usually has first tag starting at row 4.
                if not header_copied:
                    # Copy rows 1..3 from the source (A..E)
                    ws_src.Range(ws_src.Cells(1, 1), ws_src.Cells(3, 5)).Copy(ws_out.Cells(1, 1))

                    # Column widths A..E
                    for c in range(1, 6):
                        try:
                            ws_out.Columns(c).ColumnWidth = ws_src.Columns(c).ColumnWidth
                        except Exception:
                            pass

                    # Row heights 1..3
                    for r in range(1, 4):
                        try:
                            ws_out.Rows(r).RowHeight = ws_src.Rows(r).RowHeight
                        except Exception:
                            pass

                    # Try to copy page setup (optional, for printing).
                    try:
                        ps_src = ws_src.PageSetup
                        ps_out = ws_out.PageSetup
                        ps_out.Orientation = ps_src.Orientation
                        ps_out.PaperSize = ps_src.PaperSize
                        ps_out.Zoom = ps_src.Zoom
                        ps_out.FitToPagesWide = ps_src.FitToPagesWide
                        ps_out.FitToPagesTall = ps_src.FitToPagesTall
                    except Exception:
                        pass

                    header_copied = True
                    dest_row = 4

                # Copy each tag into A..E (always). If a tag is in the right column (G..K),
                # we still place it to the left for consistent printing.
                for tr in tag_ranges:
                    src_rng = ws_src.Range(ws_src.Cells(tr.start_row, tr.start_col), ws_src.Cells(tr.end_row, tr.end_col))
                    src_rng.Copy(ws_out.Cells(dest_row, 1))

                    # Copy row heights for the tag rows.
                    for off in range(_TAG_HEIGHT_ROWS):
                        try:
                            ws_out.Rows(dest_row + off).RowHeight = ws_src.Rows(tr.start_row + off).RowHeight
                        except Exception:
                            pass

                    # One blank separator row between tags.
                    sep_src_row = tr.end_row + 1
                    sep_dest_row = dest_row + _TAG_HEIGHT_ROWS
                    try:
                        sep_rng = ws_src.Range(ws_src.Cells(sep_src_row, tr.start_col), ws_src.Cells(sep_src_row, tr.end_col))
                        sep_rng.Copy(ws_out.Cells(sep_dest_row, 1))
                        try:
                            ws_out.Rows(sep_dest_row).RowHeight = ws_src.Rows(sep_src_row).RowHeight
                        except Exception:
                            pass
                    except Exception:
                        # If source has no separator row, just leave it empty.
                        pass

                    dest_row = sep_dest_row + 1

            finally:
                try:
                    wb_src.Close(False)
                except Exception:
                    pass

        if not header_copied:
            raise ValueError("Не удалось найти ни одного ценника (маркер 'Цена за порц.').")

        # Save to temp first; some folders (e.g., Desktop) can be blocked for Excel COM SaveAs.
        try:
            wb_out.SaveAs(str(tmp_path), FileFormat=file_format)
        finally:
            try:
                wb_out.Close(False)
            except Exception:
                pass

    finally:
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass

    # Move to requested destination (overwrite if exists)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        if out_path.exists():
            try:
                out_path.unlink()
            except Exception:
                # If file is open, user must close it.
                raise
        os.replace(str(tmp_path), str(out_path))
    except Exception:
        # fallback: move (may work across volumes)
        try:
            shutil.move(str(tmp_path), str(out_path))
        except Exception:
            raise
