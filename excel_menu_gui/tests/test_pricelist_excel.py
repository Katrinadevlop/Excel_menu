#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tempfile
import unittest
from pathlib import Path

import openpyxl

from app.reports.pricelist_excel import create_pricelist_xlsx
from app.services.dish_extractor import DishItem


class TestPricelistExcel(unittest.TestCase):
    def test_creates_merged_name_row_for_each_item(self):
        dishes = [
            DishItem(name="Борщ", weight="250 г", price="120"),
            DishItem(name="Цезарь", weight="200 г", price="275"),
        ]

        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "ценники.xlsx"
            create_pricelist_xlsx(dishes, str(out))
            self.assertTrue(out.exists())

            wb = openpyxl.load_workbook(str(out))
            ws = wb.active

            merged = {str(rng) for rng in ws.merged_cells.ranges}
            # Первый ценник: A1:C1
            self.assertIn("A1:C1", merged)
            # Второй ценник: у нас block_height=4 + 1 пустая строка => старт 1 + 5 = 6
            self.assertIn("A6:C6", merged)


if __name__ == '__main__':
    unittest.main(verbosity=2)
