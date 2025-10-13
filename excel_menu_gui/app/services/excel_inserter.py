#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Общий модуль для вставки блюд в Excel-шаблоны.
Содержит ядро записи строк блюд (название/вес/цена) и заготовки стратегий
для расширения логики вставки под разные цели (шаблон меню, бракеражный журнал и др.).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional, Dict, Iterable

import logging
from app.services.dish_extractor import DishItem


@dataclass
class TargetColumns:
    """
    Описание целевых колонок для вставки блюда.

    Attributes:
        name_col (int): Номер колонки для названия блюда (1-базовый индекс).
        weight_col (Optional[int]): Колонка для веса/выхода. Если None — не заполняется.
        price_col (Optional[int]): Колонка для цены. Если None — не заполняется.
    """
    name_col: int
    weight_col: Optional[int] = None
    price_col: Optional[int] = None


def fill_cells_sequential(ws, start_row: int, stop_row: int, columns: TargetColumns,
                          dishes: List[DishItem], replace_only_empty: bool = True,
                          logger: Optional[logging.Logger] = None, log_context: str = "") -> int:
    """
    Последовательно заполняет строки Excel блюдами в заданных колонках.

    Args:
        ws: Рабочий лист openpyxl.
        start_row (int): Первая строка для вставки (включительно), 1-базовая.
        stop_row (int): Первая строка после последней допустимой строки (исключая), 1-базовая.
        columns (TargetColumns): Колонки для названия/веса/цены.
        dishes (List[DishItem]): Блюда (name обязательно, weight/price — опционально).
        replace_only_empty (bool): Заполнять только пустые ячейки (по умолчанию True). Если False — перезаписывает.
        logger (Optional[logging.Logger]): Логгер для записи прогресса (INFO — сводка, DEBUG — построчно).
        log_context (str): Доп. контекст для логов (например, категория/диапазон).

    Returns:
        int: Количество вставленных строк блюд.
    """
    inserted = 0
    if start_row >= stop_row:
        return 0

    if logger:
        rng = f"{start_row}..{stop_row-1}"
        cols = f"{columns.name_col}/{columns.weight_col}/{columns.price_col}"
        ctx = f" [{log_context}]" if log_context else ""
        logger.info(f"Пишем {len(dishes)} позиций в строки {rng}, колонки {cols}{ctx}")

    row = start_row
    dish_idx = 0
    while row < stop_row and dish_idx < len(dishes):
        item = dishes[dish_idx]

        # Заполнение названия
        cell_name = ws.cell(row=row, column=columns.name_col)
        can_write_name = (not replace_only_empty) or (cell_name.value in (None, ""))
        wrote_any = False
        if can_write_name and item.name:
            cell_name.value = item.name
            wrote_any = True
        elif logger and logger.isEnabledFor(logging.DEBUG):
            logger.debug(f"Строка {row}: ячейка названия занята, пропуск (value={cell_name.value!r})")

        # Заполнение веса/выхода
        if columns.weight_col:
            cell_weight = ws.cell(row=row, column=columns.weight_col)
            can_write_weight = (not replace_only_empty) or (cell_weight.value in (None, ""))
            if can_write_weight and item.weight:
                cell_weight.value = item.weight
                wrote_any = True

        # Заполнение цены
        if columns.price_col:
            cell_price = ws.cell(row=row, column=columns.price_col)
            can_write_price = (not replace_only_empty) or (cell_price.value in (None, ""))
            if can_write_price and item.price:
                cell_price.value = item.price
                wrote_any = True

        if wrote_any:
            if logger and logger.isEnabledFor(logging.DEBUG):
                logger.debug(f"Строка {row}: вставлено name={item.name!r}, weight={item.weight!r}, price={item.price!r}")
            inserted += 1
            dish_idx += 1
        # Если строка занята и ничего не записали, просто сдвигаемся вниз
        row += 1

    if logger:
        rng = f"{start_row}..{stop_row-1}"
        ctx = f" [{log_context}]" if log_context else ""
        logger.info(f"Итого вставлено {inserted} строк в диапазон {rng}{ctx}")

    return inserted


# Ниже — заготовки стратегий для дальнейшего расширения
from abc import ABC, abstractmethod


class InsertionPolicy(ABC):
    """
    Базовый интерфейс стратегии вставки блюд в Excel.
    Определяет контракт планирования вставки под конкретную цель/шаблон.
    """

    @abstractmethod
    def plan(self, ws, categorized: Dict[str, List[DishItem]]):
        """
        Возвращает набор задач на вставку для листа.

        Args:
            ws: Рабочий лист openpyxl.
            categorized (dict[str, list[DishItem]]): Блюда по категориям.

        Returns:
            Iterable[tuple[start_row, stop_row, TargetColumns, List[DishItem]]]
        """
        raise NotImplementedError


class TemplateInsertionPolicy(InsertionPolicy):
    """
    Заготовка стратегии для шаблона меню (вставка по категориям в предусмотренные области).
    Конкретная логика поиска диапазонов и колонок должна быть добавлена при необходимости.
    """

    def plan(self, ws, categorized: Dict[str, List[DishItem]]):  # pragma: no cover (заготовка)
        return []


class JournalInsertionPolicy(InsertionPolicy):
    """
    Заготовка стратегии для бракеражного журнала (две колонки: A — завтраки/салаты, G — остальное).
    Конкретная логика поиска диапазонов и колонок должна быть добавлена при необходимости.
    """

    def plan(self, ws, categorized: Dict[str, List[DishItem]]):  # pragma: no cover (заготовка)
        return []
