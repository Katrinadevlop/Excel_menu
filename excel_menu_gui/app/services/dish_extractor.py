#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Общий модуль для извлечения блюд из различных источников данных.
Централизованная логика для работы с меню в разных форматах.
"""

import pandas as pd
import openpyxl
import re
from pathlib import Path
from datetime import datetime, date
from typing import Dict, List, Optional, Tuple, Union
from dataclasses import dataclass
from abc import ABC, abstractmethod


# Исключения
class DishExtractionError(Exception):
    """Базовая ошибка извлечения блюд"""
    pass


class FileFormatError(DishExtractionError):
    """Ошибка неподдерживаемого формата файла"""
    pass


class SheetNotFoundError(DishExtractionError):
    """Ошибка отсутствия листа в файле"""
    pass


# Модели данных
@dataclass
class DishItem:
    """Представляет отдельное блюдо с базовой информацией.
    
    Данный класс содержит основную информацию о блюде:
    название, вес, цену и категорию. Автоматически очищает
    и нормализует данные при создании объекта.
    
    Attributes:
        name (str): Название блюда
        weight (str): Вес или порция блюда (например, "150г", "1 шт")
        price (str): Цена блюда (например, "120 руб")
        category (str): Категория блюда (завтрак, салат, первое, мясо, курица, рыба, гарнир)
    """
    name: str
    weight: str = ""
    price: str = ""
    category: str = ""
    
    def __post_init__(self):
        # Очистка и нормализация данных при создании
        self.name = str(self.name).strip()
        self.weight = str(self.weight).strip() if self.weight else ""
        self.price = str(self.price).strip() if self.price else ""
        self.category = str(self.category).strip() if self.category else ""


@dataclass
class ExtractionResult:
    """Результат извлечения блюд с метаданными.
    
    Содержит результаты извлечения блюд из источника данных,
    включая общий список, группировку по категориям и дополнительную информацию.
    
    Attributes:
        dishes (List[DishItem]): Список всех извлеченных блюд
        categories (Dict[str, List[DishItem]]): Блюда, сгруппированные по категориям
        source_date (Optional[datetime]): Дата, найденная в источнике данных
        total_count (int): Общее количество блюд (вычисляется автоматически)
    """
    dishes: List[DishItem]
    categories: Dict[str, List[DishItem]]
    source_date: Optional[datetime] = None
    total_count: int = 0
    
    def __post_init__(self):
        self.total_count = len(self.dishes)


# Абстракция для источников данных
class DataSource(ABC):
    """Абстрактный базовый класс для источников данных о блюдах.
    
    Определяет интерфейс для всех типов источников данных (Excel, CSV, базы данных).
    Наследники должны реализовать методы для извлечения блюд и дат.
    """
    
    @abstractmethod
    def extract_dishes(self, source_path: str, **kwargs) -> ExtractionResult:
        """Извлекает блюда из источника данных.
        
        Args:
            source_path: Путь к источнику данных
            **kwargs: Дополнительные параметры извлечения
            
        Returns:
            ExtractionResult: Результат извлечения с блюдами и метаданными
            
        Raises:
            DishExtractionError: При ошибках извлечения данных
        """
        pass
    
    @abstractmethod
    def extract_date(self, source_path: str, **kwargs) -> Optional[datetime]:
        """Извлекает дату из источника данных.
        
        Args:
            source_path: Путь к источнику данных
            **kwargs: Дополнительные параметры извлечения
            
        Returns:
            Optional[datetime]: Найденная дата или None
        """
        pass


class ExcelDataSource(DataSource):
    """Источник данных для работы с Excel файлами меню.
    
    Обрабатывает файлы форматов .xls, .xlsx, .xlsm.
    Извлекает блюда по категориям и находит даты в содержимом файлов.
    Поддерживает различные структуры Excel-документов.
    """
    
    def __init__(self):
        self.categories_mapping = {
            'завтрак': 'завтрак',
            'салат': 'салат', 
            'первое': 'первое',
            'мясо': 'мясо',
            'курица': 'курица',
            'птица': 'птица',
            'рыба': 'рыба',
            'гарнир': 'гарнир'
        }
    
    def extract_dishes(self, source_path: str, **kwargs) -> ExtractionResult:
        """Извлекает все блюда из Excel файла"""
        if not Path(source_path).exists():
            raise FileFormatError(f"Файл не найден: {source_path}")
        
        # Определяем формат файла
        ext = Path(source_path).suffix.lower()
        if ext not in ['.xls', '.xlsx', '.xlsm']:
            raise FileFormatError(f"Неподдерживаемый формат файла: {ext}")
        
        # Извлекаем блюда по категориям
        categorized_dishes = self._extract_categorized_dishes(source_path)
        
        # Преобразуем в список DishItem
        all_dishes = []
        categories_result = {}
        
        for category, dish_names in categorized_dishes.items():
            category_dishes = []
            for dish_name in dish_names:
                dish_item = DishItem(name=dish_name, category=category)
                all_dishes.append(dish_item)
                category_dishes.append(dish_item)
            categories_result[category] = category_dishes
        
        # Извлекаем дату
        source_date = self.extract_date(source_path)
        
        return ExtractionResult(
            dishes=all_dishes,
            categories=categories_result,
            source_date=source_date
        )
    
    def extract_date(self, source_path: str, **kwargs) -> Optional[datetime]:
        """Извлекает дату из Excel файла"""
        try:
            if source_path.endswith('.xls'):
                return self._extract_date_from_xls(source_path)
            else:
                return self._extract_date_from_xlsx(source_path)
        except Exception as e:
            print(f"Ошибка при извлечении даты: {e}")
            return None
    
    def _extract_date_from_xlsx(self, path: str) -> Optional[datetime]:
        """Извлекает дату из xlsx файла"""
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            
            # Проверяем название файла
            filename = Path(path).stem
            date_from_filename = self._parse_date_string(filename)
            if date_from_filename:
                return date_from_filename
            
            # Ищем дату в листах
            for ws in wb.worksheets:
                # Проверяем название листа
                date_from_sheet = self._parse_date_string(ws.title)
                if date_from_sheet:
                    return date_from_sheet
                
                # Ищем дату в первых 20 строках
                for row in range(1, min(21, ws.max_row + 1)):
                    for col in range(1, min(ws.max_column + 1, 10)):
                        cell = ws.cell(row=row, column=col)
                        if cell.value:
                            date_from_cell = self._parse_date_string(str(cell.value))
                            if date_from_cell:
                                return date_from_cell
        except Exception:
            pass
        return None
    
    def _extract_date_from_xls(self, path: str) -> Optional[datetime]:
        """Извлекает дату из xls файла"""
        try:
            df_dict = pd.read_excel(path, sheet_name=None)
            
            for sheet_name, df in df_dict.items():
                # Проверяем название листа
                date_from_sheet = self._parse_date_string(sheet_name)
                if date_from_sheet:
                    return date_from_sheet
                
                # Ищем дату в содержимом листа
                for col in df.columns:
                    if pd.notna(col):
                        date_from_col = self._parse_date_string(str(col))
                        if date_from_col:
                            return date_from_col
                
                # Ищем дату в первых строках
                for _, row in df.head(10).iterrows():
                    for cell in row:
                        if pd.notna(cell):
                            date_from_cell = self._parse_date_string(str(cell))
                            if date_from_cell:
                                return date_from_cell
        except Exception:
            pass
        return None
    
    def _parse_date_string(self, text: str) -> Optional[datetime]:
        """Парсит дату из строки"""
        if not text:
            return None
            
        text = text.strip().lower()
        
        # Русские месяцы
        months_ru = {
            'января': 1, 'февраля': 2, 'марта': 3, 'апреля': 4, 'мая': 5, 'июня': 6,
            'июля': 7, 'августа': 8, 'сентября': 9, 'октября': 10, 'ноября': 11, 'декабря': 12
        }
        
        # Паттерны для поиска даты
        patterns = [
            r'(\d{1,2})\s+(\w+)',  # "5 сентября"
            r'(\d{1,2})\.(\d{1,2})\.(\d{4})',  # "05.09.2024"
            r'(\d{1,2})/(\d{1,2})/(\d{4})',  # "05/09/2024"
            r'(\d{1,2})\s+(\w+)\s+(\d{4})',  # "5 сентября 2024"
            r'(\d{2})\.(\d{2})\.(\d{2})',  # "05.09.24"
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                try:
                    if len(match.groups()) == 2:
                        # Формат "день месяц"
                        day = int(match.group(1))
                        month_str = match.group(2)
                        month = months_ru.get(month_str)
                        if month:
                            year = datetime.now().year
                            return datetime(year, month, day)
                    elif len(match.groups()) == 3:
                        if match.group(2) in months_ru:
                            # Формат "день месяц год"
                            day = int(match.group(1))
                            month = months_ru[match.group(2)]
                            year = int(match.group(3))
                        else:
                            # Формат "день.месяц.год"
                            day = int(match.group(1))
                            month = int(match.group(2))
                            year = int(match.group(3))
                            if year < 100:  # Двузначный год
                                year += 2000
                        return datetime(year, month, day)
                except (ValueError, TypeError):
                    continue
        
        return None
    
    def _extract_categorized_dishes(self, source_path: str) -> Dict[str, List[str]]:
        """Извлекает блюда по категориям"""
        result: Dict[str, List[str]] = {k: [] for k in self.categories_mapping.keys()}
        
        try:
            if source_path.endswith('.xls'):
                self._extract_from_xls(source_path, result)
            else:
                self._extract_from_xlsx(source_path, result)
        except Exception as e:
            print(f"Ошибка при извлечении категорий: {e}")
        
        return result
    
    def _extract_from_xlsx(self, path: str, result: Dict[str, List[str]]):
        """Извлекает блюда из xlsx файла"""
        wb = openpyxl.load_workbook(path, data_only=True)
        if wb.worksheets:
            self._extract_from_worksheet(wb.worksheets[0], result)
    
    def _extract_from_xls(self, path: str, result: Dict[str, List[str]]):
        """Извлекает блюда из xls файла"""
        df_dict = pd.read_excel(path, sheet_name=None)
        
        # Ищем лист с приоритетом
        preferred = None
        for name, df in df_dict.items():
            nm = str(name).lower()
            if 'касс' in nm:
                preferred = df
                break
        
        if preferred is None:
            for name, df in df_dict.items():
                nm = str(name).lower()
                if 'меню' in nm:
                    preferred = df
                    break
        
        if preferred is None and df_dict:
            preferred = list(df_dict.values())[0]
        
        if preferred is not None:
            self._extract_from_dataframe(preferred, result)
    
    def _extract_from_worksheet(self, ws, result: Dict[str, List[str]]):
        """Извлекает блюда из worksheet"""
        # Находим строку заголовков
        header_row = None
        for row_idx in range(1, min(11, ws.max_row + 1)):
            row_text = ''
            for col_idx in range(1, ws.max_column + 1):
                cell_val = ws.cell(row=row_idx, column=col_idx).value
                if cell_val:
                    row_text += str(cell_val).strip() + ' '
            
            if 'ЗАВТРАКИ' in row_text.upper():
                header_row = row_idx
                break
        
        if header_row is None:
            return
        
        # Извлекаем завтраки и салаты из первого столбца
        current_category = 'завтрак'
        
        for row_idx in range(header_row + 1, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=1).value  # Столбец A
            if cell_value:
                cell_str = str(cell_value).strip()
                
                # Проверяем заголовки
                if 'САЛАТ' in cell_str.upper() and 'ХОЛОДН' in cell_str.upper():
                    current_category = 'салат'
                    continue
                
                # Добавляем блюдо
                if not self._should_skip_cell(cell_str) and self._is_valid_dish(cell_str, []):
                    result[current_category].append(cell_str)
        
        # Извлекаем остальные блюда из правых столбцов
        current_category = None
        for col_idx in range(4, ws.max_column + 1):
            for row_idx in range(header_row, ws.max_row + 1):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value:
                    cell_str = str(cell_value).strip()
                    
                    # Определяем категорию
                    if 'ПЕРВЫЕ' in cell_str.upper() and 'БЛЮДА' in cell_str.upper():
                        current_category = 'первое'
                    elif 'БЛЮДА ИЗ МЯСА' in cell_str.upper():
                        current_category = 'мясо'
                    elif 'БЛЮДА ИЗ ПТИЦЫ' in cell_str.upper():
                        current_category = 'птица'
                    elif 'БЛЮДА ИЗ РЫБЫ' in cell_str.upper():
                        current_category = 'рыба'
                    elif 'ГАРНИРЫ' in cell_str.upper():
                        current_category = 'гарнир'
                    elif 'НАПИТКИ' in cell_str.upper():
                        return  # Останавливаемся на напитках
                    elif current_category and not self._should_skip_cell(cell_str) and self._is_valid_dish(cell_str, []):
                        result[current_category].append(cell_str)
    
    def _extract_from_dataframe(self, df: pd.DataFrame, result: Dict[str, List[str]]):
        """Извлекает блюда из DataFrame"""
        # Находим строку заголовков
        header_row = None
        for idx in range(min(10, len(df))):
            row = df.iloc[idx]
            row_text = ' '.join([str(cell).strip() for cell in row if pd.notna(cell) and str(cell).strip()])
            if 'ЗАВТРАКИ' in row_text.upper():
                header_row = idx
                break
        
        if header_row is None:
            return
        
        # Извлекаем завтраки из первого столбца
        for row_idx in range(header_row + 1, len(df)):
            cell_value = df.iloc[row_idx, 0]  # Столбец A
            if pd.notna(cell_value):
                cell_str = str(cell_value).strip()
                if not self._should_skip_cell(cell_str) and self._is_valid_dish(cell_str, []):
                    result['завтрак'].append(cell_str)
            elif len(result['завтрак']) > 0:
                break
        
        # Извлекаем остальные блюда
        current_category = None
        for col_idx in range(3, len(df.columns)):
            for row_idx in range(header_row, len(df)):
                cell_value = df.iloc[row_idx, col_idx]
                if pd.notna(cell_value):
                    cell_str = str(cell_value).strip()
                    
                    # Определяем категорию
                    if 'ПЕРВЫЕ' in cell_str.upper() and 'БЛЮДА' in cell_str.upper():
                        current_category = 'первое'
                    elif 'БЛЮДА ИЗ МЯСА' in cell_str.upper():
                        current_category = 'мясо'
                    elif 'БЛЮДА ИЗ ПТИЦЫ' in cell_str.upper():
                        current_category = 'птица'
                    elif 'БЛЮДА ИЗ РЫБЫ' in cell_str.upper():
                        current_category = 'рыба'
                    elif 'ГАРНИРЫ' in cell_str.upper():
                        current_category = 'гарнир'
                    elif 'НАПИТКИ' in cell_str.upper():
                        return
                    elif current_category and not self._should_skip_cell(cell_str) and self._is_valid_dish(cell_str, []):
                        result[current_category].append(cell_str)
    
    def _should_skip_cell(self, cell_str: str) -> bool:
        """Проверяет, нужно ли пропустить ячейку"""
        cell_lower = cell_str.lower().strip()
        
        if len(cell_str) < 4:
            return True
        
        skip_words = [
            'вес', 'цена', 'руб', 'ед.изм', 'утверждаю', 'директор', 
            'меню', 'столовой', 'патриот', 'москва', 'наб', 'стр',
            'попова', 'сентября', 'пятница', '_____', 'понедельник', 'вторник',
            'среда', 'четверг', 'пятница', 'суббота', 'воскресенье',
            'январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль',
            'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь',
            'завтрак', 'обед', 'ужин', 'полдник', 'время', 'дата',
            'напит', 'сок', 'чай', 'кофе', 'смузи', 'фреш',
            'соус', 'майонез', 'кетчуп',
            'наименование', 'блюда', 'час', 'мин'
        ]
        
        # Проверяем время
        if re.match(r'^\d{1,2}:\d{2}(:\d{2})?$', cell_str):
            return True
        
        for skip_word in skip_words:
            if skip_word in cell_lower:
                return True
        
        # Пропускаем объемы напитков
        if re.search(r'\d+/\d+\s*мл|\d+\s*мл|200/300мл|300мл|225\s*мл', cell_str):
            return True
        
        # Пропускаем числа и символы
        if re.match(r'^[\d\s\.,/г-]+$', cell_str):
            return True
        
        # Пропускаем адреса
        if 'ул.' in cell_lower or 'д.' in cell_lower or 'овчинниковская' in cell_lower:
            return True
            
        return False
    
    def _is_valid_dish(self, dish_name: str, existing_dishes: List[str]) -> bool:
        """Проверяет, является ли строка валидным названием блюда"""
        if dish_name in existing_dishes:
            return False
        
        if len(dish_name) < 5:
            return False
        
        # Проверяем заголовки категорий
        dish_lower = dish_name.lower().strip()
        category_titles = [
            'завтрак', 'салат', 'холодн', 'закуск', 
            'первое', 'первы', 'блюд', 'гарнир', 
            'мясо', 'курица', 'птица', 'рыба'
        ]
        
        if len(dish_name) < 35:
            for title in category_titles:
                if dish_lower == title or dish_lower.startswith(title + ' ') or dish_lower.endswith(' ' + title):
                    return False
        
        # Только числа и символы
        if re.match(r'^[\d\s\.,/г-]+$', dish_name):
            return False
            
        return True


class DishExtractorService:
    """Основной сервис для извлечения блюд"""
    
    def __init__(self):
        self._sources = {
            'excel': ExcelDataSource(),
        }
    
    def extract_dishes(self, source_path: str, source_type: str = 'auto', **kwargs) -> ExtractionResult:
        """
        Извлекает блюда из источника данных
        
        Args:
            source_path: Путь к источнику данных
            source_type: Тип источника ('excel', 'auto')
            **kwargs: Дополнительные параметры
        
        Returns:
            ExtractionResult: Результат извлечения
        """
        if source_type == 'auto':
            source_type = self._detect_source_type(source_path)
        
        if source_type not in self._sources:
            raise FileFormatError(f"Неподдерживаемый тип источника: {source_type}")
        
        source = self._sources[source_type]
        return source.extract_dishes(source_path, **kwargs)
    
    def extract_categorized_dishes(self, source_path: str, source_type: str = 'auto', **kwargs) -> Dict[str, List[str]]:
        """
        Извлекает блюда, сгруппированные по категориям
        
        Args:
            source_path: Путь к источнику данных
            source_type: Тип источника
            **kwargs: Дополнительные параметры
        
        Returns:
            Dict[str, List[str]]: Словарь категория -> список блюд
        """
        result = self.extract_dishes(source_path, source_type, **kwargs)
        return {category: [dish.name for dish in dishes] for category, dishes in result.categories.items()}
    
    def extract_dishes_by_category(self, source_path: str, category: str, source_type: str = 'auto', **kwargs) -> List[str]:
        """
        Извлекает блюда определенной категории
        
        Args:
            source_path: Путь к источнику данных
            category: Категория блюд
            source_type: Тип источника
            **kwargs: Дополнительные параметры
        
        Returns:
            List[str]: Список названий блюд
        """
        categorized = self.extract_categorized_dishes(source_path, source_type, **kwargs)
        return categorized.get(category, [])
    
    def extract_date_from_source(self, source_path: str, source_type: str = 'auto', **kwargs) -> Optional[datetime]:
        """
        Извлекает дату из источника данных
        
        Args:
            source_path: Путь к источнику данных
            source_type: Тип источника
            **kwargs: Дополнительные параметры
        
        Returns:
            Optional[datetime]: Дата из источника или None
        """
        if source_type == 'auto':
            source_type = self._detect_source_type(source_path)
        
        if source_type not in self._sources:
            raise FileFormatError(f"Неподдерживаемый тип источника: {source_type}")
        
        source = self._sources[source_type]
        return source.extract_date(source_path, **kwargs)
    
    def _detect_source_type(self, source_path: str) -> str:
        """Определяет тип источника данных по расширению файла"""
        ext = Path(source_path).suffix.lower()
        if ext in ['.xls', '.xlsx', '.xlsm']:
            return 'excel'
        else:
            raise FileFormatError(f"Не удалось определить тип источника для файла: {source_path}")


# Глобальный экземпляр сервиса для удобства использования
_extractor_service = DishExtractorService()


# Публичные функции для совместимости с существующим кодом
def extract_categorized_dishes_from_menu(menu_path: str) -> Dict[str, List[str]]:
    """Извлекает блюда по категориям из файла меню (совместимость)"""
    return _extractor_service.extract_categorized_dishes(menu_path)


def extract_dishes_by_category(menu_path: str, category: str) -> List[str]:
    """Извлекает блюда определенной категории (совместимость)"""
    return _extractor_service.extract_dishes_by_category(menu_path, category)


def extract_date_from_menu(menu_path: str) -> Optional[datetime]:
    """Извлекает дату из файла меню (совместимость)"""
    return _extractor_service.extract_date_from_source(menu_path)


def get_dish_extractor() -> DishExtractorService:
    """Возвращает экземпляр сервиса извлечения блюд"""
    return _extractor_service
