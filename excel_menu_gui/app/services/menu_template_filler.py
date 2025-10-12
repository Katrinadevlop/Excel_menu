#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Модуль для заполнения шаблона меню данными из другого файла меню
"""

import pandas as pd
import openpyxl
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from datetime import datetime
import re

from dish_extractor import extract_categorized_dishes_from_menu, extract_date_from_menu
from comparator import _find_category_ranges, _extract_dishes_from_multiple_columns, read_cell_values, normalize_dish


class MenuTemplateFiller:
    """Заполняет шаблон меню данными из другого файла"""
    
    def __init__(self):
        self.categories_mapping = {
            'завтрак': 'завтраки',
            'салат': 'холодные закуски и салаты', 
            'первое': 'первые блюда',
            'мясо': 'блюда из мяса',
            'курица': 'блюда из курицы',
            'птица': 'блюда из птицы',  # Добавляем птицу
            'рыба': 'блюда из рыбы',
            'гарнир': 'гарниры'
        }
    
    def extract_categorized_dishes(self, menu_path: str) -> Dict[str, List[str]]:
        """Извлекает блюда по категориям из файла меню (используем уже готовую логику)"""
        return extract_categorized_dishes_from_menu(menu_path)
    
    def extract_categorized_dishes_advanced(self, menu_path: str) -> Dict[str, List[str]]:
        """Улучшенное извлечение блюд по категориям используя логику из comparator.py"""
        try:
            # Определяем какой лист использовать
            from comparator import get_sheet_names
            sheets = get_sheet_names(menu_path)
            if not sheets:
                return {}
            
            # Ищем лист с приоритетом "касс" или первый доступный
            sheet_name = None
            for name in sheets:
                if 'касс' in name.lower():
                    sheet_name = name
                    break
            if not sheet_name:
                sheet_name = sheets[0]
            
            # Читаем данные из файла
            values = read_cell_values(menu_path, sheet_name)
            
            # Определяем синонимы для поиска категорий (как в comparator.py)
            synonyms_map = {
                'завтрак': 'завтраки',
                'салат': 'салаты и холодные закуски', 
                'холодн': 'салаты и холодные закуски',
                'закуск': 'салаты и холодные закуски',
                'перв': 'первые блюда',
                'мяс': 'блюда из мяса',
                'птиц': 'блюда из птицы',
                'курин': 'блюда из птицы',
                'рыб': 'блюда из рыбы',
                'гарнир': 'гарниры'
            }
            
            # Находим диапазоны категорий
            ranges = _find_category_ranges(values, synonyms_map)
            
            # Извлекаем блюда из каждой категории
            result = {
                'завтрак': [],
                'салат': [], 
                'первое': [],
                'мясо': [],
                'курица': [],
                'птица': [],
                'рыба': [],
                'гарнир': []
            }
            
            # Мапинг категорий из comparator в наши категории
            category_mapping = {
                'завтраки': 'завтрак',
                'салаты и холодные закуски': 'салат',
                'первые блюда': 'первое',
                'блюда из мяса': 'мясо',
                'блюда из птицы': 'птица', 
                'блюда из рыбы': 'рыба',
                'гарниры': 'гарнир'
            }
            
            # Извлекаем блюда для каждой найденной категории
            for comp_category, (start, end) in ranges.items():
                our_category = category_mapping.get(comp_category)
                if our_category:
                    # Извлекаем блюда из столбцов A и D для этого диапазона
                    dishes_set = _extract_dishes_from_multiple_columns(values, start, end, True, ['A', 'D', 'E'])
                    
                    # Фильтруем и очищаем блюда
                    clean_dishes = []
                    for dish in dishes_set:
                        if dish and len(dish.strip()) > 3:
                            # Пропускаем заголовки
                            dish_lower = dish.lower()
                            if not any(header in dish_lower for header in ['блюда', 'салаты', 'гарниры', 'первые', 'вес', 'цена', 'руб']):
                                clean_dishes.append(dish.strip())
                    
                    result[our_category] = clean_dishes
                    print(f"Категория {comp_category} -> {our_category}: найдено {len(clean_dishes)} блюд")
            
            return result
            
        except Exception as e:
            print(f"Ошибка в продвинутом извлечении: {e}")
            # Fallback к старому методу
            return self.extract_categorized_dishes(menu_path)
    
    def extract_dishes_with_details(self, menu_path: str) -> Dict[str, List[Dict[str, str]]]:
        """Извлекает блюда с деталями (название, вес, цена) из файла меню"""
        try:
            import openpyxl
            wb = openpyxl.load_workbook(menu_path)
            ws = wb.active
            
            result = {
                'завтрак': [],
                'салат': [], 
                'первое': [],
                'мясо': [],
                'курица': [],
                'птица': [],
                'рыба': [],
                'гарнир': []
            }
            
            # Ищем категории и извлекаем блюда с деталями
            current_category = None
            category_mapping = {
                'завтрак': 'завтрак',
                'салат': 'салат',
                'закуск': 'салат',
                'холодн': 'салат',
                'перв': 'первое',
                'мяс': 'мясо',
                'птиц': 'птица',
                'курин': 'птица', 
                'рыб': 'рыба',
                'гарнир': 'гарнир'
            }
            
            for row in range(1, min(100, ws.max_row + 1)):
                # Проверяем заголовки категорий
                cell_val = ws.cell(row=row, column=1).value
                if cell_val:
                    cell_text = str(cell_val).lower().strip()
                    for key, category in category_mapping.items():
                        if key in cell_text:
                            current_category = category
                            print(f"Найдена категория {current_category} в строке {row}: {cell_val}")
                            break
                
                # Если нашли категорию, собираем блюда
                if current_category and cell_val and current_category in result:
                    # Проверяем, не является ли это заголовком
                    if not any(header in cell_text for header in ['блюда', 'салаты', 'гарниры', 'первые', 'вес', 'цена', 'руб']):
                        # Получаем вес и цену из соседних колонок
                        weight = ws.cell(row=row, column=2).value
                        price = ws.cell(row=row, column=3).value
                        
                        dish_info = {
                            'name': str(cell_val).strip(),
                            'weight': str(weight).strip() if weight else '',
                            'price': str(price).strip() if price else ''
                        }
                        
                        result[current_category].append(dish_info)
                        print(f"Добавлено блюдо в {current_category}: {dish_info}")
            
            wb.close()
            
            # Отчет о найденных блюдах
            total = sum(len(dishes) for dishes in result.values())
            print(f"\nИзвлечено блюд с деталями:")
            for cat, dishes in result.items():
                if dishes:
                    print(f"  {cat}: {len(dishes)} блюд")
            print(f"Всего блюд: {total}")
            
            return result
            
        except Exception as e:
            print(f"Ошибка при извлечении блюд с деталями: {e}")
            return {}
    
    def find_column_by_header(self, ws, header_text: str) -> Optional[int]:
        """Находит номер колонки по заголовку"""
        header_variations = {
            'завтраки': ['завтрак', 'завтраки'],
            'холодные закуски и салаты': ['салат', 'холодн', 'закуск'],  # Работает с "САЛАТЫ и ХОЛОДНЫЕ ЗАКУСКИ"
            'первые блюда': ['первые', 'первое'],
            'блюда из мяса': ['мяса', 'мясо', 'мясн'],  # Работает с "БЛЮДА ИЗ МЯСА"
            'блюда из курицы': ['курица', 'курин', 'птица', 'птицы'],
            'блюда из птицы': ['птицы', 'птица', 'курица', 'курин'],  # Работает с "БЛЮДА ИЗ ПТИЦЫ"
            'блюда из рыбы': ['рыбы', 'рыба', 'рыбн'],  # Работает с "БЛЮДА ИЗ РЫБЫ"
            'гарниры': ['гарниры']
        }
        
        # Ищем в первых 50 строках (некоторые заголовки могут быть на 30-й строке)
        for row in range(1, min(51, ws.max_row + 1)):
            for col in range(1, ws.max_column + 1):
                cell_val = ws.cell(row=row, column=col).value
                if cell_val:
                    cell_text = str(cell_val).lower().strip()
                    
                    # Проверяем точное совпадение
                    if header_text.lower() in cell_text:
                        return col
                    
                    # Проверяем вариации
                    if header_text.lower() in header_variations:
                        for variation in header_variations[header_text.lower()]:
                            if variation in cell_text:
                                return col
        
        return None
    
    def find_data_start_row(self, ws, col: int, category: str = None) -> int:
        """Находит строку начала данных в колонке (после заголовков)"""
        # Специальная логика для гарниров - ищем именно заголовок "ГАРНИРЫ"
        if category == 'гарнир':
            for row in range(20, min(50, ws.max_row + 1)):  # Ищем с 20 по 50 строку
                cell_val = ws.cell(row=row, column=col).value
                if cell_val:
                    cell_text = str(cell_val).lower().strip()
                    if 'гарнир' in cell_text and ('гарниры' in cell_text or 'гарнир' == cell_text):
                        print(f"Найден заголовок гарниров в строке {row}: {cell_val}")
                        return row + 1  # Начинаем с следующей строки после заголовка
        
        # Для остальных категорий - находим строку с соответствующим заголовком
        header_keywords = {
            'завтрак': ['завтрак'],
            'салат': ['салат', 'закуск'],
            'первое': ['первые'],
            'мясо': ['мяса', 'мясо'],
            'курица': ['курица', 'птица'],
            'птица': ['птица', 'курица'],
            'рыба': ['рыба', 'рыбы']
        }
        
        if category in header_keywords:
            keywords = header_keywords[category]
            for row in range(1, min(50, ws.max_row + 1)):
                cell_val = ws.cell(row=row, column=col).value
                if cell_val:
                    cell_text = str(cell_val).lower().strip()
                    if any(keyword in cell_text for keyword in keywords):
                        print(f"Найден заголовок {category} в строке {row}: {cell_val}")
                        return row + 1
        
        # Общий поиск заголовка категории
        header_row = None
        for row in range(1, min(50, ws.max_row + 1)):
            cell_val = ws.cell(row=row, column=col).value
            if cell_val:
                cell_text = str(cell_val).lower().strip()
                # Ищем заголовки категорий
                if any(header in cell_text for header in ['завтрак', 'салат', 'мясо', 'курица', 'птица', 'рыба', 'гарнир', 'первые', 'блюда']):
                    header_row = row
                    break
        
        # Если нашли заголовок, начинаем с следующей строки
        if header_row:
            return header_row + 1
        
        # Иначе ищем строку с "Вес/ед.изм." и начинаем после неё
        for row in range(1, min(50, ws.max_row + 1)):
            cell_val = ws.cell(row=row, column=col).value
            if cell_val and 'вес' in str(cell_val).lower():
                return row + 1
                
        return 7  # По умолчанию начинаем с 7-й строки
    
    def handle_caesar_salad(self, dish_name: str) -> List[str]:
        """Обрабатывает Цезарь салат, разделяя по слэшу"""
        if 'цезарь' in dish_name.lower() and '/' in dish_name:
            parts = dish_name.split('/')
            result = []
            base = parts[0].strip()
            for i, part in enumerate(parts):
                if i == 0:
                    result.append(part.strip())
                else:
                    result.append(f"{base.split()[0]} {part.strip()}")
            return result
        return [dish_name]
    
    def find_category_end_row(self, ws, col: int, start_row: int, category: str) -> int:
        """Находит конечную строку для категории, ища следующий заголовок"""
        # Особая логика для гарниров - ищем до "НАПИТКИ"
        if category == 'гарнир':
            for row in range(start_row, min(50, ws.max_row + 1)):
                cell_val = ws.cell(row=row, column=col).value
                if cell_val and 'напит' in str(cell_val).lower():
                    return row - 1  # Возвращаем последнюю строку перед "Напитки"
        
        # Для остальных категорий ищем до следующего заголовка
        for row in range(start_row, min(50, ws.max_row + 1)):
            cell_val = ws.cell(row=row, column=col).value
            if cell_val:
                cell_text = str(cell_val).lower().strip()
                if any(header in cell_text for header in ['блюда', 'салаты', 'гарниры', 'первые', 'напит', 'хлеб']):
                    return row - 1  # Возвращаем последнюю строку перед новым заголовком
        
        # Если не нашли следующий заголовок, возвращаем конец листа
        return ws.max_row
    
    def fill_template_column(self, ws, col: int, dishes: List[str], start_row: int) -> int:
        """Заполняет колонку блюдами, начиная с указанной строки"""
        current_row = start_row
        filled_count = 0
        
        for dish in dishes:
            # Обрабатываем Цезарь салат
            dish_variants = self.handle_caesar_salad(dish)
            
            for variant in dish_variants:
                if current_row > ws.max_row:
                    break
                    
                # Получаем ячейку
                cell = ws.cell(row=current_row, column=col)
                
                # Проверяем, является ли ячейка объединенной
                if hasattr(cell, 'coordinate'):
                    # Проверяем, есть ли эта ячейка в списке объединенных ячеек
                    is_merged = False
                    for merged_range in ws.merged_cells.ranges:
                        if cell.coordinate in merged_range:
                            is_merged = True
                            break
                    
                    # Пропускаем объединенные ячейки, кроме главной
                    if is_merged and cell.__class__.__name__ == 'MergedCell':
                        current_row += 1
                        continue
                
                # Заполняем только пустые ячейки
                current_val = cell.value
                if not current_val or str(current_val).strip() == '':
                    try:
                        cell.value = variant
                        filled_count += 1
                    except AttributeError:
                        # Если не можем записать в эту ячейку, пропускаем
                        pass
                
                current_row += 1
        
        return filled_count
    
    def fill_template_column_limited(self, ws, col: int, dishes: List[str], start_row: int, end_row: int, category: str) -> int:
        """Заполняет колонку блюдами в ограниченном диапазоне строк"""
        current_row = start_row
        filled_count = 0
        
        for dish in dishes:
            # Обрабатываем Цезарь салат
            dish_variants = self.handle_caesar_salad(dish)
            
            for variant in dish_variants:
                if current_row > end_row:  # Ограничиваем по концу диапазона
                    print(f"Превышен лимит для {category}: конец на строке {end_row}")
                    break
                    
                # Получаем ячейку
                cell = ws.cell(row=current_row, column=col)
                
                # Проверяем, является ли ячейка объединенной
                if hasattr(cell, 'coordinate'):
                    # Проверяем, есть ли эта ячейка в списке объединенных ячеек
                    is_merged = False
                    for merged_range in ws.merged_cells.ranges:
                        if cell.coordinate in merged_range:
                            is_merged = True
                            break
                    
                    # Пропускаем объединенные ячейки, кроме главной
                    if is_merged and cell.__class__.__name__ == 'MergedCell':
                        current_row += 1
                        continue
                
                # Заполняем только пустые ячейки
                current_val = cell.value
                if not current_val or str(current_val).strip() == '':
                    try:
                        cell.value = variant
                        filled_count += 1
                        print(f"Заполнено {category} в строку {current_row}: {variant}")
                    except AttributeError:
                        # Если не можем записать в эту ячейку, пропускаем
                        pass
                else:
                    print(f"Ячейка {current_row}:{col} для {category} уже заполнена: {current_val}")
                
                current_row += 1
            
            # Прерываем, если превысили лимит
            if current_row > end_row:
                break
        
        return filled_count
    
    def fill_template_with_details(self, ws, name_col: int, weight_col: int, price_col: int, 
                                  dishes_with_details: List[Dict[str, str]], start_row: int) -> int:
        """Заполняет шаблон блюдами с деталями (название, вес, цена)"""
        current_row = start_row
        filled_count = 0
        
        for dish_info in dishes_with_details:
            if current_row > ws.max_row:
                break
            
            # Заполняем название
            name_cell = ws.cell(row=current_row, column=name_col)
            if not name_cell.value or str(name_cell.value).strip() == '':
                try:
                    name_cell.value = dish_info['name']
                except AttributeError:
                    pass
            
            # Заполняем вес
            if dish_info['weight'] and weight_col:
                weight_cell = ws.cell(row=current_row, column=weight_col)
                if not weight_cell.value or str(weight_cell.value).strip() == '':
                    try:
                        weight_cell.value = dish_info['weight']
                    except AttributeError:
                        pass
            
            # Заполняем цену
            if dish_info['price'] and price_col:
                price_cell = ws.cell(row=current_row, column=price_col)
                if not price_cell.value or str(price_cell.value).strip() == '':
                    try:
                        price_cell.value = dish_info['price']
                    except AttributeError:
                        pass
            
            filled_count += 1
            print(f"Заполнено в строку {current_row}: {dish_info['name']} | {dish_info['weight']} | {dish_info['price']}")
            current_row += 1
        
        return filled_count
    
    def fill_template_with_details_limited(self, ws, name_col: int, weight_col: int, price_col: int, 
                                          dishes_with_details: List[Dict[str, str]], start_row: int, end_row: int, category: str) -> int:
        """Заполняет шаблон блюдами с деталями в ограниченном диапазоне"""
        current_row = start_row
        filled_count = 0
        
        for dish_info in dishes_with_details:
            if current_row > end_row:
                print(f"Превышен лимит для {category}: конец на строке {end_row}")
                break
            
            # Заполняем название
            name_cell = ws.cell(row=current_row, column=name_col)
            if not name_cell.value or str(name_cell.value).strip() == '':
                try:
                    name_cell.value = dish_info['name']
                except AttributeError:
                    pass
            
            # Заполняем вес
            if dish_info['weight'] and weight_col:
                weight_cell = ws.cell(row=current_row, column=weight_col)
                if not weight_cell.value or str(weight_cell.value).strip() == '':
                    try:
                        weight_cell.value = dish_info['weight']
                    except AttributeError:
                        pass
            
            # Заполняем цену
            if dish_info['price'] and price_col:
                price_cell = ws.cell(row=current_row, column=price_col)
                if not price_cell.value or str(price_cell.value).strip() == '':
                    try:
                        price_cell.value = dish_info['price']
                    except AttributeError:
                        pass
            
            filled_count += 1
            print(f"Заполнено {category} в строку {current_row}: {dish_info['name']} | {dish_info['weight']} | {dish_info['price']}")
            current_row += 1
        
        return filled_count
    
    def extract_date_from_menu(self, menu_path: str) -> Optional[datetime]:
        """Извлекает дату из файла меню (используем уже готовую логику)"""
        return extract_date_from_menu(menu_path)
    
    def update_template_date(self, ws, menu_date: Optional[datetime]):
        """Обновляет дату в шаблоне"""
        if not menu_date:
            menu_date = datetime.now()
        
        # Ищем ячейки с датой в первых строках
        for row in range(1, min(6, ws.max_row + 1)):
            for col in range(1, min(5, ws.max_column + 1)):
                cell = ws.cell(row=row, column=col)
                cell_val = cell.value
                if cell_val:
                    cell_text = str(cell_val).lower()
                    # Если в ячейке есть упоминание месяца или это похоже на дату
                    if any(month in cell_text for month in ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 
                                                            'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь']):
                        # Обновляем дату
                        russian_months = {
                            1: 'января', 2: 'февраля', 3: 'марта', 4: 'апреля', 5: 'мая', 6: 'июня',
                            7: 'июля', 8: 'августа', 9: 'сентября', 10: 'октября', 11: 'ноября', 12: 'декабря'
                        }
                        date_str = f"{menu_date.day} {russian_months.get(menu_date.month, 'сентября')}"
                        
                        # Проверяем, не является ли ячейка объединенной MergedCell
                        if cell.__class__.__name__ == 'MergedCell':
                            continue  # Пропускаем объединенные ячейки
                        
                        try:
                            cell.value = date_str
                            return
                        except AttributeError:
                            # Если не можем записать, продолжаем поиск
                            continue
    
    def fill_menu_template(self, template_path: str, source_menu_path: str, output_path: str) -> Tuple[bool, str]:
        """
        Заполняет шаблон меню данными из исходного файла
        
        Args:
            template_path: Путь к шаблону меню
            source_menu_path: Путь к файлу-источнику данных  
            output_path: Путь к выходному файлу
            
        Returns:
            (success, message): Кортеж с результатом операции
        """
        try:
            # Проверяем существование файлов
            if not Path(template_path).exists():
                return False, f"Шаблон не найден: {template_path}"
            
            if not Path(source_menu_path).exists():
                return False, f"Исходный файл не найден: {source_menu_path}"
            
            # Извлекаем данные из исходного файла (с продвинутой логикой)
            print(f"Извлекаем данные из файла: {source_menu_path}")
            categories = self.extract_categorized_dishes_advanced(source_menu_path)
            
            # Отладочный вывод
            total_dishes = sum(len(dishes) for dishes in categories.values())
            print(f"Извлечено блюд по категориям:")
            for cat, dishes in categories.items():
                print(f"  {cat}: {len(dishes)} блюд")
            print(f"Всего блюд: {total_dishes}")
            
            if total_dishes == 0:
                return False, "Не удалось извлечь блюда из исходного файла"
            
            # Извлекаем дату из исходного файла
            menu_date = self.extract_date_from_menu(source_menu_path)
            
            # Открываем шаблон
            wb = openpyxl.load_workbook(template_path)
            
            # Находим лист "Касса" для заполнения данными (там хранятся сами блюда)
            ws = None
            for sheet in wb.worksheets:
                if 'касс' in sheet.title.lower():
                    ws = sheet
                    break
            
            if not ws:
                # Fallback к основному листу
                ws = wb.active
                print("Предупреждение: Лист 'Касса' не найден, используем основной лист")
            else:
                print(f"Используем лист: {ws.title}")
            
            # Обновляем дату в шаблоне
            self.update_template_date(ws, menu_date)
            
            # Заполняем колонки данными
            total_filled = 0
            categories_filled = 0
            
            for category, dishes in categories.items():
                if not dishes:
                    continue
                
                # Находим соответствующую колонку в шаблоне
                template_header = self.categories_mapping.get(category)
                if not template_header:
                    print(f"Предупреждение: Не найдено соответствие для категории '{category}'")
                    continue
                
                col = self.find_column_by_header(ws, template_header)
                if not col:
                    print(f"Предупреждение: Не найдена колонка для '{template_header}'")
                    continue
                
                # Находим начальную строку для данных
                start_row = self.find_data_start_row(ws, col, category)
                
                # Особая логика для гарниров - ограничиваем диапазон
                if category == 'гарнир':
                    end_row = self.find_category_end_row(ws, col, start_row, category)
                    print(f"Гарниры: ограничиваем диапазон {start_row}-{end_row}")
                    filled_count = self.fill_template_column_limited(ws, col, dishes, start_row, end_row, category)
                else:
                    # Обычное заполнение для остальных категорий
                    filled_count = self.fill_template_column(ws, col, dishes, start_row)
                total_filled += filled_count
                categories_filled += 1
                
                print(f"Категория '{category}' -> колонка {col}: добавлено {filled_count} блюд")
            
            # Сохраняем результат
            wb.save(output_path)
            
            date_str = menu_date.strftime("%d.%m.%Y") if menu_date else "текущая дата"
            message = f"Шаблон меню заполнен для даты {date_str}\n"
            message += f"Заполнено категорий: {categories_filled}\n"
            message += f"Всего добавлено блюд: {total_filled}"
            
            return True, message
            
        except Exception as e:
            return False, f"Ошибка при заполнении шаблона: {str(e)}"


    def fill_menu_template_with_details(self, template_path: str, source_menu_path: str, output_path: str) -> Tuple[bool, str]:
        """
        Заполняет шаблон меню с деталями (название, вес, цена)
        
        Args:
            template_path: Путь к шаблону меню
            source_menu_path: Путь к файлу-источнику данных  
            output_path: Путь к выходному файлу
            
        Returns:
            (success, message): Кортеж с результатом операции
        """
        try:
            # Проверяем существование файлов
            if not Path(template_path).exists():
                return False, f"Шаблон не найден: {template_path}"
            
            if not Path(source_menu_path).exists():
                return False, f"Исходный файл не найден: {source_menu_path}"
            
            # Извлекаем данные с деталями из исходного файла
            print(f"Извлекаем данные с деталями из файла: {source_menu_path}")
            categories_detailed = self.extract_dishes_with_details(source_menu_path)
            
            # Отладочный вывод
            total_dishes = sum(len(dishes) for dishes in categories_detailed.values())
            if total_dishes == 0:
                return False, "Не удалось извлечь блюда с деталями из исходного файла"
            
            # Извлекаем дату из исходного файла
            menu_date = self.extract_date_from_menu(source_menu_path)
            
            # Открываем шаблон
            wb = openpyxl.load_workbook(template_path)
            
            # Находим лист "Касса" для заполнения данными (там хранятся сами блюда)
            ws = None
            for sheet in wb.worksheets:
                if 'касс' in sheet.title.lower():
                    ws = sheet
                    break
            
            if not ws:
                ws = wb.active
                print("Предупреждение: Лист 'Касса' не найден, используем основной лист")
            else:
                print(f"Используем лист: {ws.title}")
            
            # Обновляем дату в шаблоне
            self.update_template_date(ws, menu_date)
            
            # Определяем мапинг колонок для разных категорий
            column_mapping = {
                # Левая сторона: завтраки, салаты
                'завтрак': {'name': 1, 'weight': 2, 'price': 3},
                'салат': {'name': 1, 'weight': 2, 'price': 3},
                # Правая сторона: первые, мясо, птица, рыба, гарниры
                'первое': {'name': 4, 'weight': 5, 'price': 6},
                'мясо': {'name': 4, 'weight': 5, 'price': 6},
                'птица': {'name': 4, 'weight': 5, 'price': 6},
                'рыба': {'name': 4, 'weight': 5, 'price': 6},
                'гарнир': {'name': 4, 'weight': 5, 'price': 6}
            }
            
            # Заполняем колонки данными
            total_filled = 0
            categories_filled = 0
            
            for category, dishes_details in categories_detailed.items():
                if not dishes_details:
                    continue
                
                # Получаем колонки для этой категории
                if category not in column_mapping:
                    print(f"Предупреждение: Не найдено соответствие колонок для категории '{category}'")
                    continue
                
                cols = column_mapping[category]
                name_col, weight_col, price_col = cols['name'], cols['weight'], cols['price']
                
                # Находим начальную строку для данных
                start_row = self.find_data_start_row(ws, name_col, category)
                
                # Особая логика для гарниров - ограничиваем диапазон
                if category == 'гарнир':
                    end_row = self.find_category_end_row(ws, name_col, start_row, category)
                    print(f"Гарниры: ограничиваем диапазон {start_row}-{end_row}")
                    filled_count = self.fill_template_with_details_limited(
                        ws, name_col, weight_col, price_col, dishes_details, start_row, end_row, category
                    )
                else:
                    # Обычное заполнение для остальных категорий
                    filled_count = self.fill_template_with_details(
                        ws, name_col, weight_col, price_col, dishes_details, start_row
                    )
                
                total_filled += filled_count
                categories_filled += 1
                
                print(f"Категория '{category}' -> колонки {name_col}/{weight_col}/{price_col}: добавлено {filled_count} блюд")
            
            # Сохраняем результат
            wb.save(output_path)
            
            date_str = menu_date.strftime("%d.%m.%Y") if menu_date else "текущая дата"
            message = f"Шаблон меню заполнен (с деталями) для даты {date_str}\n"
            message += f"Заполнено категорий: {categories_filled}\n"
            message += f"Всего добавлено блюд: {total_filled}"
            
            return True, message
            
        except Exception as e:
            return False, f"Ошибка при заполнении шаблона с деталями: {str(e)}"


def fill_menu_template_from_source(template_path: str, source_menu_path: str, output_path: str) -> Tuple[bool, str]:
    """Удобная функция для заполнения шаблона меню"""
    filler = MenuTemplateFiller()
    return filler.fill_menu_template(template_path, source_menu_path, output_path)

def fill_menu_template_with_details_from_source(template_path: str, source_menu_path: str, output_path: str) -> Tuple[bool, str]:
    """Удобная функция для заполнения шаблона меню с деталями"""
    filler = MenuTemplateFiller()
    return filler.fill_menu_template_with_details(template_path, source_menu_path, output_path)

def fill_breakfast_only(template_path: str, source_menu_path: str, output_path: str) -> Tuple[bool, str]:
    """Копирует только завтраки с деталями (название, вес, цена)"""
    try:
        import openpyxl
        from pathlib import Path
        
        # Проверяем существование файлов
        if not Path(template_path).exists():
            return False, f"Шаблон не найден: {template_path}"
        
        if not Path(source_menu_path).exists():
            return False, f"Исходный файл не найден: {source_menu_path}"
        
        # Открываем исходный файл и ищем завтраки
        print(f"Извлекаем завтраки из {source_menu_path}")
        source_wb = openpyxl.load_workbook(source_menu_path)
        source_ws = source_wb.active
        
        breakfast_dishes = []
        in_breakfast_section = False
        
        for row in range(1, min(100, source_ws.max_row + 1)):
            name_cell = source_ws.cell(row=row, column=1).value
            if not name_cell:
                continue
                
            name_text = str(name_cell).lower().strip()
            
            # Начало секции завтраков
            if 'завтрак' in name_text:
                in_breakfast_section = True
                print(f"Начало секции завтраков в строке {row}: {name_cell}")
                continue
            
            # Конец секции завтраков
            if in_breakfast_section and any(word in name_text for word in ['салат', 'холодн', 'перв', 'блюда']):
                print(f"Конец секции завтраков в строке {row}: {name_cell}")
                break
            
            # Собираем блюда завтраков
            if in_breakfast_section:
                # Пропускаем служебные строки
                if any(word in name_text for word in ['вес', 'цена', 'руб']):
                    continue
                
                # Получаем вес и цену
                weight = source_ws.cell(row=row, column=2).value
                price = source_ws.cell(row=row, column=3).value
                
                dish = {
                    'name': str(name_cell).strip(),
                    'weight': str(weight).strip() if weight else '',
                    'price': str(price).strip() if price else ''
                }
                
                breakfast_dishes.append(dish)
                print(f"Найдено блюдо: {dish}")
        
        source_wb.close()
        
        if not breakfast_dishes:
            return False, "Не найдено блюд для завтрака"
        
        print(f"Найдено {len(breakfast_dishes)} блюд для завтрака")
        
        # Открываем шаблон и заполняем завтраки
        template_wb = openpyxl.load_workbook(template_path)
        
        # Находим лист Касса
        template_ws = None
        for sheet in template_wb.worksheets:
            if 'касс' in sheet.title.lower():
                template_ws = sheet
                break
        
        if not template_ws:
            template_ws = template_wb.active
            print("Предупреждение: Лист 'Касса' не найден")
        
        # Находим заголовок завтраков в шаблоне
        start_row = None
        for row in range(1, 20):
            cell_val = template_ws.cell(row=row, column=1).value
            if cell_val and 'завтрак' in str(cell_val).lower():
                start_row = row + 1
                print(f"Найден заголовок завтраков в строке {row}, начинаем с {start_row}")
                break
        
        if not start_row:
            start_row = 7  # По умолчанию
            print(f"Заголовок не найден, начинаем с строки {start_row}")
        
        # Заполняем завтраки
        filled_count = 0
        current_row = start_row
        
        for dish in breakfast_dishes:
            # Пропускаем занятые строки
            while current_row <= template_ws.max_row:
                existing_name = template_ws.cell(row=current_row, column=1).value
                if not existing_name or str(existing_name).strip() == '':
                    break
                current_row += 1
            
            if current_row > template_ws.max_row:
                print(f"Достигнут конец листа")
                break
            
            # Заполняем данные
            template_ws.cell(row=current_row, column=1).value = dish['name']      # Название
            template_ws.cell(row=current_row, column=2).value = dish['weight']    # Вес
            template_ws.cell(row=current_row, column=3).value = dish['price']     # Цена
            
            filled_count += 1
            print(f"Заполнена строка {current_row}: {dish['name']} | {dish['weight']} | {dish['price']}")
            current_row += 1
        
        # Сохраняем
        template_wb.save(output_path)
        template_wb.close()
        
        message = f"Успешно скопировано {filled_count} завтраков с деталями"
        return True, message
        
    except Exception as e:
        return False, f"Ошибка: {str(e)}"
