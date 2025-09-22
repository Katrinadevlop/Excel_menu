import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from pathlib import Path
from datetime import datetime
import re
from typing import List, Dict, Optional, Tuple

class BrokerageJournalGenerator:
    """Генератор бракеражного журнала на основе меню"""
    
    def __init__(self):
        pass
    
    def extract_date_from_menu(self, menu_path: str) -> Optional[datetime]:
        """Извлекает дату из файла меню"""
        try:
            # Пробуем разные способы чтения файла
            if menu_path.endswith('.xls'):
                # Для старых xls файлов используем pandas
                df_dict = pd.read_excel(menu_path, sheet_name=None)
                
                # Ищем дату во всех листах
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
                    
                    # Ищем дату в первых нескольких строках
                    for _, row in df.head(10).iterrows():
                        for cell in row:
                            if pd.notna(cell):
                                date_from_cell = self._parse_date_string(str(cell))
                                if date_from_cell:
                                    return date_from_cell
            else:
                # Для xlsx файлов используем openpyxl
                wb = openpyxl.load_workbook(menu_path, data_only=True)
                
                # Проверяем название файла
                filename = Path(menu_path).stem
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
        except Exception as e:
            print(f"Ошибка при извлечении даты: {e}")
        
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
                            # Используем текущий год или год из контекста
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
    
    def extract_dishes_from_menu(self, menu_path: str) -> List[str]:
        """Извлекает блюда из меню последовательно до первой пустой строки"""
        dishes = []
        
        try:
            if menu_path.endswith('.xls'):
                # Для старых xls файлов
                df_dict = pd.read_excel(menu_path, sheet_name=None)
                
                # Ищем лист с кассой или с меню
                target_sheet = None
                for sheet_name, df in df_dict.items():
                    if 'касс' in str(sheet_name).lower() or 'меню' in str(sheet_name).lower():
                        target_sheet = df
                        break
                
                if target_sheet is None and df_dict:
                    target_sheet = list(df_dict.values())[0]  # Берем первый лист
                
                if target_sheet is not None:
                    dishes = self._extract_sequential_dishes_from_dataframe(target_sheet)
            else:
                # Для xlsx файлов
                wb = openpyxl.load_workbook(menu_path, data_only=True)
                if wb.worksheets:
                    dishes = self._extract_sequential_dishes_from_worksheet(wb.worksheets[0])
        
        except Exception as e:
            print(f"Ошибка при извлечении блюд: {e}")
        
        return dishes
    
    def _extract_sequential_dishes_from_dataframe(self, df: pd.DataFrame) -> List[str]:
        """Извлекает блюда последовательно из DataFrame"""
        dishes = []
        
        # Проходим по всем строкам подряд
        for row_idx, row in df.iterrows():
            row_has_dish = False
            
            # Проверяем все ячейки в строке
            for cell in row:
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    
                    # Пропускаем служебные ячейки
                    if self._should_skip_cell(cell_str):
                        continue
                    
                    # Проверяем, является ли это валидным блюдом
                    if self._is_valid_dish(cell_str, dishes) and len(cell_str) > 8:
                        dishes.append(cell_str)
                        row_has_dish = True
                        break  # Одно блюдо на строку
            
            # Если строка пустая (нет валидных блюд), останавливаемся
            if not row_has_dish and len(dishes) > 0:
                break
        
        return dishes
    
    def _extract_sequential_dishes_from_worksheet(self, ws) -> List[str]:
        """Извлекает блюда последовательно из worksheet"""
        dishes = []
        
        # Проходим по всем строкам подряд
        for row in range(1, ws.max_row + 1):
            row_has_dish = False
            
            # Проверяем все ячейки в строке
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                if cell.value:
                    cell_str = str(cell.value).strip()
                    
                    # Пропускаем служебные ячейки
                    if self._should_skip_cell(cell_str):
                        continue
                    
                    # Проверяем, является ли это валидным блюдом
                    if self._is_valid_dish(cell_str, dishes) and len(cell_str) > 8:
                        dishes.append(cell_str)
                        row_has_dish = True
                        break  # Одно блюдо на строку
            
            # Если строка пустая (нет валидных блюд), останавливаемся
            if not row_has_dish and len(dishes) > 0:
                break
        
        return dishes
    
    def extract_categorized_dishes(self, menu_path: str) -> Dict[str, List[str]]:
        """Извлекает блюда из меню, распределяя по колонкам.
        Столбец A (0) - завтраки, столбец E (4) - первые блюда, мясо, курица, рыба, гарниры
        """
        result: Dict[str, List[str]] = {k: [] for k in ['завтрак','салат','первое','мясо','курица','рыба','гарнир']}
        
        def add_from_dataframe(df: pd.DataFrame):
            print(f"\nОтладка: DataFrame имеет {len(df)} строк и {len(df.columns)} колонок")
            
            # Находим строку заголовков (где есть ЗАВТРАКИ)
            header_row = None
            for idx in range(min(10, len(df))):
                row = df.iloc[idx]
                row_text = ' '.join([str(cell).strip() for cell in row if pd.notna(cell) and str(cell).strip()])
                if 'ЗАВТРАКИ' in row_text.upper():
                    header_row = idx
                    break
            
            if header_row is None:
                print("Отладка: Не найден заголовок ЗАВТРАКИ")
                return
            
            print(f"Отладка: Найден заголовок в строке {header_row}")
            
            # Извлекаем завтраки из первого столбца (A)
            current_category = None
            for row_idx in range(header_row + 1, len(df)):
                cell_value = df.iloc[row_idx, 0]  # Столбец A
                if pd.notna(cell_value):
                    cell_str = str(cell_value).strip()
                    if not self._should_skip_cell(cell_str) and self._is_valid_dish(cell_str, []):
                        result['завтрак'].append(cell_str)
                elif len(result['завтрак']) > 0:  # Если уже есть блюда и встретили пустую ячейку
                    break
            
            # Извлекаем остальные блюда из правых столбцов (начиная с 3-го)
            current_category = None
            
            # Проверяем колонки 3, 4, 5... (D, E, F...)
            for col_idx in range(3, len(df.columns)):
                print(f"Отладка: Проверяем колонку {col_idx} ({chr(65+col_idx)})")
                
                for row_idx in range(header_row, len(df)):
                    cell_value = df.iloc[row_idx, col_idx]
                    if pd.notna(cell_value):
                        cell_str = str(cell_value).strip()
                        
                        # Проверяем, является ли это заголовком категории
                        if 'ПЕРВЫЕ' in cell_str.upper() and 'БЛЮДА' in cell_str.upper():
                            current_category = 'первое'
                            print(f"Отладка: Найден заголовок ПЕРВЫЕ БЛЮДА в колонке {col_idx}, строке {row_idx}")
                        elif 'БЛЮДА ИЗ МЯСА' in cell_str.upper():
                            current_category = 'мясо'
                            print(f"Отладка: Найден заголовок МЯСО в колонке {col_idx}, строке {row_idx}")
                        elif 'БЛЮДА ИЗ ПТИЦЫ' in cell_str.upper():
                            current_category = 'курица'
                            print(f"Отладка: Найден заголовок ПТИЦА в колонке {col_idx}, строке {row_idx}")
                        elif 'БЛЮДА ИЗ РЫБЫ' in cell_str.upper():
                            current_category = 'рыба'
                            print(f"Отладка: Найден заголовок РЫБА в колонке {col_idx}, строке {row_idx}")
                        elif 'ГАРНИРЫ' in cell_str.upper():
                            current_category = 'гарнир'
                            print(f"Отладка: Найден заголовок ГАРНИРЫ в колонке {col_idx}, строке {row_idx}")
                        elif 'НАПИТКИ' in cell_str.upper():
                            # Останавливаемся на напитках - больше ничего не добавляем
                            print(f"Отладка: Найдены НАПИТКИ в колонке {col_idx}, строке {row_idx}, прекращаем сбор")
                            return  # Полностью выходим из функции
                        elif 'САЛАТ' in cell_str.upper() or 'ХОЛОДН' in cell_str.upper():
                            current_category = 'салат'
                            print(f"Отладка: Найден заголовок САЛАТЫ в колонке {col_idx}, строке {row_idx}")
                        elif current_category and not self._should_skip_cell(cell_str) and self._is_valid_dish(cell_str, []):
                            # Это блюдо для текущей категории
                            result[current_category].append(cell_str)
                            print(f"Отладка: Добавлено блюдо '{cell_str}' в категорию '{current_category}' (колонка {col_idx})")
        
        def add_from_worksheet(ws):
            print(f"\nОтладка: Worksheet имеет {ws.max_row} строк и {ws.max_column} колонок")
            
            # Находим строку заголовков (где есть ЗАВТРАКИ)
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
                print("Отладка: Не найден заголовок ЗАВТРАКИ в worksheet")
                return
            
            print(f"Отладка: Найден заголовок в строке {header_row}")
            
            # Извлекаем завтраки и салаты из первого столбца (A)
            current_category = 'завтрак'  # Начинаем с завтраков
            
            for row_idx in range(header_row + 1, ws.max_row + 1):
                cell_value = ws.cell(row=row_idx, column=1).value  # Столбец A
                if cell_value:
                    cell_str = str(cell_value).strip()
                    
                    # Проверяем, является ли это заголовком салатов
                    if 'САЛАТ' in cell_str.upper() and 'ХОЛОДН' in cell_str.upper():
                        current_category = 'салат'  # Переключаемся на салаты
                        print(f"Отладка: Переключились на салаты в строке {row_idx}")
                        continue  # Пропускаем заголовок
                    
                    # Проверяем на другие заголовки разделов
                    if ('СЭНДВИЧ' in cell_str.upper() or 
                        'ПЕЛЬМЕН' in cell_str.upper() or
                        'ВАРЕНИК' in cell_str.upper()):
                        # Все равно добавляем блюда этих категорий в завтраки
                        if not self._should_skip_cell(cell_str) and self._is_valid_dish(cell_str, []):
                            result['завтрак'].append(cell_str)
                            print(f"Отладка: Добавлено блюдо '{cell_str}' в завтраки из раздела {cell_str[:20]}")
                        continue
                    
                    # Добавляем блюдо в текущую категорию
                    if not self._should_skip_cell(cell_str) and self._is_valid_dish(cell_str, []):
                        result[current_category].append(cell_str)
                        print(f"Отладка: Добавлено блюдо '{cell_str}' в категорию '{current_category}'")
                # Пустая ячейка - продолжаем в той же категории
            
            # Извлекаем остальные блюда со всех столбцов, начиная с четвертого
            print(f"Отладка: Поиск данных в правых столбцах (колонки 4-{ws.max_column})")
            
            # Проходим по всем столбцам начиная с 4-го (D)
            current_category = None
            for col_idx in range(4, ws.max_column + 1):  # Колонки D, E, F, G, H...
                print(f"Отладка: Проверяем колонку {col_idx} ({chr(64+col_idx)})")
                
                for row_idx in range(header_row, ws.max_row + 1):
                    cell_value = ws.cell(row=row_idx, column=col_idx).value
                    if cell_value:
                        cell_str = str(cell_value).strip()
                        
                        # Проверяем, является ли это заголовком категории
                        if 'ПЕРВЫЕ' in cell_str.upper() and 'БЛЮДА' in cell_str.upper():
                            current_category = 'первое'
                            print(f"Отладка: Найден заголовок ПЕРВЫЕ БЛЮДА в колонке {col_idx}, строке {row_idx}")
                        elif 'БЛЮДА ИЗ МЯСА' in cell_str.upper():
                            current_category = 'мясо'
                            print(f"Отладка: Найден заголовок МЯСО в колонке {col_idx}, строке {row_idx}")
                        elif 'БЛЮДА ИЗ ПТИЦЫ' in cell_str.upper():
                            current_category = 'курица'
                            print(f"Отладка: Найден заголовок ПТИЦА в колонке {col_idx}, строке {row_idx}")
                        elif 'БЛЮДА ИЗ РЫБЫ' in cell_str.upper():
                            current_category = 'рыба'
                            print(f"Отладка: Найден заголовок РЫБА в колонке {col_idx}, строке {row_idx}")
                        elif 'ГАРНИРЫ' in cell_str.upper():
                            current_category = 'гарнир'
                            print(f"Отладка: Найден заголовок ГАРНИРЫ в колонке {col_idx}, строке {row_idx}")
                        elif 'НАПИТКИ' in cell_str.upper():
                            # Останавливаемся на напитках - больше ничего не добавляем
                            print(f"Отладка: Найдены НАПИТКИ в колонке {col_idx}, строке {row_idx}, полностью прекращаем сбор")
                            return  # Полностью выходим из функции
                        elif current_category and not self._should_skip_cell(cell_str) and self._is_valid_dish(cell_str, []):
                            # Это блюдо для текущей категории
                            result[current_category].append(cell_str)
                            print(f"Отладка: Добавлено блюдо '{cell_str}' в категорию '{current_category}' (колонка {col_idx})")
        
        try:
            if menu_path.endswith('.xls'):
                df_dict = pd.read_excel(menu_path, sheet_name=None)
                # ПРИОРИТЕТ: ищем лист с названием 'касс', иначе любой другой
                preferred = None
                # Сначала ищем лист с 'касс' в названии
                for name, df in df_dict.items():
                    nm = str(name).lower()
                    if 'касс' in nm:
                        preferred = df
                        break
                # Если не нашли касс, берем любой лист с 'меню'
                if preferred is None:
                    for name, df in df_dict.items():
                        nm = str(name).lower()
                        if 'меню' in nm:
                            preferred = df
                            break
                # В крайнем случае берем первый лист
                if preferred is None and df_dict:
                    preferred = list(df_dict.values())[0]
                if preferred is not None:
                    add_from_dataframe(preferred)
            else:
                wb = openpyxl.load_workbook(menu_path, data_only=True)
                if wb.worksheets:
                    add_from_worksheet(wb.worksheets[0])
        except Exception as e:
            print(f"Ошибка при извлечении категорий: {e}")
        
        return result
    
    def _should_skip_cell(self, cell_str: str) -> bool:
        """Проверяет, нужно ли пропустить ячейку при обработке"""
        cell_lower = cell_str.lower().strip()
        
        # Пропускаем пустые или слишком короткие строки
        if len(cell_str) < 4:
            return True
        
        # Пропускаем служебные строки
        skip_words = [
            'вес', 'цена', 'руб', 'ед.изм', 'утверждаю', 'директор', 
            'меню', 'столовой', 'патриот', 'москва', 'наб', 'стр',
            'попова', 'сентября', 'пятница', '_____', 'понедельник', 'вторник',
            'среда', 'четверг', 'пятница', 'суббота', 'воскресенье',
            'январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль',
            'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь',
            'завтрак', 'обед', 'ужин', 'полдник', 'время', 'дата',
            # напитки и подобные
            'напит', 'сок', 'чай', 'кофе', 'смузи', 'фреш',
            # соусы - исключаем полностью
            'соус', 'майонез', 'кетчуп',
            # дополнительные служебные
            'наименование', 'блюда', 'час', 'мин'
        ]
        
        # Проверяем, является ли строка временем (HH:MM:SS)
        if re.match(r'^\d{1,2}:\d{2}(:\d{2})?$', cell_str):
            return True
        
        for skip_word in skip_words:
            if skip_word in cell_lower:
                return True
        
        # Пропускаем строки с объемами напитков
        if re.search(r'\d+/\d+\s*мл|\d+\s*мл|200/300мл|300мл|225\s*мл', cell_str):
            return True
        
        # Пропускаем строки, состоящие только из чисел и символов
        if re.match(r'^[\d\s\.,/г-]+$', cell_str):
            return True
        
        # Пропускаем адреса
        if 'ул.' in cell_lower or 'д.' in cell_lower or 'овчинниковская' in cell_lower:
            return True
            
        return False
    
    def _find_category_in_text(self, text: str) -> Optional[str]:
        """Находит категорию блюда в тексте по заголовкам разделов"""
        text_lower = text.lower().strip()
        
        # Ищем только явные заголовки разделов
        if 'завтрак' in text_lower and len(text_lower) < 30:
            return 'завтрак'
        if ('салат' in text_lower and 'холодн' in text_lower) or \
           ('холодн' in text_lower and 'закуск' in text_lower):
            return 'салат'
        if 'первое' in text_lower or ('первы' in text_lower and 'блюд' in text_lower):
            return 'первое'
        if ('мясо' in text_lower or 'говядин' in text_lower or 'свинин' in text_lower) and len(text_lower) < 30:
            return 'мясо'
        if ('курица' in text_lower or 'куриц' in text_lower or 'птица' in text_lower) and len(text_lower) < 30:
            return 'курица'
        if ('рыба' in text_lower or 'рыбн' in text_lower) and len(text_lower) < 30:
            return 'рыба'
        if 'гарнир' in text_lower and len(text_lower) < 30:
            return 'гарнир'
        
        return None
    
    def _is_valid_dish(self, dish_name: str, existing_dishes: List[str]) -> bool:
        """Проверяет, является ли строка валидным названием блюда"""
        # Уже есть в списке
        if dish_name in existing_dishes:
            return False
        
        # Слишком короткое название
        if len(dish_name) < 5:
            return False
        
        # Это явный заголовок категории - исключаем
        dish_lower = dish_name.lower().strip()
        category_titles = [
            'завтрак', 'салат', 'холодн', 'закуск', 
            'первое', 'первы', 'блюд', 'гарнир', 
            'мясо', 'курица', 'птица', 'рыба'
        ]
        
        # Если это короткая строка, содержащая только название категории
        if len(dish_name) < 35:
            for title in category_titles:
                if dish_lower == title or dish_lower.startswith(title + ' ') or dish_lower.endswith(' ' + title):
                    return False
        
        # Содержит только числа и специальные символы
        if re.match(r'^[\d\s\.,/г-]+$', dish_name):
            return False
            
        return True

    def _is_section_header(self, text: str) -> bool:
        """Определяет, является ли строка заголовком раздела (не блюдом)."""
        if not text:
            return False
        txt = str(text).strip()
        lower = txt.lower()
        
        # Точные заголовки разделов, которые точно нужно исключить
        exact_headers = [
            'салаты и холодные закуски',
            'сэндвичи',
            'пельмени', 
            'вареники',
            'соусы',
            'первые блюда',
            'блюда из мяса',
            'блюда из птицы',
            'блюда из рыбы',
            'гарниры',
            'завтраки'
        ]
        
        # Проверяем точное совпадение с заголовками
        if lower in exact_headers:
            return True
            
        # Строки полностью в верхнем регистре и очень короткие (вероятно заголовки)
        if txt.isupper() and len(txt) <= 15 and any(h in lower for h in ['блюд', 'салат', 'сэндвич']):
            return True
            
        return False
    
    def _extract_from_worksheet(self, ws, dishes: Dict[str, List[str]]):
        """Извлекает блюда из листа openpyxl"""
        current_category = None
        
        for row in range(1, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                if cell.value:
                    cell_str = str(cell.value).strip()
                    
                    # Пропускаем служебные строки
                    if self._should_skip_cell(cell_str):
                        continue
                    
                    # Ищем заголовки категорий
                    category_found = self._find_category_in_text(cell_str)
                    if category_found:
                        current_category = category_found
                        continue
                    
                    # Добавляем блюда в текущую категорию
                    if current_category and self._is_valid_dish(cell_str, dishes[current_category]):
                        dishes[current_category].append(cell_str)
    
    def create_brokerage_journal(self, menu_path: str, template_path: str, output_path: str) -> Tuple[bool, str]:
        """Создает бракеражный журнал на основе меню с листа касс, заполняя столбец A завтраками, а столбец G - остальными блюдами. Время не изменяем."""
        try:
            # Проверяем существование шаблона
            if not Path(template_path).exists():
                return False, f"Шаблон бракеражного журнала не найден: {template_path}"
            
            # Извлекаем дату из меню
            menu_date = self.extract_date_from_menu(menu_path)
            if not menu_date:
                menu_date = datetime.now()
            
            # Извлекаем блюда по категориям
            categories = self.extract_categorized_dishes(menu_path)
            
            # Отладочный вывод результата
            print(f"\nРезультат извлечения категорий:")
            for category, dishes in categories.items():
                print(f"{category}: {len(dishes)} блюд - {dishes}")
            
            # Собираем завтраки и салаты для левого столбца (A)
            left_list: List[str] = []
            left_list.extend(categories.get('завтрак', []))
            left_list.extend(categories.get('салат', []))
            # Удаляем заголовки разделов
            left_list = [d for d in left_list if not self._is_section_header(d)]
            print(f"\nКоличество блюд для левого столбца (A): {len(left_list)}")
            
            # Собираем остальные блюда для правого столбца (G)
            right_list: List[str] = []
            right_list.extend(categories.get('первое', []))
            right_list.extend(categories.get('мясо', []))
            right_list.extend(categories.get('курица', []))
            right_list.extend(categories.get('рыба', []))
            right_list.extend(categories.get('гарнир', []))
            # Удаляем заголовки разделов
            right_list = [d for d in right_list if not self._is_section_header(d)]
            print(f"Количество блюд для правого столбца (G): {len(right_list)}")
            
            # Открываем шаблон
            wb = openpyxl.load_workbook(template_path)
            ws = wb.active
            
            # Обновляем дату в шаблоне (строка 3, колонка 1)
            russian_months = {
                1: 'января', 2: 'февраля', 3: 'марта', 4: 'апреля', 5: 'мая', 6: 'июня',
                7: 'июля', 8: 'августа', 9: 'сентября', 10: 'октября', 11: 'ноября', 12: 'декабря'
            }
            date_str_display = f"{menu_date.day} {russian_months.get(menu_date.month, 'unknown')}"
            ws.cell(row=3, column=1, value=date_str_display)
            
            # Форматируем дату для названия листа
            date_str = menu_date.strftime('%d.%m.%y')
            ws.title = date_str
            
            # Находим строку заголовков таблицы ("НАИМЕНОВАНИЕ БЛЮДА" в колонке A)
            header_row = None
            for r in range(1, ws.max_row + 1):
                v = ws.cell(row=r, column=1).value
                if v and 'наименование' in str(v).lower() and 'блюд' in str(v).lower():
                    header_row = r
                    break
            if header_row is None:
                return False, 'Не удалось определить заголовок таблицы в шаблоне'
            
            start_row = header_row + 1  # первая строка для блюд
            
            # Определяем первую полностью пустую строку: дальше НИЧЕГО не трогаем
            stop_row = start_row
            while stop_row <= ws.max_row:
                row_empty = True
                for c in range(1, 10):  # A..I
                    if ws.cell(row=stop_row, column=c).value not in (None, ''):
                        row_empty = False
                        break
                if row_empty:
                    break
                stop_row += 1
            
            # Заполняем ТОЛЬКО пустые ячейки первого столбца (A) блюдами из left_list
            inserted_left = 0
            dish_idx = 0
            for r in range(start_row, stop_row):
                if dish_idx >= len(left_list):
                    break
                current_val = ws.cell(row=r, column=1).value
                if current_val in (None, ''):
                    ws.cell(row=r, column=1, value=left_list[dish_idx])
                    # НЕ МЕНЯЕМ время - оставляем как в шаблоне
                    dish_idx += 1
                    inserted_left += 1
                else:
                    # Ячейка занята — не трогаем её и идем дальше
                    continue
            
            # Заполняем ТОЛЬКО пустые ячейки седьмого столбца (G) блюдами из right_list
            inserted_right = 0
            dish_idx = 0
            for r in range(start_row, stop_row):
                if dish_idx >= len(right_list):
                    break
                current_val = ws.cell(row=r, column=7).value  # Столбец G = 7
                if current_val in (None, ''):
                    ws.cell(row=r, column=7, value=right_list[dish_idx])
                    # НЕ МЕНЯЕМ время - оставляем как в шаблоне
                    dish_idx += 1
                    inserted_right += 1
                else:
                    # Ячейка занята — не трогаем её и идем дальше
                    continue
            
            wb.save(output_path)
            return True, f"Бракеражный журнал создан успешно для даты {date_str} (вставлено {inserted_left} блюд в колонку A, {inserted_right} блюд в колонку G)"
        except Exception as e:
            return False, f"Ошибка при создании бракеражного журнала: {str(e)}"
    
    def _create_header(self, ws, date: datetime):
        """Создает заголовок бракеражного журнала"""
        # Дата
        date_str = date.strftime("%d %B").replace('September', 'сентября').replace('August', 'августа')
        ws.cell(row=3, column=1, value=date_str)
        
        # Заголовок
        ws.cell(row=1, column=1, value="БРАКЕРАЖНЫЙ ЖУРНАЛ")
        ws.cell(row=1, column=1).font = Font(bold=True, size=14)
        ws.cell(row=1, column=1).alignment = Alignment(horizontal='center')
        
        # Объединяем ячейки для заголовка
        ws.merge_cells('A1:G1')


def create_brokerage_journal_from_menu(menu_path: str, template_path: str, output_path: str) -> Tuple[bool, str]:
    """Удобная функция для создания бракеражного журнала"""
    generator = BrokerageJournalGenerator()
    return generator.create_brokerage_journal(menu_path, template_path, output_path)
