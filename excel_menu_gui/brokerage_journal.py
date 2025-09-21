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
        """Извлекает блюда из меню, распределяя по нужным категориям в исходном порядке.
        Категории: завтрак, салат, первое, мясо, курица, рыба, гарнир
        """
        result: Dict[str, List[str]] = {k: [] for k in ['завтрак','салат','первое','мясо','курица','рыба','гарнир']}
        
        def add_from_dataframe(df: pd.DataFrame):
            # Новая логика: все блюда с начала листа до раздела «СЭНДВИЧИ» — это завтраки
            breakfast_mode = True  # До «СЭНДВИЧИ» всё считается завтраками
            current_category: Optional[str] = None
            started = False
            
            print(f"\nОтладка: DataFrame имеет {len(df)} строк и {len(df.columns)} колонок")
            
            for row_idx, row in df.iterrows():
                row_has_value = False
                
                # Собираем всю строку в одну строку
                row_text = ' '.join([str(cell).strip() for cell in row if pd.notna(cell) and str(cell).strip()])
                row_text_upper = row_text.upper()

                # Переключение режимов: пока breakfast_mode=True, реагируем только на СЭНДВИЧИ
                if breakfast_mode:
                    if 'СЭНДВИЧ' in row_text_upper or 'СЭНДВИЧИ' in row_text_upper:
                        breakfast_mode = False
                        current_category = None
                        row_has_value = True
                        started = True
                        print("Отладка: Найден раздел СЭНДВИЧИ — завершили сбор завтраков")
                else:
                    # После завершения завтраков — обычная категоризация
                    if 'САЛАТ' in row_text_upper and ('ХОЛОДН' in row_text_upper or 'ЗАКУСК' in row_text_upper):
                        current_category = 'салат'
                        row_has_value = True
                        started = True
                    elif 'ПЕРВЫЕ' in row_text_upper and 'БЛЮДА' in row_text_upper:
                        current_category = 'первое'
                        row_has_value = True
                        started = True
                    elif 'БЛЮДА ИЗ МЯСА' in row_text_upper:
                        current_category = 'мясо'
                        row_has_value = True
                        started = True
                    elif 'БЛЮДА ИЗ ПТИЦЫ' in row_text_upper:
                        current_category = 'курица'
                        row_has_value = True
                        started = True
                    elif 'БЛЮДА ИЗ РЫБЫ' in row_text_upper:
                        current_category = 'рыба'
                        row_has_value = True
                        started = True
                    elif 'ГАРНИРЫ' in row_text_upper:
                        current_category = 'гарнир'
                        row_has_value = True
                        started = True
                
                # Если это не заголовок, ищем блюда
                if not row_has_value:
                    # Проверяем все ячейки в строке
                    for cell in row:
                        if pd.notna(cell) and str(cell).strip():
                            cell_str = str(cell).strip()
                            if not self._should_skip_cell(cell_str) and self._is_valid_dish(cell_str, []):
                                if breakfast_mode:
                                    result['завтрак'].append(cell_str)
                                    # print(f"Отладка: Добавлено блюдо завтрака: '{cell_str}'")
                                elif current_category:
                                    result[current_category].append(cell_str)
                                row_has_value = True
                                started = True
                                break  # Одно блюдо на строку
                
                # Останавливаемся на первой полностью пустой строке после начала списка
                if not row_has_value and started:
                    break
        
        def add_from_worksheet(ws):
            # Новая логика: все блюда до появления «СЭНДВИЧИ» — это завтраки
            breakfast_mode = True  # Начинаем с режима завтрака
            current_category: Optional[str] = None
            started = False
            
            for r in range(1, ws.max_row + 1):
                row_has_value = False
                
                # Собираем весь текст строки
                row_texts = []
                for c in range(1, ws.max_column + 1):
                    val = ws.cell(row=r, column=c).value
                    if val and str(val).strip():
                        row_texts.append(str(val).strip())
                
                row_text = ' '.join(row_texts)
                row_text_upper = row_text.upper()
                
                if breakfast_mode:
                    if 'СЭНДВИЧ' in row_text_upper or 'СЭНДВИЧИ' in row_text_upper:
                        breakfast_mode = False
                        current_category = None
                        row_has_value = True
                        started = True
                else:
                    if 'САЛАТ' in row_text_upper and ('ХОЛОДН' in row_text_upper or 'ЗАКУСК' in row_text_upper):
                        current_category = 'салат'
                        row_has_value = True
                        started = True
                    elif 'ПЕРВЫЕ' in row_text_upper and 'БЛЮДА' in row_text_upper:
                        current_category = 'первое'
                        row_has_value = True
                        started = True
                    elif 'БЛЮДА ИЗ МЯСА' in row_text_upper:
                        current_category = 'мясо'
                        row_has_value = True
                        started = True
                    elif 'БЛЮДА ИЗ ПТИЦЫ' in row_text_upper:
                        current_category = 'курица'
                        row_has_value = True
                        started = True
                    elif 'БЛЮДА ИЗ РЫБЫ' in row_text_upper:
                        current_category = 'рыба'
                        row_has_value = True
                        started = True
                    elif 'ГАРНИРЫ' in row_text_upper:
                        current_category = 'гарнир'
                        row_has_value = True
                        started = True
                
                # Если это не заголовок, ищем блюда
                if not row_has_value:
                    for c in range(1, ws.max_column + 1):
                        val = ws.cell(row=r, column=c).value
                        if val and str(val).strip():
                            cell_str = str(val).strip()
                            if not self._should_skip_cell(cell_str) and self._is_valid_dish(cell_str, []):
                                if breakfast_mode:
                                    result['завтрак'].append(cell_str)
                                elif current_category:
                                    result[current_category].append(cell_str)
                                row_has_value = True
                                started = True
                                break
                
                if not row_has_value and started:
                    break
        
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
            'напит', 'сок', 'чай', 'кофе',
            # дополнительные служебные
            'наименование', 'блюда', 'час', 'мин'
        ]
        
        # Проверяем, является ли строка временем (HH:MM:SS)
        if re.match(r'^\d{1,2}:\d{2}(:\d{2})?$', cell_str):
            return True
        
        for skip_word in skip_words:
            if skip_word in cell_lower:
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
        """Создает бракеражный журнал на основе меню с листа касс, заполняя только первый столбец ЗАВТРАКАМИ. Время не изменяем."""
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
            
            # Собираем ТОЛЬКО завтраки для первого столбца
            left_list: List[str] = []
            left_list.extend(categories.get('завтрак', []))
            print(f"\nКоличество завтраков для вставки: {len(left_list)}")
            
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
            
            # Заполняем ТОЛЬКО пустые ячейки первого столбца (A) блюдами из left_list, не меняя ничего другого
            inserted = 0
            dish_idx = 0
            for r in range(start_row, stop_row):
                if dish_idx >= len(left_list):
                    break
                current_val = ws.cell(row=r, column=1).value
                if current_val in (None, ''):
                    ws.cell(row=r, column=1, value=left_list[dish_idx])
                    # НЕ МЕНЯЕМ время - оставляем как в шаблоне
                    dish_idx += 1
                    inserted += 1
                else:
                    # Ячейка занята — не трогаем её и идем дальше
                    continue
            
            wb.save(output_path)
            return True, f"Бракеражный журнал создан успешно для даты {date_str} (вставлено {inserted} блюд в левую колонку)"
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
