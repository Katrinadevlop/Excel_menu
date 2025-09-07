import pandas as pd
import shutil
from pathlib import Path
from typing import List, Tuple, Optional
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from dataclasses import dataclass


@dataclass
class MenuItem:
    name: str
    weight: str
    price: str


def extract_salads_from_excel(excel_path: str) -> List[MenuItem]:
    """
    Извлекает данные салатов из Excel файла.
    Ищет секцию 'САЛАТЫ И ХОЛОДНЫЕ ЗАКУСКИ' и извлекает название, вес и цену.
    """
    try:
        # Читаем Excel файл
        df = pd.read_excel(excel_path, sheet_name='касса ', header=None)
        
        # Найдем строку с 'САЛАТЫ'
        salads_row = None
        for i, row in df.iterrows():
            if any(pd.notna(cell) and 'САЛАТ' in str(cell).upper() and 'ХОЛОДН' in str(cell).upper() for cell in row):
                salads_row = i
                break
        
        if salads_row is None:
            return []
        
        # Извлекаем данные салатов
        salads = []
        current_row = salads_row + 1
        
        # Читаем до тех пор, пока не встретим следующую категорию (заглавными буквами)
        while current_row < len(df):
            row = df.iloc[current_row]
            
            # Проверяем, не началась ли новая категория
            first_col = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
            if first_col.isupper() and len(first_col) > 3:
                break
                
            # Извлекаем данные салата
            name = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
            weight = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
            price = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
            
            # Проверяем, что это действительно блюдо (не пустая строка и не заголовок)
            if name and not name.isupper() and price and price.replace('.', '').isdigit():
                salads.append(MenuItem(name=name, weight=weight, price=f"{price} руб."))
            
            current_row += 1
            
        return salads
        
    except Exception as e:
        print(f"Ошибка при извлечении салатов: {e}")
        return []


def update_presentation_with_salads(presentation_path: str, salads: List[MenuItem], output_path: str) -> bool:
    """
    Обновляет презентацию, вставляя данные салатов во второй слайд.
    """
    try:
        # Копируем исходную презентацию
        shutil.copy2(presentation_path, output_path)
        
        # Открываем презентацию
        prs = Presentation(output_path)
        
        # Работаем со вторым слайдом (индекс 1)
        if len(prs.slides) < 2:
            return False
            
        slide = prs.slides[1]  # Второй слайд
        
        # Найдем таблицу на слайде
        table_shape = None
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                table_shape = shape
                break
                
        if table_shape is None:
            return False
            
        table = table_shape.table
        
        # Получаем количество строк в таблице
        total_rows = len(table.rows)
        
        # Очищаем все строки кроме первой (заголовки) и заполняем их салатами
        for i, salad in enumerate(salads):
            # Определяем индекс строки (начиная с 1, так как 0 - это заголовок)
            row_idx = i + 1
            
            # Проверяем, что строка существует в таблице
            if row_idx < total_rows:
                row = table.rows[row_idx]
                
                # Заполняем ячейки
                if len(row.cells) >= 3:
                    row.cells[0].text = salad.name
                    row.cells[1].text = salad.weight
                    row.cells[2].text = salad.price
            else:
                # Если строк не хватает, прерываем цикл
                break
        
        # Очищаем оставшиеся строки, если они есть
        salads_count = min(len(salads), total_rows - 1)  # -1 для заголовка
        for i in range(salads_count + 1, total_rows):
            if i < len(table.rows):
                row = table.rows[i]
                # Очищаем ячейки
                for j in range(len(row.cells)):
                    row.cells[j].text = ""
                
        # Сохраняем презентацию
        prs.save(output_path)
        return True
        
    except Exception as e:
        print(f"Ошибка при обновлении презентации: {e}")
        return False


def create_presentation_with_excel_data(template_path: str, excel_path: str, output_path: str) -> Tuple[bool, str]:
    """
    Создает презентацию с данными из Excel файла.
    
    Returns:
        Tuple[bool, str]: (успех, сообщение)
    """
    try:
        # Проверяем существование файлов
        if not Path(template_path).exists():
            return False, f"Шаблон презентации не найден: {template_path}"
            
        if not Path(excel_path).exists():
            return False, f"Excel файл не найден: {excel_path}"
        
        # Извлекаем салаты из Excel
        salads = extract_salads_from_excel(excel_path)
        if not salads:
            return False, "Не удалось найти данные салатов в Excel файле"
        
        # Обновляем презентацию
        success = update_presentation_with_salads(template_path, salads, output_path)
        
        if success:
            return True, f"Презентация создана успешно с {len(salads)} салатами"
        else:
            return False, "Ошибка при обновлении презентации"
            
    except Exception as e:
        return False, f"Ошибка: {str(e)}"
