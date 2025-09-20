#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import re
from pathlib import Path
from presentation_handler import extract_fish_dishes_from_column_e

def test_fish_extraction():
    """Тестирует извлечение рыбных блюд с подробным выводом"""
    
    # Укажите путь к вашему Excel файлу
    excel_path = input("Введите путь к Excel файлу: ").strip().strip('"')
    
    if not Path(excel_path).exists():
        print(f"❌ Файл не найден: {excel_path}")
        return
    
    print(f"📂 Открываем файл: {excel_path}")
    
    try:
        # Читаем файл и показываем структуру
        xls = pd.ExcelFile(excel_path)
        print(f"📋 Листы в файле: {xls.sheet_names}")
        
        # Выбираем лист
        sheet_name = None
        for nm in xls.sheet_names:
            if 'касс' in str(nm).strip().lower():
                sheet_name = nm
                break
        if sheet_name is None and xls.sheet_names:
            sheet_name = xls.sheet_names[0]
            
        print(f"📄 Используем лист: {sheet_name}")
        
        # Читаем данные
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, dtype=object)
        print(f"📊 Размер данных: {len(df)} строк, {len(df.columns)} столбцов")
        
        def row_text(row) -> str:
            parts = []
            for v in row:
                if pd.notna(v):
                    parts.append(str(v))
            return ' '.join(parts).strip()
        
        # Ищем заголовок "БЛЮДА ИЗ РЫБЫ"
        print("\n🔍 Ищем заголовок 'БЛЮДА ИЗ РЫБЫ':")
        fish_header_found = False
        for i in range(min(50, len(df))):
            content = row_text(df.iloc[i]).upper().replace('Ё', 'Е')
            if content.strip():
                print(f"  Строка {i+1}: {content}")
                if 'БЛЮДА ИЗ РЫБЫ' in content or ('РЫБН' in content and 'БЛЮДА' in content):
                    print(f"  ✅ НАЙДЕН заголовок рыбных блюд в строке {i+1}!")
                    fish_header_found = True
                    
                    # Показываем следующие 10 строк после заголовка
                    print(f"\n📝 Следующие 10 строк после заголовка:")
                    for j in range(i+1, min(i+11, len(df))):
                        if j < len(df):
                            row_content = row_text(df.iloc[j])
                            if row_content.strip():
                                print(f"    Строка {j+1}: {row_content}")
                                
                                # Показываем содержимое по столбцам
                                print(f"      Столбцы:")
                                for col_idx in range(len(df.columns)):
                                    if pd.notna(df.iloc[j, col_idx]):
                                        cell_val = str(df.iloc[j, col_idx]).strip()
                                        if cell_val:
                                            print(f"        Столбец {col_idx+1}: '{cell_val}'")
                    break
        
        if not fish_header_found:
            print("❌ Заголовок 'БЛЮДА ИЗ РЫБЫ' НЕ НАЙДЕН!")
            print("\n🔍 Возможные варианты в файле:")
            for i in range(min(30, len(df))):
                content = row_text(df.iloc[i]).upper()
                if 'РЫБ' in content:
                    print(f"  Строка {i+1}: {content}")
            return
        
        # Теперь тестируем основную функцию
        print("\n🧪 Тестируем функцию extract_fish_dishes_from_column_e:")
        fish_dishes = extract_fish_dishes_from_column_e(excel_path)
        
        print(f"\n📊 Результат: найдено {len(fish_dishes)} рыбных блюд")
        
        if fish_dishes:
            print("\n🐟 Найденные рыбные блюда:")
            for i, dish in enumerate(fish_dishes, 1):
                print(f"  {i}. Название: '{dish.name}'")
                print(f"     Вес: '{dish.weight}'")
                print(f"     Цена: '{dish.price}'")
                print()
        else:
            print("❌ Рыбные блюда не найдены!")
            
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_fish_extraction()
