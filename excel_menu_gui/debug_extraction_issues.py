#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Debug script to analyze issues with specific extraction functions
"""

import pandas as pd
from pathlib import Path
from presentation_handler import (
    extract_first_courses_from_excel, 
    extract_meat_dishes_from_excel,
    extract_fish_dishes_from_excel
)

def debug_extraction():
    """Debug the extraction issues"""
    excel_path = r"C:\Users\katya\Downloads\Telegram Desktop\18 сентября - четверг.xls"
    
    print("🔍 Отладка извлечения данных из Excel файла")
    print(f"Файл: {excel_path}")
    print("=" * 60)
    
    # 1. Проверяем первые блюда
    print("\n1️⃣ ПЕРВЫЕ БЛЮДА:")
    first_courses = extract_first_courses_from_excel(excel_path)
    print(f"Найдено: {len(first_courses)} первых блюд")
    for i, dish in enumerate(first_courses[:5]):  # Показываем первые 5
        print(f"   {i+1}. '{dish.name}' | '{dish.weight}' | '{dish.price}'")
    
    # 2. Проверяем мясные блюда
    print("\n2️⃣ МЯСНЫЕ БЛЮДА:")
    meat_dishes = extract_meat_dishes_from_excel(excel_path)
    print(f"Найдено: {len(meat_dishes)} мясных блюд")
    for i, dish in enumerate(meat_dishes[:5]):  # Показываем первые 5
        print(f"   {i+1}. '{dish.name}' | '{dish.weight}' | '{dish.price}'")
    
    # 3. Проверяем рыбные блюда
    print("\n3️⃣ РЫБНЫЕ БЛЮДА:")
    fish_dishes = extract_fish_dishes_from_excel(excel_path)
    print(f"Найдено: {len(fish_dishes)} рыбных блюд")
    for i, dish in enumerate(fish_dishes[:5]):  # Показываем первые 5
        print(f"   {i+1}. '{dish.name}' | '{dish.weight}' | '{dish.price}'")
    
    print("\n" + "=" * 60)
    
    # Дополнительная диагностика - прямое чтение Excel
    print("\n🔧 ДИАГНОСТИКА Excel файла:")
    try:
        df = pd.read_excel(excel_path, sheet_name=0, header=None, dtype=object)
        print(f"Размер файла: {len(df)} строк, {len(df.columns)} столбцов")
        
        # Ищем строки с ключевыми категориями
        categories_to_find = ['ПЕРВЫЕ БЛЮДА', 'БЛЮДА ИЗ МЯСА', 'БЛЮДА ИЗ РЫБЫ']
        
        for category in categories_to_find:
            print(f"\n🔍 Поиск категории '{category}':")
            found = False
            for i in range(len(df)):
                row_text = ''
                for j in range(len(df.columns)):
                    if pd.notna(df.iloc[i, j]):
                        row_text += str(df.iloc[i, j]) + ' '
                
                if category.upper() in row_text.upper():
                    print(f"   Строка {i+1}: {row_text.strip()}")
                    found = True
                    
                    # Показываем следующие 3 строки для анализа данных
                    for k in range(1, 4):
                        if i + k < len(df):
                            next_row_text = ''
                            for j in range(len(df.columns)):
                                if pd.notna(df.iloc[i + k, j]):
                                    next_row_text += f"[{chr(65+j)}]: {str(df.iloc[i + k, j])} "
                            if next_row_text.strip():
                                print(f"     Строка {i+k+1}: {next_row_text.strip()}")
                    break
            
            if not found:
                print(f"   ❌ Категория '{category}' не найдена")
    
    except Exception as e:
        print(f"❌ Ошибка при анализе Excel: {e}")

if __name__ == "__main__":
    debug_extraction()
