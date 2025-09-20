#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import os

def analyze_excel_structure():
    # Путь к основному файлу
    excel_path = "../8 сентября - понедельник (2).xls"
    
    if not os.path.exists(excel_path):
        print("❌ Файл не найден:", excel_path)
        return
    
    print(f"📊 Анализ структуры файла: {excel_path}")
    print("=" * 60)
    
    try:
        # Читаем файл
        xls = pd.ExcelFile(excel_path)
        print(f"📋 Листы в файле: {xls.sheet_names}")
        
        # Выбираем лист с "касс"
        sheet_name = None
        for nm in xls.sheet_names:
            if 'касс' in str(nm).strip().lower():
                sheet_name = nm
                break
        if sheet_name is None:
            sheet_name = xls.sheet_names[0]
        
        print(f"🔍 Анализируем лист: {sheet_name}")
        
        # Читаем лист
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, dtype=object)
        print(f"📏 Размер: {len(df)} строк, {len(df.columns)} столбцов")
        
        def row_text(row) -> str:
            parts = []
            for v in row:
                if pd.notna(v):
                    parts.append(str(v))
            return ' '.join(parts).strip()
        
        # Ищем секцию рыбных блюд
        fish_section_start = None
        fish_section_end = None
        
        for i in range(len(df)):
            row_content = row_text(df.iloc[i]).upper().replace('Ё', 'Е')
            
            if fish_section_start is None and 'БЛЮДА ИЗ РЫБЫ' in row_content:
                fish_section_start = i
                print(f"\\n🐟 Найдена секция рыбных блюд в строке {i + 1}: {row_content}")
                continue
            
            if (fish_section_start is not None and fish_section_end is None and 
                'ГАРНИРЫ' in row_content):
                fish_section_end = i
                print(f"🔚 Конец секции рыбных блюд в строке {i + 1}: {row_content}")
                break
        
        if fish_section_start is None:
            print("❌ Секция 'БЛЮДА ИЗ РЫБЫ' не найдена")
            return
        
        if fish_section_end is None:
            fish_section_end = min(fish_section_start + 10, len(df))
        
        print(f"\\n📋 Детальный анализ строк с {fish_section_start + 1} по {fish_section_end}:")
        print("=" * 80)
        
        # Анализируем каждую строку в секции
        for i in range(fish_section_start, fish_section_end):
            if i >= len(df):
                break
                
            row = df.iloc[i]
            print(f"\\nСтрока {i + 1}:")
            
            for j, cell_value in enumerate(row):
                if pd.notna(cell_value) and str(cell_value).strip():
                    column_letter = chr(65 + j)  # A, B, C, D, E, F, G...
                    print(f"  {column_letter} (индекс {j}): '{cell_value}'")
        
        print(f"\\n🎯 Попробуем найти ваши блюда:")
        expected_dishes = [
            "Окунь жареный (тушка)",
            "Котлета по-приморски", 
            "Треска с сыром и овощами",
            "Филе форели гриль"
        ]
        
        for expected in expected_dishes:
            found = False
            for i in range(len(df)):
                row = df.iloc[i]
                for j, cell_value in enumerate(row):
                    if pd.notna(cell_value):
                        cell_str = str(cell_value).strip()
                        if expected.lower() in cell_str.lower():
                            column_letter = chr(65 + j)
                            print(f"  ✓ '{expected}' найден в столбце {column_letter} (индекс {j}), строка {i + 1}")
                            found = True
                            break
                if found:
                    break
            if not found:
                print(f"  ❌ '{expected}' не найден")
                
    except Exception as e:
        print(f"❌ Ошибка: {e}")

if __name__ == "__main__":
    analyze_excel_structure()
    input("\\nНажмите Enter для выхода...")
