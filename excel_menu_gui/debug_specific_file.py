#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Диагностика конкретного файла: 4 сентября - четверг (2).xls
"""
import os
import sys
from pathlib import Path
import pandas as pd

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from presentation_handler import extract_fish_dishes_from_column_e, MenuItem

def analyze_problematic_file():
    """Анализирует проблемный файл"""
    excel_path = r"C:\Users\katya\Downloads\Telegram Desktop\4 сентября - четверг (2).xls"
    
    print("🔍 АНАЛИЗ ПРОБЛЕМНОГО ФАЙЛА")
    print(f"Файл: {Path(excel_path).name}")
    print("=" * 70)
    
    if not Path(excel_path).exists():
        print(f"❌ Файл не найден: {excel_path}")
        return
    
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
        if sheet_name is None and xls.sheet_names:
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
        
        print("\n🔍 ПОИСК СЕКЦИИ РЫБНЫХ БЛЮД:")
        for i in range(min(50, len(df))):
            row_content = row_text(df.iloc[i]).upper().replace('Ё', 'Е')
            
            if 'РЫБА' in row_content or 'РЫБН' in row_content:
                print(f"   Строка {i + 1}: {row_content}")
            
            if fish_section_start is None and 'БЛЮДА ИЗ РЫБЫ' in row_content:
                fish_section_start = i
                print(f"🎯 Найдена секция рыбных блюд в строке {i + 1}: {row_content}")
                continue
            
            if (fish_section_start is not None and fish_section_end is None and 
                ('ГАРНИРЫ' in row_content or 'ГАРНИР' in row_content)):
                fish_section_end = i
                print(f"🔚 Конец секции рыбных блюд в строке {i + 1}: {row_content}")
                break
        
        if fish_section_start is None:
            print("❌ Секция 'БЛЮДА ИЗ РЫБЫ' не найдена в первых 50 строках")
            
            # Попробуем найти любые упоминания рыбы
            print("\n🔍 ПОИСК ЛЮБЫХ УПОМИНАНИЙ РЫБЫ:")
            for i in range(len(df)):
                row_content = row_text(df.iloc[i])
                if any(word in row_content.lower() for word in ['рыб', 'окун', 'треска', 'форел', 'минтай']):
                    print(f"   Строка {i + 1}: {row_content[:100]}")
            return
        
        if fish_section_end is None:
            fish_section_end = min(fish_section_start + 15, len(df))
        
        print(f"\n📋 ДЕТАЛЬНЫЙ АНАЛИЗ СТРОК С {fish_section_start + 1} ПО {fish_section_end}:")
        print("=" * 80)
        
        # Анализируем каждую строку в секции
        for i in range(fish_section_start, fish_section_end):
            if i >= len(df):
                break
                
            row = df.iloc[i]
            print(f"\nСТРОКА {i + 1}:")
            
            row_cells = []
            for j, cell_value in enumerate(row):
                if pd.notna(cell_value) and str(cell_value).strip():
                    column_letter = chr(65 + j)  # A, B, C, D, E, F, G...
                    cell_str = str(cell_value).strip()
                    row_cells.append(f"{column_letter}({j}): '{cell_str}'")
            
            if row_cells:
                print(f"   {' | '.join(row_cells)}")
            else:
                print("   [пустая строка]")
        
        print(f"\n🧪 ТЕСТИРУЕМ ТЕКУЩУЮ ФУНКЦИЮ ИЗВЛЕЧЕНИЯ:")
        fish_dishes = extract_fish_dishes_from_column_e(excel_path)
        
        print(f"✅ Результат: найдено {len(fish_dishes)} рыбных блюд:")
        for i, dish in enumerate(fish_dishes, 1):
            print(f"   {i}. Название: '{dish.name}'")
            print(f"      Вес:      '{dish.weight}'")
            print(f"      Цена:     '{dish.price}'")
            print()
        
        if len(fish_dishes) == 0 or any(not dish.name for dish in fish_dishes):
            print("⚠️  ПРОБЛЕМА ОБНАРУЖЕНА!")
            print("\n🔧 ДИАГНОСТИКА ПРОБЛЕМЫ:")
            
            # Пробуем найти данные в разных столбцах
            print("\n📍 ПОИСК РЫБНЫХ БЛЮД ВО ВСЕХ СТОЛБЦАХ:")
            fish_keywords = ['окун', 'треска', 'форел', 'минтай', 'котлет', 'рыбн']
            
            for i in range(fish_section_start + 1, fish_section_end):
                if i >= len(df):
                    break
                
                row = df.iloc[i]
                found_dish = False
                
                for j, cell_value in enumerate(row):
                    if pd.notna(cell_value):
                        cell_str = str(cell_value).strip().lower()
                        if any(keyword in cell_str for keyword in fish_keywords):
                            column_letter = chr(65 + j)
                            print(f"   🐟 Строка {i+1}, столбец {column_letter}({j}): '{cell_value}'")
                            
                            # Показываем соседние ячейки
                            neighbors = []
                            for offset in [-2, -1, 1, 2]:
                                neighbor_col = j + offset
                                if 0 <= neighbor_col < len(row) and pd.notna(row.iloc[neighbor_col]):
                                    neighbor_letter = chr(65 + neighbor_col)
                                    neighbors.append(f"{neighbor_letter}: '{row.iloc[neighbor_col]}'")
                            
                            if neighbors:
                                print(f"      Соседние ячейки: {' | '.join(neighbors)}")
                            
                            found_dish = True
                            break
                
                if not found_dish:
                    # Показываем всю строку если не нашли явных рыбных блюд
                    row_content = row_text(row)
                    if row_content.strip():
                        print(f"   ? Строка {i+1}: {row_content}")
                        
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    analyze_problematic_file()
