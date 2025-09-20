#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Детальное изучение структуры Excel файла для поиска настоящих рыбных блюд.
"""

import pandas as pd
import os
from pathlib import Path

def analyze_excel_structure(excel_path):
    """Анализирует структуру Excel файла."""
    
    if not os.path.exists(excel_path):
        print(f"❌ Файл не найден: {excel_path}")
        return
    
    print(f"📂 Анализируем файл: {Path(excel_path).name}")
    print("=" * 80)
    
    try:
        # Читаем все листы
        xls = pd.ExcelFile(excel_path)
        print(f"📋 Листы в файле: {xls.sheet_names}")
        
        # Выбираем основной лист (с "касс" или первый)
        sheet_name = None
        for nm in xls.sheet_names:
            if 'касс' in str(nm).strip().lower():
                sheet_name = nm
                break
        if sheet_name is None and xls.sheet_names:
            sheet_name = xls.sheet_names[0]
        
        print(f"🎯 Используем лист: {sheet_name}")
        
        # Читаем данные
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, dtype=object)
        print(f"📊 Размер листа: {len(df)} строк, {len(df.columns)} столбцов")
        
        def row_text(row) -> str:
            parts = []
            for v in row:
                if pd.notna(v):
                    parts.append(str(v))
            return ' '.join(parts).strip()
        
        # Ищем заголовок "БЛЮДА ИЗ РЫБЫ"
        fish_header_row = None
        fish_end_row = None
        
        print(f"\n🔍 ПОИСК РАЗДЕЛА 'БЛЮДА ИЗ РЫБЫ'...")
        print("-" * 60)
        
        for i in range(len(df)):
            row_content = row_text(df.iloc[i]).upper().replace('Ё', 'Е')
            
            # Ищем заголовок рыбных блюд
            if fish_header_row is None:
                if 'БЛЮДА ИЗ РЫБЫ' in row_content or 'РЫБНЫЕ БЛЮДА' in row_content:
                    fish_header_row = i
                    print(f"✅ НАЙДЕН заголовок в строке {i + 1}: {row_content}")
                    continue
            
            # Ищем конец секции
            if fish_header_row is not None and fish_end_row is None:
                if any(category in row_content for category in [
                    'ГАРНИРЫ', 'НАПИТКИ', 'ДЕСЕРТЫ', 'САЛАТЫ', 'ЗАКУСКИ'
                ]):
                    fish_end_row = i
                    print(f"✅ НАЙДЕН конец секции в строке {i + 1}: {row_content}")
                    break
        
        if fish_header_row is None:
            print("❌ Заголовок 'БЛЮДА ИЗ РЫБЫ' НЕ НАЙДЕН!")
            print("🔍 Показываю все строки, чтобы найти рыбные блюда вручную:")
            print("-" * 60)
            
            for i in range(min(50, len(df))):
                content = row_text(df.iloc[i])
                if content.strip():
                    print(f"Строка {i+1:>2}: {content}")
            return
        
        if fish_end_row is None:
            fish_end_row = min(fish_header_row + 10, len(df))
            print(f"⚠️  Конец секции не найден, берем до строки {fish_end_row}")
        
        print(f"\n📋 СОДЕРЖИМОЕ РАЗДЕЛА РЫБНЫХ БЛЮД (строки {fish_header_row + 1} - {fish_end_row}):")
        print("=" * 80)
        
        # Показываем все строки в разделе рыбных блюд
        for i in range(fish_header_row, fish_end_row):
            if i >= len(df):
                break
            
            row = df.iloc[i]
            row_content = row_text(row)
            
            print(f"\nСТРОКА {i + 1}:")
            print(f"  Полное содержимое: '{row_content}'")
            
            # Показываем содержимое каждого столбца
            for col_idx in range(len(df.columns)):
                if pd.notna(df.iloc[i, col_idx]):
                    cell_content = str(df.iloc[i, col_idx]).strip()
                    if cell_content:
                        column_letter = chr(65 + col_idx)  # A, B, C, D, E, F, G...
                        print(f"    Столбец {column_letter}: '{cell_content}'")
        
        print(f"\n🤔 АНАЛИЗ:")
        print("- Показаны ВСЕ строки из раздела 'БЛЮДА ИЗ РЫБЫ'")
        print("- Если здесь нет настоящих рыбных блюд, значит они в другом месте")
        print("- Возможно, нужно искать в других листах или разделах")
        
    except Exception as e:
        print(f"❌ ОШИБКА при анализе: {e}")
        import traceback
        traceback.print_exc()

def main():
    """Главная функция."""
    
    # Ищем реальные файлы меню
    real_menu_files = [
        r"C:\Users\katya\Desktop\menurepit\5  сентября - пятница.xlsx",
        r"C:\Users\katya\Desktop\menurepit\01  августа - пятница.xls",
        r"C:\Users\katya\Desktop\menurepit\8 сентября - понедельник (2).xls"
    ]
    
    excel_path = None
    for file_path in real_menu_files:
        if os.path.exists(file_path):
            excel_path = file_path
            break
    
    if not excel_path:
        print("❌ Реальные файлы меню не найдены!")
        return
    
    analyze_excel_structure(excel_path)

if __name__ == "__main__":
    main()
