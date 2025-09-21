#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Анализ структуры разных Excel файлов для понимания различий в извлечении данных
"""

import pandas as pd
import os
from pathlib import Path
import re

def analyze_excel_structure(excel_path: str):
    """Анализирует структуру Excel файла"""
    try:
        print(f"\n{'='*80}")
        print(f"📄 АНАЛИЗ ФАЙЛА: {Path(excel_path).name}")
        print(f"{'='*80}")
        
        # Читаем все листы
        xls = pd.ExcelFile(excel_path)
        print(f"📊 Листы в файле: {xls.sheet_names}")
        
        # Определяем основной лист
        sheet_name = None
        for nm in xls.sheet_names:
            if 'касс' in str(nm).strip().lower():
                sheet_name = nm
                break
        if sheet_name is None and xls.sheet_names:
            sheet_name = xls.sheet_names[0]
            
        print(f"📋 Используемый лист: '{sheet_name}'")
        
        # Читаем лист
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, dtype=object)
        print(f"📊 Размер: {len(df)} строк, {len(df.columns)} столбцов")
        
        def row_text(row) -> str:
            parts = []
            for v in row:
                if pd.notna(v):
                    parts.append(str(v))
            return ' '.join(parts).strip()
        
        # Ищем ключевые категории и их позиции
        categories = {
            'ПЕРВЫЕ БЛЮДА': None,
            'БЛЮДА ИЗ МЯСА': None,
            'БЛЮДА ИЗ ПТИЦЫ': None, 
            'БЛЮДА ИЗ РЫБЫ': None,
            'САЛАТЫ': None,
            'ГАРНИРЫ': None
        }
        
        print(f"\n🔍 ПОИСК КАТЕГОРИЙ:")
        for i in range(min(50, len(df))):
            row_content = row_text(df.iloc[i]).upper().replace('Ё', 'Е')
            
            for category in categories.keys():
                if categories[category] is None:
                    if category in row_content:
                        categories[category] = i
                        print(f"   {category}: строка {i+1}")
                        
                        # Показываем содержимое этой строки по столбцам
                        print(f"      Содержимое по столбцам:")
                        for j, cell in enumerate(df.iloc[i]):
                            if pd.notna(cell) and str(cell).strip():
                                print(f"         Столбец {j+1}: '{str(cell).strip()}'")
        
        # Анализируем структуру данных в каждой категории
        print(f"\n📋 АНАЛИЗ СТРУКТУРЫ ДАННЫХ:")
        
        for category, start_row in categories.items():
            if start_row is not None:
                print(f"\n   📌 {category} (строка {start_row+1}):")
                
                # Показываем несколько строк после заголовка категории
                for i in range(1, min(6, len(df) - start_row)):  # До 5 строк после заголовка
                    row_idx = start_row + i
                    if row_idx < len(df):
                        row = df.iloc[row_idx]
                        row_content = row_text(row)
                        
                        if row_content.strip():
                            # Проверяем, не является ли это следующей категорией
                            is_next_category = any(cat in row_content.upper().replace('Ё', 'Е') 
                                                 for cat in categories.keys() 
                                                 if cat != category)
                            
                            if is_next_category:
                                break
                                
                            print(f"      Строка {row_idx+1}: {row_content[:100]}")
                            
                            # Показываем детальное содержимое по столбцам
                            has_data = False
                            for j, cell in enumerate(row):
                                if pd.notna(cell) and str(cell).strip():
                                    cell_text = str(cell).strip()
                                    if not cell_text.isupper() and len(cell_text) > 2:
                                        has_data = True
                                        
                            if has_data:
                                print(f"         Детали:")
                                for j, cell in enumerate(row):
                                    if pd.notna(cell) and str(cell).strip():
                                        cell_text = str(cell).strip()
                                        # Определяем тип данных
                                        data_type = "неопределен"
                                        if re.search(r'\d+.*?(г|гр|грамм|мл|л|кг|шт)', cell_text, re.IGNORECASE):
                                            data_type = "ВЕС"
                                        elif re.match(r'^\d+([.,]\d+)?\s*(руб|₽|р\.?)?$', cell_text):
                                            data_type = "ЦЕНА"
                                        elif not cell_text.isupper() and len(cell_text) > 3:
                                            data_type = "НАЗВАНИЕ"
                                            
                                        print(f"            Столбец {j+1} ({chr(65+j)}): '{cell_text}' -> {data_type}")
                            break
        
        print(f"\n{'='*80}")
        
    except Exception as e:
        print(f"❌ Ошибка при анализе файла {excel_path}: {e}")

def main():
    """Основная функция"""
    print("🔍 АНАЛИЗ СТРУКТУРЫ EXCEL ФАЙЛОВ")
    print("="*80)
    
    # Список файлов для анализа
    download_dir = r"C:\Users\katya\Downloads\Telegram Desktop"
    
    files_to_analyze = [
        "15 сентября - понедельник.xlsx",
        "17 сентябя-среда.xlsx", 
        "11 сентября - четверг.xlsx"
    ]
    
    for filename in files_to_analyze:
        file_path = os.path.join(download_dir, filename)
        if os.path.exists(file_path):
            analyze_excel_structure(file_path)
        else:
            print(f"❌ Файл не найден: {filename}")

if __name__ == "__main__":
    main()
