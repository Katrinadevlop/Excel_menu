#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Поиск настоящих рыбных блюд во всех листах Excel файла.
"""

import pandas as pd
import os
from pathlib import Path

def find_fish_in_all_sheets(excel_path):
    """Ищет рыбные блюда во всех листах Excel файла."""
    
    if not os.path.exists(excel_path):
        print(f"❌ Файл не найден: {excel_path}")
        return
    
    print(f"📂 Ищем рыбные блюда в файле: {Path(excel_path).name}")
    print("=" * 80)
    
    try:
        xls = pd.ExcelFile(excel_path)
        print(f"📋 Найдено листов: {len(xls.sheet_names)}")
        
        def row_text(row) -> str:
            parts = []
            for v in row:
                if pd.notna(v):
                    parts.append(str(v))
            return ' '.join(parts).strip()
        
        # Проверяем каждый лист
        for sheet_idx, sheet_name in enumerate(xls.sheet_names, 1):
            print(f"\n🔍 ЛИСТ {sheet_idx}: '{sheet_name}'")
            print("-" * 60)
            
            try:
                df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, dtype=object)
                print(f"📊 Размер: {len(df)} строк, {len(df.columns)} столбцов")
                
                # Ищем слова связанные с рыбой
                fish_keywords = [
                    'рыб', 'форел', 'семг', 'лосос', 'треск', 'хек', 'судак', 
                    'карп', 'щука', 'окун', 'сом', 'минтай', 'пангасиус',
                    'котлет', 'филе'  # часто используется с рыбой
                ]
                
                found_fish_rows = []
                
                for i in range(len(df)):
                    row_content = row_text(df.iloc[i]).lower()
                    
                    # Проверяем наличие рыбных слов
                    for keyword in fish_keywords:
                        if keyword in row_content and len(row_content.strip()) > 5:
                            found_fish_rows.append((i, row_content))
                            break
                
                if found_fish_rows:
                    print(f"🐟 Найдено {len(found_fish_rows)} строк с рыбными блюдами:")
                    
                    for row_idx, content in found_fish_rows:
                        print(f"  Строка {row_idx + 1}: {content}")
                        
                        # Показываем детали этой строки по столбцам
                        row = df.iloc[row_idx]
                        for col_idx in range(len(df.columns)):
                            if pd.notna(df.iloc[row_idx, col_idx]):
                                cell_content = str(df.iloc[row_idx, col_idx]).strip()
                                if cell_content:
                                    column_letter = chr(65 + col_idx)
                                    print(f"    {column_letter}: '{cell_content}'")
                        print()
                else:
                    print("❌ Рыбных блюд не найдено")
                    
            except Exception as e:
                print(f"❌ Ошибка при чтении листа: {e}")
    
    except Exception as e:
        print(f"❌ ОШИБКА: {e}")

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
    
    find_fish_in_all_sheets(excel_path)
    
    print(f"\n💡 РЕКОМЕНДАЦИИ:")
    print("1. Проверьте, в каком листе находятся НАСТОЯЩИЕ рыбные блюда")
    print("2. Убедитесь, что данные в правильном формате (название | вес | цена)")
    print("3. Если рыбных блюд нет - возможно файл содержит только тестовые данные")

if __name__ == "__main__":
    main()
