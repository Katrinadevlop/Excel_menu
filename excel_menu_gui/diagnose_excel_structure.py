#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Диагностика структуры Excel файла для понимания расположения данных
"""
import os
import sys
from pathlib import Path
import pandas as pd

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def diagnose_excel_structure():
    """Анализирует структуру Excel файла"""
    print("🔍 ДИАГНОСТИКА СТРУКТУРЫ EXCEL ФАЙЛА")
    print("=" * 70)
    
    excel_path = Path(r"C:\Users\katya\Downloads\Telegram Desktop\4 сентября - четверг (2).xls")
    
    if not excel_path.exists():
        print(f"❌ Excel файл не найден: {excel_path}")
        return
    
    try:
        # Выбираем лист
        xls = pd.ExcelFile(excel_path)
        sheet_name = None
        for nm in xls.sheet_names:
            if 'касс' in str(nm).strip().lower():
                sheet_name = nm
                break
        if sheet_name is None and xls.sheet_names:
            sheet_name = xls.sheet_names[0]

        print(f"📊 Листы в файле: {xls.sheet_names}")
        print(f"📄 Анализируем лист: {sheet_name}")
        
        # Читаем весь лист
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, dtype=object)
        
        def row_text(row) -> str:
            parts = []
            for v in row:
                if pd.notna(v):
                    parts.append(str(v))
            return ' '.join(parts).strip()
        
        print(f"\n📏 Размер данных: {len(df)} строк × {len(df.columns)} столбцов")
        
        # Ищем все категории и показываем структуру
        categories = {
            'ПЕРВЫЕ БЛЮДА': [],
            'БЛЮДА ИЗ МЯСА': [],
            'БЛЮДА ИЗ ПТИЦЫ': [],
            'БЛЮДА ИЗ РЫБЫ': [],
            'САЛАТЫ И ХОЛОДНЫЕ ЗАКУСКИ': [],
            'ГАРНИРЫ': []
        }
        
        for i in range(len(df)):
            row_content = row_text(df.iloc[i]).upper().replace('Ё', 'Е')
            
            for category in categories.keys():
                if category in row_content:
                    print(f"\n🎯 {category} (строка {i+1}):")
                    print(f"   Содержимое строки: {row_content}")
                    
                    # Показываем структуру следующих 10 строк
                    for j in range(1, min(11, len(df) - i)):
                        row_idx = i + j
                        if row_idx >= len(df):
                            break
                            
                        row = df.iloc[row_idx]
                        
                        # Показываем содержимое каждого столбца
                        row_data = []
                        for col_idx in range(min(10, len(df.columns))):  # Показываем первые 10 столбцов
                            if col_idx < len(row) and pd.notna(row.iloc[col_idx]):
                                cell_content = str(row.iloc[col_idx]).strip()
                                if cell_content:
                                    row_data.append(f"[{col_idx}]='{cell_content}'")
                        
                        if row_data:
                            print(f"   {j:2d}. {' | '.join(row_data)}")
                        else:
                            empty_count = 0
                            for next_j in range(j, min(j+3, len(df) - i)):
                                next_row_idx = i + next_j
                                if next_row_idx < len(df) and not row_text(df.iloc[next_row_idx]).strip():
                                    empty_count += 1
                                else:
                                    break
                            if empty_count >= 2:
                                print(f"   {j:2d}. [пустая строка - конец секции?]")
                                break
                    
                    categories[category] = list(range(i+1, min(i+11, len(df))))
                    break
        
        print(f"\n📋 ИТОГОВАЯ СТРУКТУРА:")
        for category, rows in categories.items():
            if rows:
                print(f"   {category}: строки {min(rows)}-{max(rows)}")
            
    except Exception as e:
        print(f"❌ Ошибка при анализе файла: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    diagnose_excel_structure()
