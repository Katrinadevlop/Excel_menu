#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import sys
import os

def analyze_file(file_path):
    """Подробный анализ файла 18 сентября"""
    try:
        print(f"=== АНАЛИЗ ФАЙЛА: {os.path.basename(file_path)} ===\n")
        
        # Читаем Excel файл
        xls = pd.ExcelFile(file_path)
        print(f"📋 Листы в файле: {xls.sheet_names}")
        
        # Выбираем лист для анализа (приоритет листу с "касс")
        sheet_name = None
        for nm in xls.sheet_names:
            if 'касс' in str(nm).strip().lower():
                sheet_name = nm
                break
        
        if sheet_name is None and xls.sheet_names:
            sheet_name = xls.sheet_names[0]
            
        print(f"🎯 Используемый лист: {sheet_name}")
        
        # Читаем данные
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, dtype=object)
        print(f"📊 Размер данных: {len(df)} строк, {len(df.columns)} столбцов")
        
        def row_text(row) -> str:
            parts = []
            for v in row:
                if pd.notna(v):
                    parts.append(str(v))
            return ' '.join(parts).strip()
        
        # Ищем ключевые категории в первых 100 строках
        categories_found = {}
        categories_to_find = {
            'САЛАТЫ И ХОЛОДНЫЕ ЗАКУСКИ': ['САЛАТ', 'ХОЛОДН', 'ЗАКУСК'],
            'ПЕРВЫЕ БЛЮДА': ['ПЕРВЫЕ', 'БЛЮДА'],
            'БЛЮДА ИЗ МЯСА': ['БЛЮДА', 'МЯСА'],
            'БЛЮДА ИЗ ПТИЦЫ': ['БЛЮДА', 'ПТИЦЫ'],
            'БЛЮДА ИЗ РЫБЫ': ['БЛЮДА', 'РЫБЫ'],
            'ГАРНИРЫ': ['ГАРНИРЫ', 'ГАРНИР']
        }
        
        print("\n🔍 ПОИСК КАТЕГОРИЙ:")
        for i in range(min(100, len(df))):
            row_content = row_text(df.iloc[i]).upper().replace('Ё', 'Е')
            if not row_content.strip():
                continue
                
            for category_name, keywords in categories_to_find.items():
                if category_name not in categories_found:
                    if any(kw in row_content for kw in keywords if len(kw) > 2):
                        categories_found[category_name] = i
                        print(f"  ✅ {category_name}: строка {i + 1}")
                        print(f"      Содержимое: {row_content[:100]}")
                        
                        # Показываем распределение по столбцам для этой строки
                        print(f"      Распределение по столбцам:")
                        for j in range(min(8, len(df.columns))):
                            if pd.notna(df.iloc[i, j]):
                                cell_content = str(df.iloc[i, j]).strip()
                                if cell_content:
                                    col_letter = chr(65 + j)  # A, B, C, D, E, F, G, H
                                    print(f"        {col_letter}: {cell_content}")
                        print()
        
        if not categories_found:
            print("  ❌ Категории не найдены!")
            return
            
        # Анализируем блюда из птицы подробно
        if 'БЛЮДА ИЗ ПТИЦЫ' in categories_found:
            print("🐔 ПОДРОБНЫЙ АНАЛИЗ БЛЮД ИЗ ПТИЦЫ:")
            start_row = categories_found['БЛЮДА ИЗ ПТИЦЫ']
            
            # Ищем конец секции (блюда из рыбы или другая категория)
            end_row = len(df)
            for category_name, row_idx in categories_found.items():
                if category_name in ['БЛЮДА ИЗ РЫБЫ', 'ГАРНИРЫ'] and row_idx > start_row:
                    end_row = min(end_row, row_idx)
            
            # Ограничиваем анализ разумными пределами
            end_row = min(end_row, start_row + 50)
            
            print(f"  Анализируем строки {start_row + 1} - {end_row}")
            
            dishes_found = 0
            for i in range(start_row + 1, end_row):
                if i >= len(df):
                    break
                    
                row = df.iloc[i]
                row_content = row_text(row)
                
                # Пропускаем пустые строки
                if not row_content.strip():
                    continue
                    
                # Пропускаем заголовки
                if row_content.isupper() and len(row_content) > 10:
                    continue
                
                print(f"\n  📝 Строка {i + 1}: {row_content}")
                print(f"      Распределение по столбцам:")
                
                # Анализируем каждый столбец
                row_data = {}
                for j in range(min(8, len(df.columns))):
                    if pd.notna(df.iloc[i, j]):
                        cell_content = str(df.iloc[i, j]).strip()
                        if cell_content:
                            col_letter = chr(65 + j)  # A, B, C, D, E, F, G, H
                            row_data[col_letter] = cell_content
                            
                            # Определяем тип данных
                            data_type = "неизвестно"
                            if cell_content.replace('.', '').replace(',', '').isdigit():
                                data_type = "цена?"
                            elif any(unit in cell_content.lower() for unit in ['г', 'мл', 'л', 'кг', 'шт']):
                                data_type = "вес?"
                            elif not cell_content.isupper() and len(cell_content) > 3:
                                data_type = "название?"
                            
                            print(f"        {col_letter}: {cell_content} ({data_type})")
                
                # Проверяем, есть ли потенциальное название блюда
                potential_dishes = []
                for col, content in row_data.items():
                    if (not content.isupper() and 
                        len(content) > 3 and 
                        not content.replace('.', '').replace(',', '').isdigit() and
                        not any(unit in content.lower() for unit in ['г', 'мл', 'л', 'кг', 'шт'])):
                        potential_dishes.append((col, content))
                
                if potential_dishes:
                    dishes_found += 1
                    print(f"      🍽️ Потенциальные блюда: {potential_dishes}")
                    
                    # Анализируем старый метод (столбец D) vs новый метод (столбец E)
                    old_method_dish = row_data.get('D', '')  # Старый метод
                    new_method_dish = row_data.get('E', '')  # Новый метод
                    
                    print(f"      🔄 Сравнение методов:")
                    print(f"         Старый (столбец D): '{old_method_dish}'")
                    print(f"         Новый (столбец E): '{new_method_dish}'")
                    
            print(f"\n  📊 Итого найдено потенциальных блюд: {dishes_found}")
        
        # Показываем общую структуру данных
        print(f"\n📋 ОБЩАЯ СТРУКТУРА ДАННЫХ:")
        print(f"   Левая часть (A-C): завтраки, дополнительные блюда")
        print(f"   Правая часть (E-G): основные категории блюд")
        
        # Анализируем несколько строк с данными в разных частях
        print(f"\n🔍 ОБРАЗЦЫ ДАННЫХ ИЗ РАЗНЫХ ЧАСТЕЙ:")
        
        sample_rows = []
        for i in range(min(100, len(df))):
            row = df.iloc[i]
            # Ищем строки с данными и в левой, и в правой части
            left_data = any(pd.notna(row.iloc[j]) and str(row.iloc[j]).strip() for j in range(min(3, len(row))))
            right_data = any(pd.notna(row.iloc[j]) and str(row.iloc[j]).strip() for j in range(4, min(7, len(row))))
            
            if left_data and right_data and len(sample_rows) < 5:
                sample_rows.append(i)
        
        for i in sample_rows:
            row = df.iloc[i]
            print(f"\n  Строка {i + 1}:")
            print(f"    Левая часть (A-C): ", end="")
            for j in range(min(3, len(row))):
                if pd.notna(row.iloc[j]):
                    content = str(row.iloc[j]).strip()[:20]
                    print(f"{chr(65+j)}:'{content}' ", end="")
            print()
            
            print(f"    Правая часть (E-G): ", end="")
            for j in range(4, min(7, len(row))):
                if pd.notna(row.iloc[j]):
                    content = str(row.iloc[j]).strip()[:20]
                    print(f"{chr(65+j)}:'{content}' ", end="")
            print()
            
    except Exception as e:
        print(f"❌ Ошибка при анализе файла: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    file_path = r"C:\Users\katya\Downloads\Telegram Desktop\18 сентября - четверг.xls"
    
    if not os.path.exists(file_path):
        print(f"❌ Файл не найден: {file_path}")
        sys.exit(1)
    
    analyze_file(file_path)
