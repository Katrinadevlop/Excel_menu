#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
from pathlib import Path

# Добавляем текущую директорию в путь для импорта
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from presentation_handler import extract_fish_dishes_from_excel, extract_fish_dishes_by_range

def test_fish_extraction():
    """Тестирует извлечение блюд из рыбы из Excel файла."""
    
    # Используем файл из папки templates
    excel_path = r"C:\Users\katya\Desktop\menurepit\excel_menu_gui\templates\Шаблон меню пример.xlsx"
    
    if not Path(excel_path).exists():
        print("❌ Excel файл не найден")
        print("Проверяем папку templates:")
        templates_dir = Path(r"C:\Users\katya\Desktop\menurepit\excel_menu_gui\templates")
        if templates_dir.exists():
            for f in templates_dir.glob("*.xlsx"):
                print(f"  - {f.name}")
        return
    
    print(f"📂 Используется файл: {Path(excel_path).name}")
    
    print("=" * 80)
    print("ТЕСТИРОВАНИЕ ИЗВЛЕЧЕНИЯ БЛЮД ИЗ РЫБЫ")
    print("=" * 80)
    
    # Тест 1: Извлечение по точному диапазону
    print("\n📋 Тест 1: Извлечение по точному диапазону (extract_fish_dishes_by_range)")
    print("-" * 40)
    dishes_by_range = extract_fish_dishes_by_range(excel_path)
    
    if dishes_by_range:
        print(f"✅ Найдено {len(dishes_by_range)} блюд по диапазону:")
        for i, dish in enumerate(dishes_by_range[:10], 1):  # Показываем первые 10
            print(f"  {i}. {dish.name}")
            print(f"     Вес: {dish.weight if dish.weight else 'не указан'}")
            print(f"     Цена: {dish.price if dish.price else 'не указана'}")
            print()
    else:
        print("❌ Блюда не найдены по диапазону")
    
    # Тест 2: Извлечение через основную функцию
    print("\n📋 Тест 2: Извлечение через основную функцию (extract_fish_dishes_from_excel)")
    print("-" * 40)
    dishes_main = extract_fish_dishes_from_excel(excel_path)
    
    if dishes_main:
        print(f"✅ Найдено {len(dishes_main)} блюд через основную функцию:")
        for i, dish in enumerate(dishes_main[:10], 1):  # Показываем первые 10
            print(f"  {i}. {dish.name}")
            print(f"     Вес: {dish.weight if dish.weight else 'не указан'}")
            print(f"     Цена: {dish.price if dish.price else 'не указана'}")
            print()
    else:
        print("❌ Блюда не найдены через основную функцию")
    
    # Сравнение результатов
    print("\n📊 СРАВНЕНИЕ РЕЗУЛЬТАТОВ:")
    print("-" * 40)
    print(f"По диапазону: {len(dishes_by_range)} блюд")
    print(f"Основная функция: {len(dishes_main)} блюд")
    
    if len(dishes_main) > 0:
        print("\n✅ Рыбные блюда успешно извлекаются из файла!")
    else:
        print("\n⚠️ Проблема с извлечением рыбных блюд - проверьте структуру Excel файла")
    
    # Дополнительная диагностика
    if len(dishes_main) == 0:
        print("\n🔍 ДИАГНОСТИКА:")
        print("-" * 40)
        import pandas as pd
        
        try:
            xls = pd.ExcelFile(excel_path)
            print(f"Листы в файле: {xls.sheet_names}")
            
            # Ищем лист с 'касс'
            sheet_name = None
            for nm in xls.sheet_names:
                if 'касс' in str(nm).strip().lower():
                    sheet_name = nm
                    break
            if sheet_name is None and xls.sheet_names:
                sheet_name = xls.sheet_names[0]
            
            print(f"Используется лист: {sheet_name}")
            
            df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, dtype=object)
            
            print(f"Размер листа: {len(df)} строк, {len(df.columns)} столбцов")
            
            # Ищем строки с "РЫБА" или "РЫБН"
            print("\nПоиск строк с упоминанием рыбы:")
            for i in range(min(100, len(df))):
                row_text = ' '.join(str(v) for v in df.iloc[i] if pd.notna(v))
                if 'РЫБ' in row_text.upper():
                    print(f"  Строка {i+1}: {row_text[:100]}...")
                    
                    # Показываем содержимое столбцов E, F, G для этой строки
                    if len(df.columns) > 6:
                        print(f"    Столбец E (индекс 4): {df.iloc[i, 4] if pd.notna(df.iloc[i, 4]) else 'пусто'}")
                        print(f"    Столбец F (индекс 5): {df.iloc[i, 5] if pd.notna(df.iloc[i, 5]) else 'пусто'}")
                        print(f"    Столбец G (индекс 6): {df.iloc[i, 6] if pd.notna(df.iloc[i, 6]) else 'пусто'}")
                    
        except Exception as e:
            print(f"Ошибка диагностики: {e}")

if __name__ == "__main__":
    test_fish_extraction()
