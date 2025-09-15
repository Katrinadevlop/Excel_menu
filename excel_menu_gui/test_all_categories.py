#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
from pathlib import Path

# Добавляем текущую директорию в путь для импорта
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from presentation_handler import (
    extract_salads_from_excel,
    extract_first_courses_from_excel,
    extract_meat_dishes_from_excel,
    extract_poultry_dishes_from_excel,
    extract_fish_dishes_from_excel,
    extract_side_dishes_from_excel
)

def test_all_categories():
    """Тестирует извлечение всех категорий блюд из Excel файла."""
    
    # Используем файл из папки templates
    excel_path = r"C:\Users\katya\Desktop\menurepit\excel_menu_gui\templates\Шаблон меню пример.xlsx"
    
    if not Path(excel_path).exists():
        print("❌ Excel файл не найден")
        return
    
    print(f"📂 Используется файл: {Path(excel_path).name}")
    print("=" * 80)
    print("ТЕСТИРОВАНИЕ ИЗВЛЕЧЕНИЯ ВСЕХ КАТЕГОРИЙ БЛЮД")
    print("=" * 80)
    
    # Тест 1: Салаты
    print("\n🥗 САЛАТЫ И ХОЛОДНЫЕ ЗАКУСКИ")
    print("-" * 40)
    salads = extract_salads_from_excel(excel_path)
    if salads:
        print(f"✅ Найдено {len(salads)} салатов:")
        for i, dish in enumerate(salads[:5], 1):
            print(f"  {i}. {dish.name} | {dish.weight} | {dish.price}")
    else:
        print("❌ Салаты не найдены")
    
    # Тест 2: Первые блюда
    print("\n🍲 ПЕРВЫЕ БЛЮДА")
    print("-" * 40)
    first_courses = extract_first_courses_from_excel(excel_path)
    if first_courses:
        print(f"✅ Найдено {len(first_courses)} первых блюд:")
        for i, dish in enumerate(first_courses[:5], 1):
            print(f"  {i}. {dish.name} | {dish.weight} | {dish.price}")
    else:
        print("❌ Первые блюда не найдены")
    
    # Тест 3: Блюда из мяса
    print("\n🥩 БЛЮДА ИЗ МЯСА")
    print("-" * 40)
    meat_dishes = extract_meat_dishes_from_excel(excel_path)
    if meat_dishes:
        print(f"✅ Найдено {len(meat_dishes)} мясных блюд:")
        for i, dish in enumerate(meat_dishes[:5], 1):
            print(f"  {i}. {dish.name} | {dish.weight} | {dish.price}")
    else:
        print("❌ Мясные блюда не найдены")
    
    # Тест 4: Блюда из птицы
    print("\n🍗 БЛЮДА ИЗ ПТИЦЫ")
    print("-" * 40)
    poultry_dishes = extract_poultry_dishes_from_excel(excel_path)
    if poultry_dishes:
        print(f"✅ Найдено {len(poultry_dishes)} блюд из птицы:")
        for i, dish in enumerate(poultry_dishes[:5], 1):
            print(f"  {i}. {dish.name} | {dish.weight} | {dish.price}")
    else:
        print("❌ Блюда из птицы не найдены")
    
    # Тест 5: Блюда из рыбы
    print("\n🐟 БЛЮДА ИЗ РЫБЫ")
    print("-" * 40)
    fish_dishes = extract_fish_dishes_from_excel(excel_path)
    if fish_dishes:
        print(f"✅ Найдено {len(fish_dishes)} рыбных блюд:")
        for i, dish in enumerate(fish_dishes[:5], 1):
            print(f"  {i}. {dish.name} | {dish.weight} | {dish.price}")
    else:
        print("❌ Рыбные блюда не найдены")
    
    # Тест 6: Гарниры
    print("\n🍚 ГАРНИРЫ")
    print("-" * 40)
    side_dishes = extract_side_dishes_from_excel(excel_path)
    if side_dishes:
        print(f"✅ Найдено {len(side_dishes)} гарниров:")
        for i, dish in enumerate(side_dishes[:5], 1):
            print(f"  {i}. {dish.name} | {dish.weight} | {dish.price}")
    else:
        print("❌ Гарниры не найдены")
    
    # Итоговая статистика
    print("\n" + "=" * 80)
    print("📊 ИТОГОВАЯ СТАТИСТИКА:")
    print("-" * 40)
    total = len(salads) + len(first_courses) + len(meat_dishes) + len(poultry_dishes) + len(fish_dishes) + len(side_dishes)
    print(f"Салаты и холодные закуски: {len(salads)} блюд")
    print(f"Первые блюда: {len(first_courses)} блюд")
    print(f"Блюда из мяса: {len(meat_dishes)} блюд")
    print(f"Блюда из птицы: {len(poultry_dishes)} блюд")
    print(f"Блюда из рыбы: {len(fish_dishes)} блюд")
    print(f"Гарниры: {len(side_dishes)} блюд")
    print(f"ВСЕГО: {total} блюд")
    
    # Дополнительная диагностика структуры файла
    print("\n🔍 ДИАГНОСТИКА СТРУКТУРЫ ФАЙЛА:")
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
        
        # Ищем все заголовки категорий
        print("\nНайденные заголовки категорий:")
        categories = ["САЛАТ", "ПЕРВЫЕ", "МЯСН", "МЯСА", "ПТИЦ", "РЫБ", "ГАРНИР"]
        
        for i in range(min(100, len(df))):
            row_text = ' '.join(str(v) for v in df.iloc[i] if pd.notna(v)).upper()
            for cat in categories:
                if cat in row_text:
                    print(f"  Строка {i+1}: {row_text[:80]}...")
                    break
                    
    except Exception as e:
        print(f"Ошибка диагностики: {e}")

if __name__ == "__main__":
    test_all_categories()
