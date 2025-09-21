#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Финальная диагностика функций извлечения данных и создания презентации
"""
import os
import sys
from pathlib import Path
import pandas as pd

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from presentation_handler import (
    extract_fish_dishes_from_column_e,
    extract_side_dishes_from_excel,
    extract_salads_from_excel,
    extract_first_courses_from_excel,
    extract_meat_dishes_from_excel,
    extract_poultry_dishes_from_excel,
    create_presentation_with_excel_data,
    MenuItem
)

def test_excel_structure():
    """Анализирует структуру Excel файла"""
    print("🔍 АНАЛИЗ СТРУКТУРЫ EXCEL ФАЙЛА")
    print("=" * 60)
    
    # Попробуем найти файл в разных местах
    possible_files = [
        Path(r"C:\Users\katya\Downloads\Telegram Desktop\11 сентября - четверг.xlsx"),
        Path(r"C:\Users\katya\Desktop\11 сентября - четверг.xlsx"),
        Path(r"C:\Users\katya\Downloads\11 сентября - четверг.xlsx")
    ]
    
    test_file = None
    for file_path in possible_files:
        if file_path.exists():
            test_file = file_path
            break
    
    if not test_file:
        print("❌ Тестовый файл не найден в стандартных местах")
        print("📍 Попробуйте разместить файл '11 сентября - четверг.xlsx' в одном из:")
        for p in possible_files:
            print(f"   - {p}")
        return None
    
    print(f"📄 Анализируем файл: {test_file}")
    
    try:
        xls = pd.ExcelFile(str(test_file))
        print(f"📊 Листы в файле: {xls.sheet_names}")
        
        # Выбираем лист с "касс"
        sheet_name = None
        for nm in xls.sheet_names:
            if 'касс' in str(nm).strip().lower():
                sheet_name = nm
                break
        if sheet_name is None and xls.sheet_names:
            sheet_name = xls.sheet_names[0]
            
        df = pd.read_excel(str(test_file), sheet_name=sheet_name, header=None, dtype=object)
        print(f"📊 Используемый лист: '{sheet_name}'")
        print(f"📊 Размер: {len(df)} строк, {len(df.columns)} столбцов")
        
        def row_text(row) -> str:
            parts = []
            for v in row:
                if pd.notna(v):
                    parts.append(str(v))
            return ' '.join(parts).strip()
        
        # Ищем ключевые категории
        categories_found = {}
        for i in range(min(100, len(df))):
            row_content = row_text(df.iloc[i]).upper().replace('Ё', 'Е')
            
            if 'САЛАТЫ' in row_content and 'ХОЛОДН' in row_content:
                categories_found['salads'] = i + 1
            elif 'ПЕРВЫЕ БЛЮДА' in row_content:
                categories_found['first'] = i + 1
            elif 'БЛЮДА ИЗ МЯСА' in row_content or 'МЯСНЫЕ БЛЮДА' in row_content:
                categories_found['meat'] = i + 1
            elif 'БЛЮДА ИЗ ПТИЦЫ' in row_content:
                categories_found['poultry'] = i + 1
            elif 'БЛЮДА ИЗ РЫБЫ' in row_content or 'РЫБНЫЕ БЛЮДА' in row_content:
                categories_found['fish'] = i + 1
            elif 'ГАРНИРЫ' in row_content:
                categories_found['garnish'] = i + 1
        
        print("\n📍 НАЙДЕННЫЕ КАТЕГОРИИ:")
        for category, line_num in categories_found.items():
            print(f"   {category}: строка {line_num}")
        
        return str(test_file)
        
    except Exception as e:
        print(f"❌ Ошибка анализа файла: {e}")
        return None

def test_all_extraction_functions(excel_path):
    """Тестирует все функции извлечения данных"""
    print("\n🧪 ТЕСТИРОВАНИЕ ФУНКЦИЙ ИЗВЛЕЧЕНИЯ")
    print("=" * 60)
    
    categories = [
        ("Салаты", extract_salads_from_excel),
        ("Первые блюда", extract_first_courses_from_excel),
        ("Мясные блюда", extract_meat_dishes_from_excel),
        ("Блюда из птицы", extract_poultry_dishes_from_excel),
        ("Рыбные блюда", extract_fish_dishes_from_column_e),
        ("Гарниры", extract_side_dishes_from_excel),
    ]
    
    results = {}
    
    for category_name, extract_func in categories:
        print(f"\n🔍 Тестируем {category_name}...")
        try:
            dishes = extract_func(excel_path)
            results[category_name] = dishes
            print(f"   ✅ Найдено: {len(dishes)} блюд")
            
            # Показываем первые 3 блюда
            for i, dish in enumerate(dishes[:3], 1):
                status = "✅" if dish.name and len(dish.name) > 2 else "❌"
                print(f"   {i}. {status} {dish.name or '[БЕЗ НАЗВАНИЯ]'} | {dish.weight or '[БЕЗ ВЕСА]'} | {dish.price or '[БЕЗ ЦЕНЫ]'}")
            
            if len(dishes) > 3:
                print(f"   ... и ещё {len(dishes) - 3} блюд")
                
            # Проверяем качество данных
            no_name = sum(1 for d in dishes if not d.name or len(d.name) < 3)
            no_weight = sum(1 for d in dishes if not d.weight)
            no_price = sum(1 for d in dishes if not d.price)
            
            if no_name > 0:
                print(f"   ⚠️  Блюд без названия: {no_name}")
            if no_weight > 0:
                print(f"   ⚠️  Блюд без веса: {no_weight}")
            if no_price > 0:
                print(f"   ⚠️  Блюд без цены: {no_price}")
                
        except Exception as e:
            print(f"   ❌ Ошибка: {e}")
            results[category_name] = []
    
    return results

def test_presentation_creation(excel_path):
    """Тестирует создание презентации"""
    print("\n🎯 ТЕСТИРОВАНИЕ СОЗДАНИЯ ПРЕЗЕНТАЦИИ")
    print("=" * 60)
    
    # Проверяем наличие шаблона
    template_candidates = [
        Path("templates/presentation_template.pptx"),
        Path("excel_menu_gui/templates/presentation_template.pptx"),
        Path("C:/Users/katya/Desktop/menurepit/excel_menu_gui/templates/presentation_template.pptx")
    ]
    
    template_path = None
    for template in template_candidates:
        if template.exists():
            template_path = str(template)
            break
    
    if not template_path:
        print("❌ Шаблон презентации не найден")
        print("📍 Ожидаемые места:")
        for template in template_candidates:
            print(f"   - {template}")
        return False
    
    print(f"📄 Используем шаблон: {template_path}")
    
    # Создаем тестовую презентацию
    output_path = Path("test_presentation_output.pptx")
    
    try:
        success, message = create_presentation_with_excel_data(
            template_path, excel_path, str(output_path)
        )
        
        if success:
            print("✅ Презентация создана успешно!")
            print(f"📊 Результат: {message}")
            print(f"📁 Файл сохранен: {output_path}")
            return True
        else:
            print(f"❌ Ошибка создания презентации: {message}")
            return False
            
    except Exception as e:
        print(f"❌ Исключение при создании презентации: {e}")
        return False

def main():
    """Основная функция диагностики"""
    print("🚀 ФИНАЛЬНАЯ ДИАГНОСТИКА СИСТЕМЫ МЕНЮ")
    print("=" * 80)
    
    # Шаг 1: Анализ структуры Excel
    excel_path = test_excel_structure()
    if not excel_path:
        return
    
    # Шаг 2: Тестирование извлечения данных
    extraction_results = test_all_extraction_functions(excel_path)
    
    # Шаг 3: Тестирование создания презентации
    presentation_success = test_presentation_creation(excel_path)
    
    # Итоговый отчет
    print("\n📋 ИТОГОВЫЙ ОТЧЕТ")
    print("=" * 60)
    
    total_dishes = sum(len(dishes) for dishes in extraction_results.values())
    print(f"📊 Всего извлечено блюд: {total_dishes}")
    
    for category, dishes in extraction_results.items():
        quality_score = 0
        if dishes:
            good_dishes = sum(1 for d in dishes if d.name and len(d.name) > 2 and d.weight and d.price)
            quality_score = (good_dishes / len(dishes)) * 100
        
        status = "✅" if len(dishes) > 0 and quality_score > 70 else "⚠️" if len(dishes) > 0 else "❌"
        print(f"{status} {category}: {len(dishes)} блюд (качество: {quality_score:.0f}%)")
    
    print(f"\n🎯 Создание презентации: {'✅ Успешно' if presentation_success else '❌ Ошибка'}")
    
    if total_dishes > 0 and presentation_success:
        print("\n🎉 СИСТЕМА ГОТОВА К РАБОТЕ!")
    else:
        print("\n⚠️  ТРЕБУЮТСЯ ДОРАБОТКИ")

if __name__ == "__main__":
    main()
