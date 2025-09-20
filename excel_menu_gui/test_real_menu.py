#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Тестирование извлечения блюд из реальных файлов меню.
"""

import os
import sys
from pathlib import Path

# Добавляем текущую папку в путь для импорта
sys.path.insert(0, str(Path(__file__).parent))

from presentation_handler import (
    extract_fish_dishes_from_excel,
    extract_salads_from_excel,
    extract_first_courses_from_excel,
    extract_meat_dishes_from_excel,
    extract_poultry_dishes_from_excel,
    extract_side_dishes_from_excel
)

def test_real_menu():
    """Тестирует реальные файлы меню."""
    
    # Реальные файлы меню
    real_menu_files = [
        r"C:\Users\katya\Desktop\menurepit\5  сентября - пятница.xlsx",
        r"C:\Users\katya\Desktop\menurepit\01  августа - пятница.xls",
        r"C:\Users\katya\Desktop\menurepit\8 сентября - понедельник (2).xls",
        r"C:\Users\katya\Desktop\menurepit\5  сентября - пятница.xls"
    ]
    
    # Найдем первый существующий файл
    excel_path = None
    for file_path in real_menu_files:
        if os.path.exists(file_path):
            excel_path = file_path
            break
    
    if not excel_path:
        print("❌ Реальные файлы меню не найдены!")
        print("Проверьте следующие пути:")
        for file_path in real_menu_files:
            print(f"  - {file_path}")
        return
    
    print(f"📂 Тестируем реальный файл меню: {Path(excel_path).name}")
    print("=" * 80)
    
    # Тестируем только рыбные блюда с диагностикой
    print("\n🐟 ИЗВЛЕЧЕНИЕ РЫБНЫХ БЛЮД ИЗ РЕАЛЬНОГО МЕНЮ")
    print("=" * 60)
    
    try:
        fish_dishes = extract_fish_dishes_from_excel(excel_path)
        
        print(f"\n📊 РЕЗУЛЬТАТ: Найдено {len(fish_dishes)} рыбных блюд")
        print("=" * 60)
        
        if fish_dishes:
            print("\n📋 РЫБНЫЕ БЛЮДА ДЛЯ ПРЕЗЕНТАЦИИ:")
            print("-" * 60)
            print(f"{'№':<3} {'Название':<35} {'Вес':<12} {'Цена':<10}")
            print("-" * 60)
            
            for i, dish in enumerate(fish_dishes, 1):
                # Очищаем цену для презентации
                import re
                clean_price = re.sub(r'\s*(руб\.?|рублей|р\.?|₽|RUB)', '', dish.price, flags=re.IGNORECASE).strip()
                
                name = dish.name[:32] + "..." if len(dish.name) > 35 else dish.name
                weight = dish.weight[:9] + "..." if len(dish.weight) > 12 else dish.weight
                price = clean_price[:7] + "..." if len(clean_price) > 10 else clean_price
                
                print(f"{i:<3} {name:<35} {weight:<12} {price:<10}")
            
            print("-" * 60)
            print(f"✅ Эти {len(fish_dishes)} блюд будут вставлены в слайд 6 презентации")
            
        else:
            print("❌ Рыбные блюда в реальном файле НЕ НАЙДЕНЫ!")
            print("\n🔍 ВОЗМОЖНЫЕ ПРИЧИНЫ:")
            print("- Нет раздела 'БЛЮДА ИЗ РЫБЫ' в файле")
            print("- Данные находятся в других столбцах")
            print("- Структура файла отличается от ожидаемой")
        
    except Exception as e:
        print(f"❌ ОШИБКА при извлечении: {e}")
        import traceback
        print(f"Детали ошибки: {traceback.format_exc()}")
    
    # Быстрая проверка всех категорий
    print(f"\n\n🔍 БЫСТРАЯ ПРОВЕРКА ВСЕХ КАТЕГОРИЙ")
    print("=" * 60)
    
    categories = [
        ("Салаты", extract_salads_from_excel),
        ("Первые блюда", extract_first_courses_from_excel), 
        ("Мясные блюда", extract_meat_dishes_from_excel),
        ("Блюда из птицы", extract_poultry_dishes_from_excel),
        ("Рыбные блюда", extract_fish_dishes_from_excel),
        ("Гарниры", extract_side_dishes_from_excel)
    ]
    
    total_found = 0
    for category_name, extract_func in categories:
        try:
            dishes = extract_func(excel_path)
            count = len(dishes)
            total_found += count
            status = "✅" if count > 0 else "❌"
            print(f"{status} {category_name:<20} - {count:>3} блюд")
        except Exception as e:
            print(f"❌ {category_name:<20} - ОШИБКА: {str(e)[:40]}...")
    
    print("-" * 40)
    print(f"📈 ВСЕГО найдено: {total_found} блюд")
    
    if total_found == 0:
        print(f"\n⚠️  В файле {Path(excel_path).name} НЕ НАЙДЕНО блюд!")
        print("Это может означать:")
        print("- Структура этого файла отличается от поддерживаемой")
        print("- Данные находятся в других местах")
        print("- Нужно настроить функции под этот конкретный формат")

if __name__ == "__main__":
    test_real_menu()
