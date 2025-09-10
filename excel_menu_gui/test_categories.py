#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
from pathlib import Path
from presentation_handler import (
    extract_salads_from_excel,
    extract_first_courses_from_excel,
    extract_meat_dishes_from_excel,
    extract_poultry_dishes_from_excel,
    extract_fish_dishes_from_excel,
    extract_side_dishes_from_excel
)

def test_excel_categories(excel_path: str):
    """Тестирует извлечение всех категорий из Excel файла"""
    
    if not Path(excel_path).exists():
        print(f"❌ Файл не найден: {excel_path}")
        return
    
    print(f"🧪 Тестируем файл: {excel_path}")
    print("=" * 80)
    
    # Тестируем каждую категорию
    categories = [
        ("Салаты и холодные закуски", extract_salads_from_excel),
        ("Первые блюда", extract_first_courses_from_excel),
        ("Мясные блюда", extract_meat_dishes_from_excel),
        ("Блюда из птицы", extract_poultry_dishes_from_excel),
        ("Рыбные блюда", extract_fish_dishes_from_excel),
        ("Гарниры", extract_side_dishes_from_excel),
    ]
    
    total_found = 0
    results = {}
    
    for category_name, extract_func in categories:
        print(f"\n🔍 Категория: {category_name}")
        print("-" * 40)
        
        try:
            dishes = extract_func(excel_path)
            results[category_name] = dishes
            
            if dishes:
                print(f"✅ Найдено {len(dishes)} блюд:")
                for i, dish in enumerate(dishes[:5], 1):  # Показываем первые 5
                    print(f"  {i}. {dish.name} | {dish.weight} | {dish.price}")
                if len(dishes) > 5:
                    print(f"  ... и еще {len(dishes) - 5} блюд")
                total_found += len(dishes)
            else:
                print("❌ Блюда не найдены")
                
        except Exception as e:
            print(f"❌ Ошибка: {e}")
            results[category_name] = []
    
    print("\n" + "=" * 80)
    print("📊 ИТОГОВАЯ СТАТИСТИКА:")
    print("-" * 40)
    
    for category_name, dishes in results.items():
        status = "✅" if dishes else "❌"
        print(f"{status} {category_name}: {len(dishes)} блюд")
    
    print(f"\n🎯 Всего найдено блюд: {total_found}")
    
    if total_found == 0:
        print("\n💡 Возможные причины:")
        print("  • Неправильные названия категорий в Excel файле")
        print("  • Категории находятся не в том листе")
        print("  • Файл имеет неожиданную структуру")
        print("\n🔧 Рекомендации:")
        print("  • Запустите debug_categories.py для подробного анализа")
        print("  • Проверьте, что категории написаны заглавными буквами")
        print("  • Убедитесь, что используется правильный лист Excel")
    else:
        print(f"\n🎉 Успех! Найдено {total_found} блюд в {sum(1 for dishes in results.values() if dishes)} категориях")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        excel_path = sys.argv[1]
    else:
        excel_path = input("Введите путь к Excel файлу: ").strip().strip('"')
    
    test_excel_categories(excel_path)
