#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тест исправленной функции выбора таблиц на всех слайдах
"""
import os
import sys
from pathlib import Path

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from presentation_handler import (
    extract_salads_from_excel,
    extract_first_courses_from_excel,
    extract_meat_dishes_from_excel,
    extract_poultry_dishes_from_excel,
    extract_fish_dishes_from_column_e,
    extract_side_dishes_from_excel,
    update_presentation_with_all_categories
)

def test_complete_presentation():
    """Тестирует создание полной презентации"""
    print("🧪 ТЕСТ ИСПРАВЛЕННОЙ СИСТЕМЫ ВЫБОРА ТАБЛИЦ")
    print("=" * 70)
    
    # Используем тестовый Excel файл
    test_excel = Path(r"C:\Users\katya\Downloads\Telegram Desktop\11 сентября - четверг.xlsx")
    template_path = Path("templates/presentation_template.pptx")
    output_path = Path("test_fixed_presentation.pptx")
    
    if not test_excel.exists():
        print(f"❌ Excel файл не найден: {test_excel}")
        return
        
    if not template_path.exists():
        print(f"❌ Шаблон не найден: {template_path}")
        return
    
    print(f"📄 Используем Excel: {test_excel.name}")
    print(f"📄 Шаблон: {template_path}")
    print(f"💾 Результат: {output_path}")
    print()
    
    # Извлекаем данные по категориям
    print("🔍 ИЗВЛЕЧЕНИЕ ДАННЫХ:")
    
    print("1. Салаты...")
    salads = extract_salads_from_excel(str(test_excel))
    print(f"   ✅ Найдено: {len(salads)} блюд")
    
    print("2. Первые блюда...")
    first_courses = extract_first_courses_from_excel(str(test_excel))
    print(f"   ✅ Найдено: {len(first_courses)} блюд")
    
    print("3. Мясные блюда...")
    meat_dishes = extract_meat_dishes_from_excel(str(test_excel))
    print(f"   ✅ Найдено: {len(meat_dishes)} блюд")
    
    print("4. Блюда из птицы...")
    poultry_dishes = extract_poultry_dishes_from_excel(str(test_excel))
    print(f"   ✅ Найдено: {len(poultry_dishes)} блюд")
    
    print("5. Рыбные блюда...")
    fish_dishes = extract_fish_dishes_from_column_e(str(test_excel))
    print(f"   ✅ Найдено: {len(fish_dishes)} блюд")
    
    print("6. Гарниры...")
    side_dishes = extract_side_dishes_from_excel(str(test_excel))
    print(f"   ✅ Найдено: {len(side_dishes)} блюд")
    
    # Подготавливаем данные
    all_dishes = {
        'salads': salads,
        'first_courses': first_courses,
        'meat': meat_dishes,
        'poultry': poultry_dishes,
        'fish': fish_dishes,
        'side_dishes': side_dishes,
    }
    
    total_dishes = sum(len(dishes) for dishes in all_dishes.values())
    print(f"\n📊 ИТОГО ИЗВЛЕЧЕНО: {total_dishes} блюд")
    
    if total_dishes == 0:
        print("❌ Нет данных для создания презентации")
        return
    
    # Создаем презентацию
    print(f"\n🎯 СОЗДАНИЕ ПРЕЗЕНТАЦИИ:")
    print("Применяем исправленную логику выбора таблиц...")
    
    try:
        success = update_presentation_with_all_categories(
            str(template_path),
            all_dishes,
            str(output_path)
        )
        
        if success:
            print(f"✅ Презентация создана успешно: {output_path}")
            if output_path.exists():
                size = output_path.stat().st_size
                print(f"📏 Размер файла: {size:,} байт")
        else:
            print("❌ Ошибка создания презентации")
            
    except Exception as e:
        print(f"❌ Исключение: {e}")
        import traceback
        traceback.print_exc()
    
    print("\n🎉 ТЕСТ ЗАВЕРШЕН!")
    print(f"Проверьте файл {output_path} чтобы убедиться что:")
    print("- Названия блюд отображаются на всех слайдах")
    print("- Выбираются правильные таблицы")
    print("- Данные корректно форматированы")

if __name__ == "__main__":
    test_complete_presentation()
