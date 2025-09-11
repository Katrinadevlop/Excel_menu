#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Тестовый скрипт для проверки функций извлечения блюд из Excel файлов.
"""

import os
import sys
from pathlib import Path

# Добавляем текущую папку в путь для импорта
sys.path.insert(0, str(Path(__file__).parent))

from presentation_handler import (
    extract_salads_from_excel,
    extract_first_courses_from_excel,
    extract_meat_dishes_from_excel
)


def test_extraction(excel_path: str):
    """Тестирует извлечение блюд из Excel файла."""
    
    if not Path(excel_path).exists():
        print(f"❌ Файл не найден: {excel_path}")
        return
    
    print(f"🧪 Тестируем извлечение из файла: {Path(excel_path).name}")
    print("=" * 80)
    
    # Тестируем извлечение салатов
    print("\n🥗 САЛАТЫ И ХОЛОДНЫЕ ЗАКУСКИ:")
    print("-" * 40)
    try:
        salads = extract_salads_from_excel(excel_path)
        print(f"✅ Найдено {len(salads)} салатов")
        for i, dish in enumerate(salads[:5], 1):  # Показываем первые 5
            print(f"  {i}. {dish.name} | {dish.weight} | {dish.price}")
        if len(salads) > 5:
            print(f"  ... и ещё {len(salads) - 5} салатов")
    except Exception as e:
        print(f"❌ Ошибка при извлечении салатов: {e}")
    
    # Тестируем извлечение первых блюд
    print("\n🍲 ПЕРВЫЕ БЛЮДА:")
    print("-" * 40)
    try:
        first_courses = extract_first_courses_from_excel(excel_path)
        print(f"✅ Найдено {len(first_courses)} первых блюд")
        for i, dish in enumerate(first_courses[:5], 1):  # Показываем первые 5
            print(f"  {i}. {dish.name} | {dish.weight} | {dish.price}")
        if len(first_courses) > 5:
            print(f"  ... и ещё {len(first_courses) - 5} первых блюд")
    except Exception as e:
        print(f"❌ Ошибка при извлечении первых блюд: {e}")
    
    # Тестируем извлечение мясных блюд
    print("\n🥩 БЛЮДА ИЗ МЯСА:")
    print("-" * 40)
    try:
        meat_dishes = extract_meat_dishes_from_excel(excel_path)
        print(f"✅ Найдено {len(meat_dishes)} мясных блюд")
        for i, dish in enumerate(meat_dishes[:5], 1):  # Показываем первые 5
            print(f"  {i}. {dish.name} | {dish.weight} | {dish.price}")
        if len(meat_dishes) > 5:
            print(f"  ... и ещё {len(meat_dishes) - 5} мясных блюд")
    except Exception as e:
        print(f"❌ Ошибка при извлечении мясных блюд: {e}")
    
    # Суммарная статистика
    total_dishes = len(salads) + len(first_courses) + len(meat_dishes)
    print("\n📊 ИТОГО:")
    print("-" * 40)
    print(f"Салаты: {len(salads)} блюд")
    print(f"Первые блюда: {len(first_courses)} блюд") 
    print(f"Мясные блюда: {len(meat_dishes)} блюд")
    print(f"ВСЕГО: {total_dishes} блюд")
    
    if total_dishes == 0:
        print("\n⚠️  Не найдено ни одного блюда!")
        print("Возможные причины:")
        print("- Структура файла отличается от ожидаемой")
        print("- Названия категорий написаны по-другому")
        print("- Данные находятся в другом листе")
    
    print("\n" + "=" * 80)
    return total_dishes > 0


def main():
    """Главная функция."""
    print("🧪 ТЕСТ ИЗВЛЕЧЕНИЯ БЛЮД ИЗ EXCEL")
    print("=" * 80)
    
    # Ищем Excel файлы в папке templates
    templates_dir = Path(__file__).parent / "templates"
    excel_files = []
    
    if templates_dir.exists():
        excel_files.extend(templates_dir.glob("*.xlsx"))
        excel_files.extend(templates_dir.glob("*.xls"))
    
    if not excel_files:
        print("❌ Не найдены Excel файлы в папке templates/")
        print("Положите тестовый Excel файл с меню в папку templates/ и запустите снова.")
        return
    
    # Тестируем каждый найденный файл
    for excel_file in excel_files:
        success = test_extraction(str(excel_file))
        if success:
            print("✅ Тест прошёл успешно!")
        else:
            print("❌ Тест не прошёл - блюда не найдены")


if __name__ == "__main__":
    main()
