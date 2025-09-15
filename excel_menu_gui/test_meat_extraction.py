#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тестовый скрипт для проверки извлечения блюд из мяса из Excel файла.
"""

import os
import sys
from pathlib import Path

# Добавляем текущую папку в путь для импорта
sys.path.append(str(Path(__file__).parent))

from presentation_handler import extract_meat_dishes_from_excel, extract_meat_dishes_by_range

def test_meat_extraction():
    """Тестирует извлечение мясных блюд из Excel файла."""
    
    print("🧪 ТЕСТИРОВАНИЕ ИЗВЛЕЧЕНИЯ МЯСНЫХ БЛЮД")
    print("=" * 50)
    
    # Ищем Excel файлы в текущей папке и родительских папках
    current_dir = Path(__file__).parent
    excel_files = []
    
    # Проверяем текущую папку
    for ext in ['*.xlsx', '*.xls']:
        excel_files.extend(list(current_dir.glob(ext)))
    
    # Проверяем родительские папки
    for parent in [current_dir.parent, current_dir.parent.parent]:
        for ext in ['*.xlsx', '*.xls']:
            excel_files.extend(list(parent.glob(ext)))
    
    if not excel_files:
        print("❌ Не найдено Excel файлов для тестирования")
        print("📋 Пожалуйста, убедитесь, что файл Excel находится в одной из следующих папок:")
        print(f"   - {current_dir}")
        print(f"   - {current_dir.parent}")
        print(f"   - {current_dir.parent.parent}")
        return
    
    # Выбираем первый найденный файл
    excel_path = str(excel_files[0])
    print(f"📁 Найден Excel файл: {excel_path}")
    print()
    
    # Тест 1: Извлечение через основную функцию
    print("🔍 ТЕСТ 1: Основная функция extract_meat_dishes_from_excel")
    print("-" * 50)
    
    try:
        meat_dishes = extract_meat_dishes_from_excel(excel_path)
        print(f"✅ Найдено мясных блюд: {len(meat_dishes)}")
        
        if meat_dishes:
            print("\n📋 Первые 5 блюд:")
            for i, dish in enumerate(meat_dishes[:5]):
                print(f"   {i+1}. {dish.name} | {dish.weight} | {dish.price}")
        else:
            print("❌ Мясные блюда не найдены основной функцией")
            
    except Exception as e:
        print(f"❌ Ошибка в основной функции: {e}")
    
    print("\n" + "=" * 50)
    
    # Тест 2: Извлечение через функцию по диапазону
    print("🔍 ТЕСТ 2: Функция по диапазону extract_meat_dishes_by_range")
    print("-" * 50)
    
    try:
        meat_dishes_range = extract_meat_dishes_by_range(excel_path)
        print(f"✅ Найдено мясных блюд: {len(meat_dishes_range)}")
        
        if meat_dishes_range:
            print("\n📋 Первые 5 блюд:")
            for i, dish in enumerate(meat_dishes_range[:5]):
                print(f"   {i+1}. {dish.name} | {dish.weight} | {dish.price}")
        else:
            print("❌ Мясные блюда не найдены функцией по диапазону")
            
    except Exception as e:
        print(f"❌ Ошибка в функции по диапазону: {e}")
    
    print("\n" + "=" * 50)
    print("🏁 ТЕСТИРОВАНИЕ ЗАВЕРШЕНО")

if __name__ == "__main__":
    test_meat_extraction()
