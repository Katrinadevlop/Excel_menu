#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тестовый скрипт для проверки функции extract_dishes_from_column_d7_d38
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__)))

from app.services.dish_extractor import extract_dishes_from_column_d7_d38

def test_extraction():
    """Тестирует извлечение блюд из диапазона D7:D38"""
    
    # Здесь нужно будет указать путь к реальному файлу меню для тестирования
    test_file_path = input("Введите путь к файлу меню для тестирования: ").strip()
    
    if not test_file_path or not os.path.exists(test_file_path):
        print("Файл не найден или не указан")
        return
    
    try:
        print(f"Тестируем извлечение из файла: {test_file_path}")
        print("=" * 60)
        
        dishes = extract_dishes_from_column_d7_d38(test_file_path)
        
        print(f"\nРезультат извлечения:")
        print(f"Всего найдено блюд: {len(dishes)}")
        print("-" * 40)
        
        for i, dish in enumerate(dishes, 1):
            print(f"{i:2d}. {dish}")
        
        print("=" * 60)
        print("Тест завершен успешно!")
        
    except Exception as e:
        print(f"Ошибка при тестировании: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_extraction()