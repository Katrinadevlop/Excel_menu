#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from presentation_handler import extract_side_dishes_from_excel
import sys
import os

def test_side_dishes_extraction(file_path):
    """Тестируем исправленную функцию извлечения гарниров"""
    
    print(f"=== ТЕСТ ИЗВЛЕЧЕНИЯ ГАРНИРОВ ===")
    print(f"Файл: {os.path.basename(file_path)}")
    print()
    
    try:
        # Извлекаем гарниры
        side_dishes = extract_side_dishes_from_excel(file_path)
        
        print(f"🥔 РЕЗУЛЬТАТЫ:")
        print(f"   Найдено гарниров: {len(side_dishes)}")
        
        if side_dishes:
            print("\n📋 СПИСОК ГАРНИРОВ:")
            for i, dish in enumerate(side_dishes, 1):
                print(f"   {i:2d}. {dish.name}")
                print(f"       Вес: {dish.weight if dish.weight else 'не указан'}")
                print(f"       Цена: {dish.price if dish.price else 'не указана'}")
                print()
        else:
            print("❌ Гарниры не найдены!")
    
    except Exception as e:
        print(f"❌ Ошибка при тестировании: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    file_path = r"C:\Users\katya\Downloads\Telegram Desktop\18 сентября - четверг.xls"
    
    if not os.path.exists(file_path):
        print(f"❌ Файл не найден: {file_path}")
        sys.exit(1)
    
    test_side_dishes_extraction(file_path)
