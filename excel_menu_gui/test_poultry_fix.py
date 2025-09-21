#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from presentation_handler import extract_poultry_dishes_from_excel
import sys
import os

def test_poultry_extraction(file_path):
    """Тестируем исправленную функцию извлечения блюд из птицы"""
    
    print(f"=== ТЕСТ ИЗВЛЕЧЕНИЯ БЛЮД ИЗ ПТИЦЫ ===")
    print(f"Файл: {os.path.basename(file_path)}")
    print()
    
    try:
        # Извлекаем блюда из птицы
        poultry_dishes = extract_poultry_dishes_from_excel(file_path)
        
        print(f"🐔 РЕЗУЛЬТАТЫ:")
        print(f"   Найдено блюд из птицы: {len(poultry_dishes)}")
        
        if poultry_dishes:
            print("\n📋 СПИСОК БЛЮД:")
            for i, dish in enumerate(poultry_dishes, 1):
                print(f"   {i:2d}. {dish.name}")
                print(f"       Вес: {dish.weight if dish.weight else 'не указан'}")
                print(f"       Цена: {dish.price if dish.price else 'не указана'}")
                print()
        else:
            print("❌ Блюда из птицы не найдены!")
    
    except Exception as e:
        print(f"❌ Ошибка при тестировании: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    file_path = r"C:\Users\katya\Downloads\Telegram Desktop\18 сентября - четверг.xls"
    
    if not os.path.exists(file_path):
        print(f"❌ Файл не найден: {file_path}")
        sys.exit(1)
    
    test_poultry_extraction(file_path)
