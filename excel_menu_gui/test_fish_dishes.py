#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os

# Добавляем текущую директорию в путь поиска модулей
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from presentation_handler import extract_fish_dishes_from_excel

def test_fish_dishes_extraction():
    print("🐟 Тестирование извлечения блюд из рыбы")
    print("=" * 50)
    
    # Путь к Excel файлу (ищем доступные файлы)
    possible_files = [
        "../8 сентября - понедельник (2).xls",
        "../5  сентября - пятница.xlsx", 
        "menu.xlsx",
        "меню.xlsx", 
        "меню на неделю.xlsx",
        "меню на неделю эксель касса 2.xlsx",
        "templates/Шаблон меню пример.xlsx"
    ]
    
    excel_path = None
    for file_path in possible_files:
        full_path = os.path.join(os.getcwd(), file_path)
        if os.path.exists(full_path):
            excel_path = full_path
            print(f"✓ Найден файл: {file_path}")
            break
    
    if not excel_path:
        print("✗ Excel файл не найден. Проверьте наличие файла меню.")
        return False
    
    try:
        # Извлекаем блюда из рыбы
        fish_dishes = extract_fish_dishes_from_excel(excel_path)
        
        print(f"\n📊 Результат извлечения:")
        print(f"Найдено рыбных блюд: {len(fish_dishes)}")
        
        if fish_dishes:
            print("\n🍽️ Список блюд из рыбы:")
            print("-" * 60)
            for i, dish in enumerate(fish_dishes, 1):
                print(f"{i:2d}. {dish.name}")
                print(f"    Вес/объем: {dish.weight if dish.weight else 'не указан'}")
                print(f"    Цена: {dish.price if dish.price else 'не указана'}")
                print()
        else:
            print("⚠️ Блюда из рыбы не найдены.")
            print("Возможные причины:")
            print("- В Excel файле нет секции 'БЛЮДА ИЗ РЫБЫ'")
            print("- Данные находятся не в столбце E")
            print("- Неправильная структура файла")
        
        return len(fish_dishes) > 0
        
    except Exception as e:
        print(f"❌ Ошибка при извлечении: {e}")
        return False

if __name__ == "__main__":
    print("Тест извлечения блюд из рыбы из Excel файла")
    print("Функция ищет данные в столбце E от заголовка 'БЛЮДА ИЗ РЫБЫ' до 'ГАРНИРЫ'")
    print()
    
    success = test_fish_dishes_extraction()
    
    if success:
        print("\n✅ Тест пройден успешно!")
    else:
        print("\n❌ Тест не пройден. Проверьте Excel файл и структуру данных.")
    
    input("\nНажмите Enter для выхода...")
