#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тестовый скрипт для отладки извлечения завтраков
"""

import sys
import os

# Добавляем текущий каталог в путь для импорта
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from brokerage_journal import BrokerageJournalGenerator

def main():
    # Путь к файлу меню (замените на актуальный путь)
    menu_path = r"C:\Users\katya\Desktop\menurepit\excel_menu_gui\test_all_breakfast_before_salads.xlsx"
    
    # Проверяем существование файла
    if not os.path.exists(menu_path):
        print(f"Файл меню не найден: {menu_path}")
        return
    
    print(f"Тестируем извлечение завтраков из файла: {menu_path}")
    
    # Создаем генератор
    generator = BrokerageJournalGenerator()
    
    # Тестируем извлечение категорий
    try:
        categories = generator.extract_categorized_dishes(menu_path)
        
        print(f"\n=== ИТОГОВЫЙ РЕЗУЛЬТАТ ===")
        for category, dishes in categories.items():
            if dishes:
                print(f"\n{category.upper()}: {len(dishes)} блюд")
                for i, dish in enumerate(dishes, 1):
                    print(f"  {i}. {dish}")
            else:
                print(f"\n{category.upper()}: 0 блюд")
                
        # Специально проверяем завтраки
        breakfasts = categories.get('завтрак', [])
        print(f"\n=== СТАТИСТИКА ЗАВТРАКОВ ===")
        print(f"Найдено завтраков: {len(breakfasts)}")
        if breakfasts:
            print("Список завтраков:")
            for i, breakfast in enumerate(breakfasts, 1):
                print(f"  {i}. '{breakfast}'")
        else:
            print("ЗАВТРАКИ НЕ НАЙДЕНЫ!")
            
    except Exception as e:
        print(f"Ошибка при извлечении: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
