#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Финальный тест создания бракеражного журнала
"""

import sys
import os

# Добавляем текущий каталог в путь для импорта
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from brokerage_journal import BrokerageJournalGenerator

def main():
    print("=== ТЕСТ СОЗДАНИЯ БРАКЕРАЖНОГО ЖУРНАЛА ===")
    
    # Пути к файлам
    menu_path = r"C:\Users\katya\Desktop\menurepit\excel_menu_gui\test_correct_menu.xlsx"
    template_path = r"C:\Users\katya\Desktop\menurepit\excel_menu_gui\templates\Бракеражный журнал шаблон.xlsx"
    output_path = r"C:\Users\katya\Desktop\menurepit\excel_menu_gui\test_output_journal.xlsx"
    
    # Проверяем существование файлов
    if not os.path.exists(menu_path):
        print(f"ОШИБКА: Файл меню не найден: {menu_path}")
        # Попробуем другие файлы
        alt_files = [
            "test_breakfast_column_fixed.xlsx",
            "test_final_corrected.xlsx", 
            "test_only_breakfast.xlsx"
        ]
        for alt_file in alt_files:
            alt_path = os.path.join(os.path.dirname(menu_path), alt_file)
            if os.path.exists(alt_path):
                menu_path = alt_path
                print(f"Используем альтернативный файл: {menu_path}")
                break
        else:
            print("Не найден ни один подходящий файл меню!")
            return
            
    if not os.path.exists(template_path):
        print(f"ОШИБКА: Шаблон не найден: {template_path}")
        return
    
    # Создаем генератор
    generator = BrokerageJournalGenerator()
    
    # Создаем бракеражный журнал
    success, message = generator.create_brokerage_journal(menu_path, template_path, output_path)
    
    if success:
        print(f"\n✅ УСПЕШНО: {message}")
        print(f"📄 Файл сохранен: {output_path}")
    else:
        print(f"\n❌ ОШИБКА: {message}")

if __name__ == "__main__":
    main()
