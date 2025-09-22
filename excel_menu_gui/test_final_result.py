#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Финальный тест системы создания бракеражного журнала
"""

from brokerage_journal import create_brokerage_journal_from_menu
import os

def main():
    print("=== ФИНАЛЬНЫЙ ТЕСТ СИСТЕМЫ ===")
    print()
    
    # Файлы для тестирования
    menu_file = 'templates/Шаблон меню пример.xlsx'
    template_file = 'templates/Бракеражный журнал шаблон.xlsx' 
    output_file = 'ИТОГОВЫЙ_БРАКЕРАЖНЫЙ_ЖУРНАЛ.xlsx'
    
    # Проверяем существование файлов
    print("Проверяем файлы:")
    print(f"✓ Меню: {os.path.exists(menu_file)} - {menu_file}")
    print(f"✓ Шаблон: {os.path.exists(template_file)} - {template_file}")
    print()
    
    # Создаем бракеражный журнал
    print("🚀 Создаем бракеражный журнал...")
    success, message = create_brokerage_journal_from_menu(menu_file, template_file, output_file)
    
    print()
    print("=== РЕЗУЛЬТАТ ===")
    if success:
        print("✅ УСПЕШНО!")
        print(f"📄 Сообщение: {message}")
        if os.path.exists(output_file):
            print(f"📁 Создан файл: {output_file}")
            file_size = os.path.getsize(output_file)
            print(f"📊 Размер файла: {file_size} bytes")
        else:
            print("❌ Файл не был создан")
    else:
        print("❌ ОШИБКА!")
        print(f"📄 Сообщение: {message}")
    
    print()
    print("=== ЗАВЕРШЕНИЕ ===")

if __name__ == "__main__":
    main()
