#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Тестовый скрипт для проверки функции извлечения данных из столбца E 
от "блюда из рыбы" до "гарниров" и вставки на 6-й слайд презентации.
"""

import sys
from pathlib import Path

# Добавляем текущую папку в путь для импорта
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

from presentation_handler import extract_fish_to_side_dishes_from_column_e, create_presentation_with_fish_and_side_dishes


def test_extract_function():
    """Тест функции извлечения данных из Excel"""
    print("🧪 Тестируем функцию извлечения данных из столбца E...")
    
    # Здесь должен быть путь к тестовому Excel файлу
    # Замените на реальный путь к вашему файлу
    test_excel_path = input("Введите путь к Excel файлу для тестирования (или нажмите Enter для пропуска): ")
    
    if not test_excel_path.strip():
        print("⏭️  Тест пропущен - файл не указан")
        return
    
    if not Path(test_excel_path).exists():
        print(f"❌ Файл не найден: {test_excel_path}")
        return
    
    try:
        dishes = extract_fish_to_side_dishes_from_column_e(test_excel_path)
        print(f"✅ Извлечено {len(dishes)} блюд из столбца E:")
        
        for i, dish in enumerate(dishes[:10]):  # Показываем первые 10 блюд
            print(f"   {i+1}. {dish.name} | {dish.weight} | {dish.price}")
        
        if len(dishes) > 10:
            print(f"   ... и ещё {len(dishes) - 10} блюд")
            
    except Exception as e:
        print(f"❌ Ошибка при извлечении: {e}")


def test_presentation_function():
    """Тест функции создания презентации"""
    print("\n🧪 Тестируем функцию создания презентации...")
    
    excel_path = input("Введите путь к Excel файлу: ")
    if not excel_path.strip():
        print("⏭️  Тест пропущен - файл не указан")
        return
    
    template_path = input("Введите путь к шаблону презентации (*.pptx): ")
    if not template_path.strip():
        print("⏭️  Тест пропущен - шаблон не указан")
        return
    
    if not Path(excel_path).exists():
        print(f"❌ Excel файл не найден: {excel_path}")
        return
    
    if not Path(template_path).exists():
        print(f"❌ Шаблон не найден: {template_path}")
        return
    
    output_path = Path.home() / "Desktop" / "test_fish_garnish_presentation.pptx"
    
    try:
        success, message = create_presentation_with_fish_and_side_dishes(
            template_path, excel_path, str(output_path)
        )
        
        if success:
            print(f"✅ {message}")
            print(f"📁 Файл сохранён: {output_path}")
        else:
            print(f"❌ {message}")
            
    except Exception as e:
        print(f"❌ Ошибка при создании презентации: {e}")


def main():
    print("🔧 Тестирование новой функциональности 'Рыба + гарниры на 6 слайд'")
    print("=" * 60)
    
    # Тест 1: Извлечение данных
    test_extract_function()
    
    # Тест 2: Создание презентации
    test_presentation_function()
    
    print("\n✨ Тестирование завершено!")


if __name__ == "__main__":
    main()
