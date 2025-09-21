#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from brokerage_journal import BrokerageJournalGenerator
from pathlib import Path

def test_final_brokerage():
    """Финальный тест создания бракеражного журнала"""
    
    generator = BrokerageJournalGenerator()
    
    # Ищем реальный файл меню
    menu_files = [
        r"C:\Users\katya\Downloads\Telegram Desktop\5  сентября - пятница (3).xls",
        "templates/Шаблон меню пример.xlsx"
    ]
    
    menu_path = None
    for file in menu_files:
        if Path(file).exists():
            menu_path = file
            break
    
    if not menu_path:
        print("❌ Не найден файл меню для тестирования")
        return
    
    template_path = "templates/Бракеражный журнал шаблон.xlsx"
    output_path = "ФИНАЛЬНЫЙ_бракеражный_журнал.xlsx"
    
    print("🔄 Финальный тест создания бракеражного журнала...")
    print(f"📁 Файл меню: {Path(menu_path).name}")
    print(f"📋 Шаблон: {Path(template_path).name}")
    print(f"💾 Выходной файл: {output_path}")
    
    # Проверяем извлечение данных
    menu_date = generator.extract_date_from_menu(menu_path)
    dishes = generator.extract_dishes_from_menu(menu_path)
    
    print(f"📅 Дата из меню: {menu_date.strftime('%d.%m.%Y') if menu_date else 'Не найдена'}")
    print(f"🍽️ Извлечено блюд: {len(dishes)}")
    
    if dishes:
        print("📋 Первые 10 блюд:")
        for i, dish in enumerate(dishes[:10]):
            print(f"  {i+1:2d}. {dish}")
        if len(dishes) > 10:
            print(f"  ... и еще {len(dishes) - 10} блюд")
    
    # Создаем журнал
    success, message = generator.create_brokerage_journal(menu_path, template_path, output_path)
    
    if success:
        print(f"\n✅ {message}")
        
        # Проверяем результат
        if Path(output_path).exists():
            size = Path(output_path).stat().st_size
            print(f"📊 Размер файла: {size} байт")
            print(f"📄 Файл сохранен: {output_path}")
        else:
            print("❌ Файл не был создан!")
    else:
        print(f"\n❌ Ошибка: {message}")

if __name__ == "__main__":
    test_final_brokerage()
