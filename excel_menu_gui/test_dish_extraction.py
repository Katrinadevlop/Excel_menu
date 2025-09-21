#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from brokerage_journal import BrokerageJournalGenerator

def test_dish_extraction():
    """Тестируем извлечение блюд из меню"""
    
    generator = BrokerageJournalGenerator()
    menu_path = "templates/Шаблон меню пример.xlsx"
    
    print("🔄 Тестируем извлечение блюд...")
    print(f"📁 Файл меню: {menu_path}")
    
    # Тестируем извлечение даты
    menu_date = generator.extract_date_from_menu(menu_path)
    print(f"📅 Извлеченная дата: {menu_date}")
    
    # Тестируем извлечение блюд
    dishes = generator.extract_dishes_from_menu(menu_path)
    print(f"🍽️ Найдено блюд: {len(dishes)}")
    
    if dishes:
        print("📋 Список извлеченных блюд:")
        for i, dish in enumerate(dishes[:20]):  # Показываем первые 20
            print(f"  {i+1:2d}. {dish}")
        
        if len(dishes) > 20:
            print(f"  ... и еще {len(dishes) - 20} блюд")
    else:
        print("❌ Блюда не были извлечены")
        
        # Попробуем более детальную диагностику
        print("\n🔍 Детальная диагностика...")
        import pandas as pd
        df_dict = pd.read_excel(menu_path, sheet_name=None)
        
        for sheet_name, df in df_dict.items():
            print(f"\n📋 Лист: {sheet_name}")
            print(f"📏 Размер: {len(df)} строк, {len(df.columns)} столбцов")
            
            # Показываем первые несколько строк
            for i, (_, row) in enumerate(df.head(10).iterrows()):
                row_content = []
                for cell in row:
                    if pd.notna(cell):
                        row_content.append(str(cell).strip())
                
                if row_content:
                    print(f"  Строка {i+1}: {' | '.join(row_content[:5])}")

if __name__ == "__main__":
    test_dish_extraction()
