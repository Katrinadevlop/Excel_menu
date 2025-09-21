"""
Финальный тест интеграции бракеражного журнала
"""
from pathlib import Path
from brokerage_journal import create_brokerage_journal_from_menu

def test_final_integration():
    """Финальный тест создания бракеражного журнала"""
    
    # Пути к файлам
    menu_file = r"C:\Users\katya\Downloads\Telegram Desktop\5  сентября - пятница (3).xls"
    output_file = r"C:\Users\katya\Desktop\ФИНАЛЬНЫЙ_бракеражный_журнал_05.09.2025.xlsx"
    
    if not Path(menu_file).exists():
        print(f"❌ Файл меню не найден: {menu_file}")
        return
    
    print("🔄 Создание финального бракеражного журнала...")
    print(f"📁 Исходный файл: {Path(menu_file).name}")
    print(f"💾 Выходной файл: {Path(output_file).name}")
    
    # Создаем бракеражный журнал
    success, message = create_brokerage_journal_from_menu(menu_file, output_file)
    
    if success:
        print(f"✅ {message}")
        print(f"📄 Файл создан: {output_file}")
        
        # Проверяем размер файла
        if Path(output_file).exists():
            size = Path(output_file).stat().st_size
            print(f"📊 Размер файла: {size} байт")
            
            # Показываем что именно было извлечено
            from brokerage_journal import BrokerageJournalGenerator
            generator = BrokerageJournalGenerator()
            
            print("\n📋 Извлеченные данные:")
            
            # Дата
            date = generator.extract_date_from_menu(menu_file)
            print(f"📅 Дата: {date.strftime('%d.%m.%Y') if date else 'Не найдена'}")
            
            # Блюда по категориям
            dishes = generator.extract_dishes_from_menu(menu_file)
            total_dishes = 0
            for category, dish_list in dishes.items():
                if dish_list:
                    count = min(len(dish_list), 20)  # Ограничиваем как в коде
                    total_dishes += count
                    print(f"🍽️ {category.upper()}: {count} блюд")
            
            print(f"🔢 Всего блюд в журнале: {total_dishes}")
            
        else:
            print("❌ Файл не был создан!")
    else:
        print(f"❌ Ошибка: {message}")

if __name__ == "__main__":
    test_final_integration()
