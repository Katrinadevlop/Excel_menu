from brokerage_journal import BrokerageJournalGenerator, create_brokerage_journal_from_menu
from pathlib import Path

def test_brokerage_journal():
    """Тестируем создание бракеражного журнала"""
    
    # Пути к файлам
    menu_file = r"C:\Users\katya\Downloads\Telegram Desktop\5  сентября - пятница (3).xls"
    output_file = r"C:\Users\katya\Desktop\test_brokerage_journal.xlsx"
    
    if not Path(menu_file).exists():
        print(f"Файл меню не найден: {menu_file}")
        return
    
    print("Создаем бракеражный журнал...")
    
    # Создаем генератор
    generator = BrokerageJournalGenerator()
    
    # Тестируем извлечение даты
    print("\n1. Извлечение даты из меню:")
    date = generator.extract_date_from_menu(menu_file)
    print(f"Найденная дата: {date}")
    
    # Тестируем извлечение блюд
    print("\n2. Извлечение блюд из меню:")
    dishes = generator.extract_dishes_from_menu(menu_file)
    for category, dish_list in dishes.items():
        if dish_list:
            print(f"\n{category.upper()}:")
            for dish in dish_list[:5]:  # Показываем первые 5 блюд
                print(f"  - {dish}")
            if len(dish_list) > 5:
                print(f"  ... и еще {len(dish_list) - 5} блюд")
    
    # Создаем бракеражный журнал
    print("\n3. Создание бракеражного журнала:")
    success, message = generator.create_brokerage_journal(menu_file, output_file)
    
    if success:
        print(f"✓ {message}")
        print(f"Файл сохранен: {output_file}")
        
        # Проверяем, что файл действительно создан
        if Path(output_file).exists():
            size = Path(output_file).stat().st_size
            print(f"Размер файла: {size} байт")
        else:
            print("❌ Файл не был создан!")
    else:
        print(f"❌ Ошибка: {message}")

if __name__ == "__main__":
    test_brokerage_journal()
