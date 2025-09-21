from brokerage_journal import create_brokerage_journal_from_menu
from pathlib import Path

def test_updated_brokerage():
    """Тестируем обновленный функционал с шаблоном"""
    
    # Пути к файлам
    menu_file = r"C:\Users\katya\Downloads\Telegram Desktop\5  сентября - пятница (3).xls"
    template_file = r"C:\Users\katya\Desktop\menurepit\excel_menu_gui\templates\Бракеражный журнал шаблон.xlsx"
    output_file = r"C:\Users\katya\Desktop\ТЕСТ_новый_бракеражный_журнал.xlsx"
    
    print("🔄 Тестируем обновленный функционал бракеражного журнала...")
    
    # Проверяем наличие файлов
    if not Path(menu_file).exists():
        print(f"❌ Файл меню не найден: {menu_file}")
        return
        
    if not Path(template_file).exists():
        print(f"❌ Шаблон не найден: {template_file}")
        return
    
    print(f"📁 Файл меню: {Path(menu_file).name}")
    print(f"📋 Шаблон: {Path(template_file).name}")
    print(f"💾 Выходной файл: {Path(output_file).name}")
    
    # Создаем бракеражный журнал
    success, message = create_brokerage_journal_from_menu(menu_file, template_file, output_file)
    
    if success:
        print(f"✅ {message}")
        print(f"📄 Файл создан: {output_file}")
        
        # Проверяем размер файла
        if Path(output_file).exists():
            size = Path(output_file).stat().st_size
            print(f"📊 Размер файла: {size} байт")
        else:
            print("❌ Файл не был создан!")
    else:
        print(f"❌ Ошибка: {message}")

if __name__ == "__main__":
    test_updated_brokerage()
