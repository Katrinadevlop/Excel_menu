#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from presentation_handler import create_presentation_with_excel_data
import sys
import os

def test_18_september():
    """Тестируем полный процесс создания презентации с файлом 18 сентября"""
    
    print(f"=== ПОЛНЫЙ ТЕСТ С ФАЙЛОМ 18 СЕНТЯБРЯ ===")
    print()
    
    # Пути к файлам
    template_path = r"C:\Users\katya\Desktop\Template_menu.pptx"
    excel_path = r"C:\Users\katya\Downloads\Telegram Desktop\18 сентября - четверг.xls"
    output_path = r"C:\Users\katya\Desktop\меню_18_сентября_исправлено.pptx"
    
    print(f"📄 Шаблон: {os.path.basename(template_path)}")
    print(f"📊 Excel: {os.path.basename(excel_path)}")
    print(f"💾 Выходной файл: {os.path.basename(output_path)}")
    print()
    
    # Проверяем существование файлов
    if not os.path.exists(template_path):
        print(f"❌ Шаблон не найден: {template_path}")
        # Попробуем найти другие шаблоны
        for possible_template in [
            r"C:\Users\katya\Desktop\menurepit\excel_menu_gui\templates\presentation_template.pptx",
            r"C:\Users\katya\Desktop\menurepit\templates\presentation_template.pptx"
        ]:
            if os.path.exists(possible_template):
                template_path = possible_template
                print(f"✅ Найден альтернативный шаблон: {template_path}")
                break
        else:
            print("❌ Не найден ни один шаблон!")
            return False
        
    if not os.path.exists(excel_path):
        print(f"❌ Excel файл не найден: {excel_path}")
        return False
    
    try:
        print("🚀 Запускаем создание презентации...")
        print("-" * 60)
        
        success, message = create_presentation_with_excel_data(
            template_path=template_path,
            excel_path=excel_path, 
            output_path=output_path
        )
        
        print("-" * 60)
        print(f"📋 РЕЗУЛЬТАТ:")
        print(f"   Статус: {'✅ Успешно' if success else '❌ Ошибка'}")
        print(f"   Сообщение: {message}")
        
        if success:
            print(f"   Файл сохранен: {output_path}")
            if os.path.exists(output_path):
                size = os.path.getsize(output_path)
                print(f"   Размер файла: {size:,} байт")
        
        return success
        
    except Exception as e:
        print(f"❌ Исключение при создании презентации: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    test_18_september()
