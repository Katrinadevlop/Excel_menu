#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тест извлечения рыбных блюд из Excel файлов в папке Downloads
"""
import os
import sys
from pathlib import Path
import pandas as pd

# Добавляем текущую папку в путь для импорта наших модулей
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from presentation_handler import extract_fish_dishes_from_column_e, MenuItem

def test_excel_files_in_downloads():
    """
    Тестирует извлечение рыбных блюд из всех Excel файлов в папке Downloads\Telegram Desktop
    """
    downloads_path = Path(r"C:\Users\katya\Downloads\Telegram Desktop")
    
    if not downloads_path.exists():
        print(f"❌ Папка {downloads_path} не найдена!")
        return
    
    # Находим все Excel файлы
    excel_files = []
    for pattern in ['*.xlsx', '*.xls']:
        excel_files.extend(downloads_path.glob(pattern))
    
    if not excel_files:
        print(f"❌ В папке {downloads_path} не найдены Excel файлы!")
        return
    
    print(f"🔍 Найдено Excel файлов: {len(excel_files)}")
    print("=" * 60)
    
    for i, excel_file in enumerate(excel_files[:5], 1):  # Тестируем первые 5 файлов
        print(f"\n📄 ФАЙЛ {i}: {excel_file.name}")
        print("-" * 50)
        
        try:
            # Тестируем извлечение рыбных блюд
            fish_dishes = extract_fish_dishes_from_column_e(str(excel_file))
            
            if fish_dishes:
                print(f"✅ Извлечено рыбных блюд: {len(fish_dishes)}")
                print("\n🐟 НАЙДЕННЫЕ РЫБНЫЕ БЛЮДА:")
                for j, dish in enumerate(fish_dishes, 1):
                    print(f"  {j:2d}. {dish.name}")
                    if dish.weight:
                        print(f"      Вес: {dish.weight}")
                    if dish.price:
                        print(f"      Цена: {dish.price}")
                    print()
            else:
                print("❌ Рыбные блюда не найдены")
                
                # Диагностика - посмотрим структуру файла
                try:
                    print("\n🔍 ДИАГНОСТИКА ФАЙЛА:")
                    xls = pd.ExcelFile(str(excel_file))
                    print(f"   Листы в файле: {xls.sheet_names}")
                    
                    # Берем первый лист
                    sheet_name = xls.sheet_names[0]
                    df = pd.read_excel(str(excel_file), sheet_name=sheet_name, header=None, dtype=object)
                    
                    print(f"   Размер листа '{sheet_name}': {len(df)} строк, {len(df.columns)} столбцов")
                    
                    # Ищем упоминания рыбы в первых 50 строках
                    fish_mentions = []
                    for i in range(min(50, len(df))):
                        row_text = ' '.join([str(v) for v in df.iloc[i] if pd.notna(v)]).upper()
                        if 'РЫБ' in row_text or 'FISH' in row_text:
                            fish_mentions.append((i+1, row_text[:100]))
                    
                    if fish_mentions:
                        print("   Найдены упоминания рыбы:")
                        for row_num, text in fish_mentions:
                            print(f"     Строка {row_num}: {text}")
                    else:
                        print("   Упоминания рыбы в первых 50 строках не найдены")
                        
                except Exception as diag_e:
                    print(f"   Ошибка диагностики: {diag_e}")
        
        except Exception as e:
            print(f"❌ Ошибка при обработке файла: {e}")
        
        print("=" * 60)

def test_presentation_creation():
    """
    Тестирует создание презентации с рыбными блюдами
    """
    downloads_path = Path(r"C:\Users\katya\Downloads\Telegram Desktop")
    excel_files = list(downloads_path.glob('*.xlsx')) + list(downloads_path.glob('*.xls'))
    
    if not excel_files:
        print("❌ Нет Excel файлов для тестирования презентации")
        return
    
    # Берем первый файл с рыбными блюдами
    test_excel = None
    for excel_file in excel_files[:3]:
        dishes = extract_fish_dishes_from_column_e(str(excel_file))
        if dishes:
            test_excel = excel_file
            break
    
    if not test_excel:
        print("❌ Не найден Excel файл с рыбными блюдами для тестирования презентации")
        return
    
    print(f"\n🎯 ТЕСТИРОВАНИЕ СОЗДАНИЯ ПРЕЗЕНТАЦИИ")
    print(f"Используем файл: {test_excel.name}")
    
    # Импортируем функцию создания презентации
    from presentation_handler import create_presentation_with_fish_and_side_dishes
    
    # Проверяем наличие шаблона
    template_path = Path("template.pptx")
    if not template_path.exists():
        print(f"❌ Шаблон презентации {template_path} не найден")
        # Поищем шаблон в других местах
        possible_templates = [
            Path("templates/template.pptx"),
            Path("../template.pptx"),
            Path("presentation_template.pptx")
        ]
        
        for t in possible_templates:
            if t.exists():
                template_path = t
                print(f"✅ Найден шаблон: {template_path}")
                break
        else:
            print("❌ Шаблон презентации не найден. Создание презентации невозможно.")
            return
    
    # Создаем тестовую презентацию
    output_path = Path("test_fish_presentation.pptx")
    
    try:
        success, message = create_presentation_with_fish_and_side_dishes(
            str(template_path),
            str(test_excel),
            str(output_path)
        )
        
        if success:
            print(f"✅ Презентация успешно создана: {output_path}")
            print(f"Сообщение: {message}")
            
            if output_path.exists():
                size = output_path.stat().st_size
                print(f"Размер файла: {size} байт")
        else:
            print(f"❌ Ошибка создания презентации: {message}")
            
    except Exception as e:
        print(f"❌ Исключение при создании презентации: {e}")

if __name__ == "__main__":
    print("🧪 ТЕСТИРОВАНИЕ ИЗВЛЕЧЕНИЯ РЫБНЫХ БЛЮД")
    print("=" * 60)
    
    test_excel_files_in_downloads()
    
    print("\n" + "=" * 60)
    test_presentation_creation()
