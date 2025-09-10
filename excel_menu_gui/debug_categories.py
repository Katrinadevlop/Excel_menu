import pandas as pd
import sys
from pathlib import Path

def debug_excel_categories(excel_path: str):
    """Отладочная функция для анализа содержимого Excel файла"""
    
    if not Path(excel_path).exists():
        print(f"❌ Файл не найден: {excel_path}")
        return
    
    print(f"📁 Анализируем файл: {excel_path}")
    print("-" * 50)
    
    try:
        # Получаем список листов
        xls = pd.ExcelFile(excel_path)
        print(f"📋 Найденные листы: {xls.sheet_names}")
        
        # Выбираем лист
        sheet_name = None
        for nm in xls.sheet_names:
            if 'касс' in str(nm).strip().lower():
                sheet_name = nm
                break
        
        if sheet_name is None and xls.sheet_names:
            sheet_name = xls.sheet_names[0]
            
        print(f"🎯 Выбранный лист: {sheet_name}")
        print("-" * 50)
        
        # Читаем лист
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, dtype=object)
        print(f"📊 Размер данных: {len(df)} строк, {len(df.columns)} колонок")
        print("-" * 50)
        
        def row_text(row) -> str:
            parts = []
            for v in row:
                if pd.notna(v):
                    parts.append(str(v))
            return ' '.join(parts).strip()
        
        # Ищем категории
        categories_to_find = [
            ['САЛАТ ХОЛ'], ['САЛАТ ЗАКУСК'], ['САЛАТЫ И ХОЛОДНЫЕ ЗАКУСКИ'],
            ['ПЕРВЫЕ БЛЮДА'],
            ['БЛЮДА ИЗ МЯСА'], ['МЯСНЫЕ БЛЮДА'],
            ['БЛЮДА ИЗ ПТИЦЫ'],
            ['БЛЮДА ИЗ РЫБЫ'], ['РЫБНЫЕ БЛЮДА'],
            ['ГАРНИРЫ'], ['ГАРНИР']
        ]
        
        print("🔍 Поиск категорий в файле:")
        print("=" * 50)
        
        found_categories = []
        
        # Показываем первые 50 строк с их содержимым
        max_rows_to_show = min(50, len(df))
        for i in range(max_rows_to_show):
            row_content = row_text(df.iloc[i])
            row_upper = row_content.upper().replace('Ё', 'Е')
            
            if row_content.strip():  # Показываем только непустые строки
                print(f"Строка {i+1:2d}: {row_content[:100]}")
                
                # Проверяем на категории
                for category_keywords in categories_to_find:
                    for keyword_set in category_keywords:
                        if all(kw in row_upper for kw in keyword_set.split(' ')):
                            found_categories.append((i+1, keyword_set, row_content))
                            print(f"  ✅ НАЙДЕНА КАТЕГОРИЯ: {keyword_set}")
                            break
        
        print("=" * 50)
        
        if found_categories:
            print("🎉 Найденные категории:")
            for row_num, category, content in found_categories:
                print(f"  • Строка {row_num}: {category}")
                print(f"    Содержимое: {content}")
        else:
            print("❌ Категории не найдены!")
            print("\n💡 Попробуем более гибкий поиск:")
            
            # Более гибкий поиск
            flexible_keywords = [
                'салат', 'закуск', 'первые', 'блюда', 'мясн', 'птиц', 'рыб', 'гарнир'
            ]
            
            for i in range(min(30, len(df))):
                row_content = row_text(df.iloc[i])
                row_lower = row_content.lower()
                
                if row_content.strip():
                    for keyword in flexible_keywords:
                        if keyword in row_lower:
                            print(f"  Строка {i+1}: '{row_content}' содержит '{keyword}'")
                            break
        
    except Exception as e:
        print(f"❌ Ошибка при анализе файла: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        excel_path = sys.argv[1]
    else:
        excel_path = input("Введите путь к Excel файлу: ").strip().strip('"')
    
    debug_excel_categories(excel_path)
