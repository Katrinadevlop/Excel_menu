#!/usr/bin/env python3
"""
Конвертирует SVG иконку в ICO формат для использования в PyInstaller
"""

try:
    from cairosvg import svg2png
    from PIL import Image
    import io
    
    def svg_to_ico(svg_path, ico_path):
        """Конвертирует SVG файл в ICO"""
        # Читаем SVG файл
        with open(svg_path, 'r', encoding='utf-8') as f:
            svg_content = f.read()
        
        # Создаем разные размеры иконки
        sizes = [16, 32, 48, 64, 128, 256]
        images = []
        
        for size in sizes:
            # Конвертируем SVG в PNG с нужным размером
            png_data = svg2png(bytestring=svg_content.encode('utf-8'), 
                             output_width=size, output_height=size)
            
            # Создаем PIL Image из PNG данных
            img = Image.open(io.BytesIO(png_data))
            images.append(img)
        
        # Сохраняем как ICO файл
        images[0].save(ico_path, format='ICO', sizes=[(img.width, img.height) for img in images])
        print(f"Иконка успешно создана: {ico_path}")
        
    if __name__ == "__main__":
        svg_to_ico("icon.svg", "icon.ico")
        
except ImportError as e:
    print("Нужно установить дополнительные библиотеки:")
    print("pip install cairosvg pillow")
    print(f"Ошибка: {e}")