#!/usr/bin/env python3
"""
Создает ICO иконку напрямую с помощью Pillow
"""

from PIL import Image, ImageDraw
import io

def create_icon():
    """Создает красивую иконку для приложения"""
    
    # Создаем изображение 256x256
    img = Image.new('RGBA', (256, 256), color=(0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # Основной фон (зеленый градиент)
    draw.rounded_rectangle([0, 0, 255, 255], radius=32, fill=(46, 139, 87, 255))
    
    # Основная форма документа (белый прямоугольник)
    draw.rounded_rectangle([48, 32, 192, 224], radius=8, fill=(255, 255, 255, 255), outline=(31, 95, 63), width=2)
    
    # Заголовок документа (зеленая полоса)
    draw.rounded_rectangle([56, 48, 184, 72], radius=4, fill=(76, 175, 80, 255))
    
    # Строки таблицы
    colors = [(232, 245, 232, 255), (240, 248, 240, 255)]
    y_positions = [88, 108, 128, 148, 168]
    
    for i, y in enumerate(y_positions):
        color = colors[i % 2]
        draw.rounded_rectangle([56, y, 184, y+12], radius=2, fill=color)
    
    # Вертикальные разделители
    line_color = (200, 230, 201, 255)
    for x in [88, 120, 152]:
        draw.line([(x, 88), (x, 180)], fill=line_color, width=1)
    
    # Значок меню в углу (красный круг с линиями)
    draw.ellipse([176, 56, 224, 104], fill=(255, 107, 53, 255))
    # Линии меню
    for y in [74, 79, 84]:
        draw.rectangle([192, y, 208, y+2], fill=(255, 255, 255, 255))
    
    # Декоративные элементы
    draw.ellipse([64, 192, 80, 208], fill=(76, 175, 80, 180))
    draw.ellipse([162, 194, 174, 206], fill=(102, 187, 106, 180))
    
    # Создаем различные размеры для ICO файла
    sizes = [16, 32, 48, 64, 128, 256]
    images = []
    
    for size in sizes:
        resized = img.resize((size, size), Image.Resampling.LANCZOS)
        images.append(resized)
    
    # Сохраняем как ICO
    images[0].save('icon.ico', format='ICO', sizes=[(img.width, img.height) for img in images])
    print("Иконка успешно создана: icon.ico")

if __name__ == "__main__":
    create_icon()