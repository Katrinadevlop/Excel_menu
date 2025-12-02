#!/usr/bin/env python3
"""
Создает ICO иконку с буквой М на зелёном фоне
"""

from PIL import Image, ImageDraw, ImageFont
import io

def create_icon():
    """Создает иконку с буквой М на зелёном фоне"""
    
    # Создаем изображение 256x256
    img = Image.new('RGBA', (256, 256), color=(0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # Основной фон (зеленый)
    draw.rounded_rectangle([0, 0, 255, 255], radius=32, fill=(46, 139, 87, 255))
    
    # Пытаемся загрузить шрифт для буквы М
    try:
        font = ImageFont.truetype('arial.ttf', 180)
    except:
        try:
            font = ImageFont.truetype('C:\\Windows\\Fonts\\arial.ttf', 180)
        except:
            try:
                font = ImageFont.truetype('C:\\Windows\\Fonts\\segoeui.ttf', 180)
            except:
                # Если не удалось загрузить шрифт, рисуем букву М вручную
                # Левая вертикальная линия
                draw.rectangle([50, 60, 80, 196], fill=(255, 255, 255, 255))
                # Правая вертикальная линия
                draw.rectangle([176, 60, 206, 196], fill=(255, 255, 255, 255))
                # Левая диагональ
                draw.polygon([(80, 60), (128, 120), (108, 120), (60, 60)], fill=(255, 255, 255, 255))
                # Правая диагональ
                draw.polygon([(176, 60), (148, 120), (128, 120), (196, 60)], fill=(255, 255, 255, 255))
                font = None
    
    # Рисуем букву М белым цветом по центру (если шрифт загружен)
    if font:
        draw.text((128, 128), 'М', fill=(255, 255, 255, 255), font=font, anchor='mm')
    
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