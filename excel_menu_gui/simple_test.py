#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from brokerage_journal import BrokerageJournalGenerator

def test_with_sample_data():
    """Тестируем логику с образцом данных из вашего примера"""
    
    # Данные как в вашем примере:
    sample_dishes = [
        # Завтраки (должны попасть в категорию завтрак)
        'Блин с ветчиной и сыром',
        'Блин с яблоком', 
        'Блинчик без начинки',
        'Горячий бутерброд с ветчиной и сыром',
        'Дополнительно: ветчина,помидор,сыр',
        'Драники картофельные',
        'Жареный сулугуни с брусничным соусом',
        'Каша геркулесовая молочная',
        'Каша дружба молочная (рис, пшено)',
        'Каша рисовая на кокосовом молоке',
        'Крок-мадам',
        'Оладьи из кабачков',
        'Омлет на заказ',
        'Омлет/Омлет с помидором',
        'Пудинг Рафаэлло',
        'Сосиска жаренная/отварная',
        'Сырники',
        'Хашбраун с яйцом и беконом',
        'Яйцо жареное/отварное',
        
        # Точка остановки - сэндвичи
        'СЭНДВИЧИ',  # После этого все остальное не должно быть завтраком
        'Сэндвич с бужениной',
        'Сэндвич с ветчиной и сыром',
        'Сэндвич с куриной грудкой',
        'Сэндвич с форелью',
        'ПЕЛЬМЕНИ',
        'Пельмени Домашние'
    ]
    
    generator = BrokerageJournalGenerator()
    result = {'завтрак': [], 'салат': [], 'первое': [], 'мясо': [], 'курица': [], 'рыба': [], 'гарнир': []}
    
    breakfast_mode = True
    
    for dish in sample_dishes:
        if breakfast_mode:
            if 'СЭНДВИЧ' in dish.upper():
                breakfast_mode = False
                continue
            
            # Проверяем, является ли это валидным блюдом
            if not generator._should_skip_cell(dish) and generator._is_valid_dish(dish, []):
                result['завтрак'].append(dish)
        
    print("\n=== ТЕСТ С ОБРАЗЦОМ ДАННЫХ ===")
    print(f"Завтраки: {len(result['завтрак'])} блюд")
    for i, dish in enumerate(result['завтрак'], 1):
        print(f"  {i}. {dish}")
    
    return result

if __name__ == "__main__":
    test_with_sample_data()
