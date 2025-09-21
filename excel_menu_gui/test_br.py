#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from brokerage_journal import BrokerageJournalGenerator

# Тестируем создание бракеражного журнала
generator = BrokerageJournalGenerator()

menu_path = "templates/Шаблон меню пример.xlsx"
template_path = "templates/Бракеражный журнал шаблон.xlsx"
output_path = "test_brokerage_journal.xlsx"

success, message = generator.create_brokerage_journal(menu_path, template_path, output_path)
print(f"Результат: {'Успех' if success else 'Ошибка'}")
print(f"Сообщение: {message}")
