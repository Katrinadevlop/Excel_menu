#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from brokerage_journal import create_brokerage_journal_from_menu
import os

def test_improved_extraction():
    """Test improved salad and right column extraction."""
    menu_file = 'templates/Шаблон меню пример.xlsx'
    template_file = 'templates/Бракеражный журнал шаблон.xlsx'
    output_file = 'test_improved_extraction.xlsx'

    success, message = create_brokerage_journal_from_menu(menu_file, template_file, output_file)
    
    print(f'Success: {success}')
    print(f'Message: {message}')
    
    if os.path.exists(output_file):
        print(f'Output file created: {output_file}')
    
    return success

if __name__ == '__main__':
    test_improved_extraction()
