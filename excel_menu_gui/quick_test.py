#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from brokerage_journal import BrokerageJournalGenerator

def quick_test():
    """Quick test of the improved extraction logic."""
    generator = BrokerageJournalGenerator()
    menu_file = 'templates/Шаблон меню пример.xlsx'
    
    print("=== TESTING IMPROVED EXTRACTION LOGIC ===\n")
    
    try:
        # Test the categorized extraction
        categories = generator.extract_categorized_dishes(menu_file)
        
        print("Extraction results:")
        for category, dishes in categories.items():
            print(f"\n{category.upper()} ({len(dishes)} dishes):")
            for dish in dishes:
                print(f"  - {dish}")
        
        # Test the journal creation
        template_file = 'templates/Бракеражный журнал шаблон.xlsx'
        output_file = 'test_improved_output.xlsx'
        
        print("\n=== CREATING BROKERAGE JOURNAL ===\n")
        success, message = generator.create_brokerage_journal(menu_file, template_file, output_file)
        
        print(f"Success: {success}")
        print(f"Message: {message}")
        
    except Exception as e:
        print(f"Error during test: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    quick_test()
