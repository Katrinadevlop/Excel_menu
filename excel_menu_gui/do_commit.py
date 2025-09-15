#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import subprocess
import sys

def run_git_command(command):
    """Выполняет git команду и выводит результат."""
    try:
        result = subprocess.run(
            command, 
            shell=True, 
            capture_output=True, 
            text=True,
            encoding='utf-8'
        )
        if result.stdout:
            print(result.stdout)
        if result.stderr:
            print(result.stderr, file=sys.stderr)
        return result.returncode == 0
    except Exception as e:
        print(f"Ошибка: {e}")
        return False

def main():
    print("=" * 60)
    print("GIT COMMIT")
    print("=" * 60)
    
    # Добавляем все файлы
    print("\n📝 Добавление всех измененных файлов...")
    if run_git_command("git add -A"):
        print("✅ Файлы добавлены")
    else:
        print("❌ Ошибка при добавлении файлов")
        return
    
    # Проверяем статус
    print("\n📊 Текущий статус:")
    run_git_command("git status --short")
    
    # Делаем коммит
    commit_message = "Add extract_fish_dishes_by_range function and update extract_fish_dishes_from_excel to use columns E, F, G"
    print(f"\n💾 Создание коммита: {commit_message}")
    
    if run_git_command(f'git commit -m "{commit_message}"'):
        print("✅ Коммит успешно создан!")
        
        # Показываем последний коммит
        print("\n📋 Последний коммит:")
        run_git_command("git log --oneline -1")
    else:
        print("❌ Ошибка при создании коммита")
        print("Возможно, нет изменений для коммита или не настроен git")

if __name__ == "__main__":
    main()
    input("\nНажмите Enter для завершения...")
