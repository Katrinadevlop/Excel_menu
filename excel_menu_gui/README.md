# Excel Menu GUI

Небольшое настольное приложение для Windows, которое помогает формировать меню на основе Excel‑шаблонов (папка `templates/`), сравнивать списки блюд и заполнять готовые формы.

## Требования и установка
- Windows, Python 3.9+
- Установите зависимости:
```
pip install -r requirements.txt
```
В проекте используется библиотека `pandas` (импортируется в `dish_extractor.py` и `menu_template_filler.py`), поэтому она добавлена в `requirements.txt`.

## Запуск приложения
Из корневой директории проекта выполните:
```
python main.py
```

## Запуск тестов
Автоматические тесты находятся в `tests/` и запускаются командой:
```
python -m unittest discover -s tests -v
```
GUI‑тесты, требующие ручного запуска, находятся в `manual_tests/` и не входят в автоматический прогон.

## Краткая структура проекта
- Корень (основной код, импорты не менялись): `main.py`, `comparator.py`, `menu_template_filler.py`, `dish_extractor.py`, `brokerage_journal.py`, `presentation_handler.py`, `template_linker.py`, `theme.py`, `ui_styles.py`
- `tests/` — автоматические тесты (без открытия GUI)
- `manual_tests/` — ручные/GUI‑тесты
- `templates/` — шаблоны Excel (без изменений)
- Служебные скрипты сборки: `build_exe.bat`, `create_distribution.bat`, `install.bat`, `menu_app.spec`

Примечание: импорты модулей выполняются напрямую из корня, поэтому запуск приложения и тестов рекомендуется производить из корневой директории проекта.

## Сборка .exe
Инструкции по сборке исполняемого файла могут быть добавлены позже при необходимости.
