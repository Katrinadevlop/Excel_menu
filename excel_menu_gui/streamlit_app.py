"""
Веб-приложение для работы с меню
Streamlit версия десктопного приложения
"""

import streamlit as st
import tempfile
import os
from pathlib import Path
from datetime import date

# Настройка страницы
st.set_page_config(
    page_title="Работа с меню",
    page_icon="🍽️",
    layout="wide",
    initial_sidebar_state="collapsed",  # сворачиваем боковую панель
)

# Добавляем путь к модулям приложения
import sys
sys.path.insert(0, str(Path(__file__).parent))

# Импорты из существующего кода
from app.services.comparator import compare_and_highlight, get_sheet_names, ColumnParseError
from app.reports.presentation_handler import create_presentation_with_excel_data
from app.reports.brokerage_journal import create_brokerage_journal_from_menu
from app.services.menu_template_filler import MenuTemplateFiller
from app.services.template_linker import default_template_path


def find_template(filename: str) -> str | None:
    """Ищет шаблон в директории templates"""
    base = Path(__file__).parent
    candidates = [
        base / "templates" / filename,
        base / "excel_menu_gui" / "templates" / filename,
    ]
    for p in candidates:
        if p.exists():
            return str(p)
    return None


def save_uploaded_file(uploaded_file) -> str:
    """Сохраняет загруженный файл во временную директорию"""
    temp_dir = tempfile.mkdtemp()
    file_path = os.path.join(temp_dir, uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return file_path


def main():
    """Главная точка входа веб-приложения с простым минималистичным интерфейсом."""
    # Немного сжимаем отступы сверху/снизу
    st.markdown(
        """
        <style>
        .block-container {padding-top: 25px; padding-bottom: 10px;}
        h1, h2 {margin-bottom: 2px;}
        /* минимальные зазоры вокруг загрузчика файлов и кнопок */
        .stFileUploader {margin-top: 1px; margin-bottom: 1px;}
        .stButton {margin-top: 1px;}
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("Работа с меню")

    tabs = st.tabs([
        "Сравнение меню",
        "Презентация",
        "Бракеражный журнал",
        "Шаблон меню",
        "Шаблоны",
    ])

    with tabs[0]:
        compare_menus_page()
    with tabs[1]:
        create_presentation_page()
    with tabs[2]:
        brokerage_journal_page()
    with tabs[3]:
        fill_template_page()
    with tabs[4]:
        download_template_page()


def compare_menus_page():
    """Страница сравнения меню"""
    st.header("Сравнение меню")

    # Два файла вертикально списком
    file1 = st.file_uploader(
        "Первый файл",
        type=["xlsx", "xls", "xlsm"],
        key="file1"
    )

    file2 = st.file_uploader(
        "Второй файл",
        type=["xlsx", "xls", "xlsm"],
        key="file2"
    )

    # Параметры сразу под загрузкой файлов, в одной колонке
    st.markdown("**Параметры**")
    ignore_case = st.checkbox("Игнорировать регистр", value=True)
    use_fuzzy = st.checkbox("Использовать нечёткое сравнение")

    # Ползунок уже включен всегда, просто используется только если включено нечёткое сравнение
    slider_col, _ = st.columns([1, 2])
    with slider_col:
        fuzzy_threshold = st.slider(
            "Порог похожести",
            min_value=50,
            max_value=100,
            value=85,
        )

    if st.button("Сравнить и подсветить", type="primary", disabled=not (file1 and file2)):
        if file1 and file2:
            with st.spinner("Сравнение файлов..."):
                try:
                    # Сохраняем загруженные файлы
                    path1 = save_uploaded_file(file1)
                    path2 = save_uploaded_file(file2)
                    
                    # Определяем листы
                    sheets1 = get_sheet_names(path1)
                    sheets2 = get_sheet_names(path2)
                    
                    # Выбираем лист с "касс" или первый
                    def select_sheet(sheets):
                        for s in sheets:
                            if "касс" in s.lower():
                                return s
                        return sheets[0] if sheets else None
                    
                    sheet1 = select_sheet(sheets1)
                    sheet2 = select_sheet(sheets2)
                    
                    if not sheet1 or not sheet2:
                        st.error("Не удалось определить листы для сравнения")
                        return
                    
                    # Выполняем сравнение
                    out_path, matches = compare_and_highlight(
                        path1=path1, sheet1=sheet1,
                        path2=path2, sheet2=sheet2,
                        col1="A", col2="E",
                        header_row1=1, header_row2=1,
                        ignore_case=ignore_case,
                        use_fuzzy=use_fuzzy,
                        fuzzy_threshold=fuzzy_threshold,
                        final_choice=0
                    )
                    
                    st.success(f"Сравнение завершено. Найдено совпадений: {matches}")
                    
                    # Кнопка скачивания результата
                    with open(out_path, "rb") as f:
                            st.download_button(
                                label="Скачать результат",
                            data=f,
                            file_name=f"сравнение_меню_{date.today().strftime('%d.%m.%Y')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                
                except ColumnParseError as e:
                    st.error(f"Ошибка колонки: {e}")
                except Exception as e:
                    st.error(f"Ошибка: {e}")


def create_presentation_page():
    """Страница создания презентации"""
    st.header("Создание презентации")
    
    excel_file = st.file_uploader(
        "",
        type=["xlsx", "xls", "xlsm"],
        key="excel_presentation"
    )
    
    if st.button("Создать презентацию", type="primary", disabled=not excel_file):
        if excel_file:
            with st.spinner("Создание презентации..."):
                try:
                    # Сохраняем файл
                    excel_path = save_uploaded_file(excel_file)
                    
                    # Ищем шаблон презентации
                    template_path = find_template("presentation_template.pptx")
                    if not template_path:
                        st.error("Шаблон презентации не найден")
                        return
                    
                    # Создаём временный файл для результата
                    temp_dir = tempfile.mkdtemp()
                    output_path = os.path.join(temp_dir, f"презентация_меню_{date.today().strftime('%d.%m.%Y')}.pptx")
                    
                    # Создаём презентацию (сигнатура: template_path, excel_path, output_path)
                    success, message = create_presentation_with_excel_data(
                        template_path,
                        excel_path,
                        output_path,
                    )
                    
                    if success:
                        st.success(message)
                        
                        with open(output_path, "rb") as f:
                            st.download_button(
                                label="Скачать презентацию",
                                data=f,
                                file_name=f"презентация_меню_{date.today().strftime('%d.%m.%Y')}.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )
                    else:
                        st.error(f"Ошибка: {message}")
                
                except Exception as e:
                    st.error(f"Ошибка: {e}")


def brokerage_journal_page():
    """Страница создания бракеражного журнала"""
    st.header("Бракеражный журнал")
    
    menu_file = st.file_uploader(
        "",
        type=["xlsx", "xls", "xlsm"],
        key="menu_brokerage"
    )
    
    if st.button("Создать журнал", type="primary", disabled=not menu_file):
        if menu_file:
            with st.spinner("Создание бракеражного журнала..."):
                try:
                    # Сохраняем файл
                    menu_path = save_uploaded_file(menu_file)
                    
                    # Ищем шаблон журнала
                    template_path = find_template("Бракеражный журнал шаблон.xlsx")
                    if not template_path:
                        st.error("Шаблон бракеражного журнала не найден")
                        return
                    
                    # Создаём временный файл для результата
                    temp_dir = tempfile.mkdtemp()
                    output_path = os.path.join(temp_dir, f"бракеражный_журнал_{date.today().strftime('%d.%m.%Y')}.xlsx")
                    
                    # Создаём журнал
                    success, message = create_brokerage_journal_from_menu(
                        menu_path, template_path, output_path
                    )
                    
                    if success:
                        st.success(message)
                        
                        with open(output_path, "rb") as f:
                            st.download_button(
                                label="Скачать журнал",
                                data=f,
                                file_name=f"бракеражный_журнал_{date.today().strftime('%d.%m.%Y')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.error(f"Ошибка: {message}")
                
                except Exception as e:
                    st.error(f"Ошибка: {e}")


def fill_template_page():
    """Страница заполнения шаблона меню"""
    st.header("Заполнение шаблона меню")
    
    source_file = st.file_uploader(
        "",
        type=["xlsx", "xls", "xlsm"],
        key="source_menu"
    )
    
    if st.button("Заполнить шаблон", type="primary", disabled=not source_file):
        if source_file:
            with st.spinner("Заполнение шаблона меню..."):
                try:
                    # Сохраняем файл
                    source_path = save_uploaded_file(source_file)
                    
                    # Ищем шаблон меню
                    template_path = default_template_path()
                    if not template_path or not Path(template_path).exists():
                        template_path = find_template("Шаблон меню пример.xlsx")
                    
                    if not template_path:
                        st.error("Шаблон меню не найден")
                        return
                    
                    # Создаём временный файл для результата
                    temp_dir = tempfile.mkdtemp()
                    output_path = os.path.join(temp_dir, f"меню_{date.today().strftime('%d.%m.%Y')}.xlsx")
                    
                    # Заполняем шаблон
                    filler = MenuTemplateFiller()
                    success, message = filler.fill_menu_template(
                        template_path, source_path, output_path
                    )
                    
                    if success:
                        st.success(message)
                        
                        with open(output_path, "rb") as f:
                            st.download_button(
                                label="Скачать заполненный шаблон",
                                data=f,
                                file_name=f"меню_{date.today().strftime('%d.%m.%Y')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.error(f"Ошибка: {message}")
                
                except Exception as e:
                    st.error(f"Ошибка: {e}")


def download_template_page():
    """Страница скачивания шаблонов"""
    st.header("Скачать шаблоны")

    # Собираем все доступные шаблоны в один список
    items = []

    # Основной шаблон меню из template_linker (если есть)
    default_tpl = default_template_path()
    if default_tpl and Path(default_tpl).exists():
        items.append(("Основной шаблон меню", default_tpl, "шаблон_меню.xlsx"))

    # Остальные файлы из папки templates
    for name, filename in [
        ("Шаблон меню", "Шаблон меню пример.xlsx"),
        ("Шаблон бракеражного журнала", "Бракеражный журнал шаблон.xlsx"),
        ("Шаблон презентации", "presentation_template.pptx"),
    ]:
        template_path = find_template(filename)
        if template_path and Path(template_path).exists():
            items.append((name, template_path, filename))

    if not items:
        st.warning("Шаблоны не найдены")
        return

    # Кнопки вертикальным списком в левой узкой колонке, одинаковой ширины
    col, spacer = st.columns([1, 3])
    with col:
        for name, path, download_name in items:
            with open(path, "rb") as f:
                mime_type = (
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    if download_name.endswith(".xlsx")
                    else "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                st.download_button(
                    label=name,
                    data=f,
                    file_name=download_name,
                    mime=mime_type,
                    key=name,
                    use_container_width=True,
                )


if __name__ == "__main__":
    main()
