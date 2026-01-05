"""
Centralized UI styling module for the Menu application.

This module provides all styling constants, font definitions, button presets,
and utility functions for consistent UI appearance across the application.
"""

from PySide6.QtCore import QSize
from PySide6.QtGui import QFont, QPalette, QColor, QIcon, QPixmap, QPainter, QPen, QBrush, QLinearGradient
from PySide6.QtWidgets import QSizePolicy
from PySide6.QtCore import Qt
from typing import Dict, Any, Optional


class AppStyles:
    """Основные константы стилей приложения и пресеты.
    
    Содержит все основные параметры внешнего вида приложения:
    размеры окон, отступы, шрифты, цвета и размеры компонентов.
    Все значения централизованы для обеспечения консистентности.
    """
    
    # === WINDOW SETTINGS ===
    # По умолчанию более широкая ширина окна, чтобы кнопки верхней панели помещались
    WINDOW_DEFAULT_SIZE = (1520, 800)
    # Минимальный размер окна (по ширине — чтобы все кнопки верхней панели помещались)
    WINDOW_MIN_SIZE = (1520, 650)
    WINDOW_ICON_SIZE = 256
    
    # === SPACING AND MARGINS ===
    DEFAULT_MARGIN = 12
    DEFAULT_SPACING = 10
    CONTENT_SPACING = 8
    COMPACT_SPACING = 4
    
    # === FONT SETTINGS ===
    DEFAULT_FONT_FAMILY = "Segoe UI"
    DEFAULT_FONT_SIZE = 14
    CAPTION_FONT_WEIGHT = True  # Bold
    BUTTON_FONT_SIZE = 14
    # Чуть более жирный шрифт для лучшей читаемости кнопок
    BUTTON_FONT_WEIGHT = 700

    # === CONTROL SIZES ===
    # Единая высота для кнопок/инпутов, чтобы интерфейс выглядел гармонично.
    # (чем меньше — тем компактнее интерфейс)
    CONTROL_HEIGHT = 20
    BUTTON_HEIGHT = CONTROL_HEIGHT

    # === BORDER RADIUS ===
    DEFAULT_BORDER_RADIUS = 8
    SMALL_BORDER_RADIUS = 6
    TINY_BORDER_RADIUS = 4

    # === BUTTON PADDING ===
    # Более компактные отступы, чтобы кнопки были поменьше по высоте и ширине
    DEFAULT_PADDING = "4px 8px"
    COMPACT_PADDING = "3px 6px"
    LARGE_PADDING = "8px 16px"
    
    # === COMPONENT HEIGHTS ===
    # Важно: при повышенном DPI (125%/150%) слишком маленькие фиксированные высоты
    # «сминают» содержимое групп. Поэтому делаем чуть больше.
    FILE_GROUP_MIN_HEIGHT = 70
    EXCEL_GROUP_MIN_HEIGHT = 140
    PARAMS_GROUP_MAX_HEIGHT = 90
    SPINBOX_MIN_HEIGHT = CONTROL_HEIGHT
    
    # === ICON GRADIENT COLORS ===
    ICON_GRADIENT_START = "#FF7E5F"
    ICON_GRADIENT_END = "#FD3A69"
    ICON_BORDER_COLOR = (255, 255, 255, 230)
    ICON_BORDER_WIDTH = 6
    ICON_TEXT_COLOR = (255, 255, 255)
    ICON_FONT_FAMILY = "Segoe UI"
    ICON_FONT_SIZE = 120
    
    @classmethod
    def create_app_icon(cls) -> QIcon:
        """Создает иконку приложения с градиентным фоном.
        
        Создает круглую иконку с градиентом от оранжевого к розовому,
        светлой каймой и буквой "М" по центру.
        
        Returns:
            QIcon: Готовая иконка приложения размером 256x256 пикселей
            
        Note:
            Требует инициализированного QGuiApplication для создания QPixmap
        """
        size = cls.WINDOW_ICON_SIZE
        pix = QPixmap(size, size)
        pix.fill(Qt.transparent)
        p = QPainter(pix)
        try:
            p.setRenderHint(QPainter.Antialiasing, True)
            
            # Background circle with gradient
            grad = QLinearGradient(0, 0, size, size)
            grad.setColorAt(0.0, QColor(cls.ICON_GRADIENT_START))
            grad.setColorAt(1.0, QColor(cls.ICON_GRADIENT_END))
            p.setBrush(QBrush(grad))
            p.setPen(Qt.NoPen)
            margin = 12
            p.drawEllipse(margin, margin, size - 2 * margin, size - 2 * margin)

            # Light border
            p.setPen(QPen(QColor(*cls.ICON_BORDER_COLOR), cls.ICON_BORDER_WIDTH))
            p.setBrush(Qt.NoBrush)
            p.drawEllipse(margin + 3, margin + 3, size - 2 * (margin + 3), size - 2 * (margin + 3))

            # Letter "M"
            f = QFont()
            f.setFamily(cls.ICON_FONT_FAMILY)
            f.setBold(True)
            f.setPointSize(cls.ICON_FONT_SIZE)
            p.setFont(f)
            p.setPen(QColor(*cls.ICON_TEXT_COLOR))
            p.drawText(pix.rect(), Qt.AlignCenter, "М")
        finally:
            p.end()
        return QIcon(pix)


class ButtonStyles:
    """Пресеты стилей кнопок и утилиты для их применения.
    
    Содержит предустановленные стили для различных типов кнопок:
    кнопки панели инструментов, кнопки обзора, кнопки действий.
    Предоставляет утилиты для применения стилей к кнопкам.
    """
    
    # === BUTTON PADDING ===
    # Используем более компактные отступы, чтобы кнопки были поменьше
    DEFAULT_PADDING = "4px 8px"
    COMPACT_PADDING = "3px 6px"
    LARGE_PADDING = "8px 16px"
    
    # === BUTTON PRESETS ===
    TOOLBAR_BUTTON = {
        "padding": DEFAULT_PADDING,
        "font-size": f"{AppStyles.BUTTON_FONT_SIZE}px",
        "font-weight": AppStyles.BUTTON_FONT_WEIGHT,
    }
    
    BROWSE_BUTTON = {
        "padding": DEFAULT_PADDING,
        "min-width": "80px",
    }
    
    ACTION_BUTTON = {
        "padding": DEFAULT_PADDING,
        "font-size": f"{AppStyles.BUTTON_FONT_SIZE}px",
        "font-weight": AppStyles.BUTTON_FONT_WEIGHT,
        "min-width": "120px",
    }

    # Более компактные кнопки для раздела "Документы"
    DOC_BUTTON = {
        "padding": DEFAULT_PADDING,
        "font-size": f"{AppStyles.BUTTON_FONT_SIZE}px",
        "font-weight": AppStyles.BUTTON_FONT_WEIGHT,
        "min-width": "90px",
    }
    
    @classmethod
    def apply_button_style(cls, button, style_preset: Dict[str, Any]) -> None:
        """Применяет пресет стилей к кнопке QPushButton.
        
        Args:
            button: Кнопка QPushButton для стилизации
            style_preset: Словарь со CSS-свойствами и их значениями
                         Например: {"padding": "6px 12px", "color": "white"}
        
        Note:
            CSS-свойства применяются как stylesheet к кнопке
        """
        if hasattr(button, 'setStyleSheet'):
            css_rules = []
            for prop, value in style_preset.items():
                css_rules.append(f"{prop}: {value}")
            
            css = "QPushButton { " + "; ".join(css_rules) + " }"
            button.setStyleSheet(css)


class LayoutStyles:
    """Константы стилей макетов и контейнеров.
    
    Предоставляет предустановленные значения отступов, размерных политик
    и утилиты для их применения к макетам и виджетам.
    """
    
    # === LAYOUT MARGINS ===
    NO_MARGINS = (0, 0, 0, 0)
    DEFAULT_MARGINS = (AppStyles.DEFAULT_MARGIN, AppStyles.DEFAULT_MARGIN, 
                      AppStyles.DEFAULT_MARGIN, AppStyles.DEFAULT_MARGIN)
    TOPBAR_MARGINS = (AppStyles.DEFAULT_MARGIN, 8, AppStyles.DEFAULT_MARGIN, 8)
    MINIMAL_TOP_MARGIN = (0, 2, 0, 0)
    CONTENT_TOP_MARGIN = (0, 8, 0, 0)
    
    # === SIZE POLICIES ===
    EXPANDING_FIXED = (QSizePolicy.Expanding, QSizePolicy.Fixed)
    EXPANDING_EXPANDING = (QSizePolicy.Expanding, QSizePolicy.Expanding)
    
    @classmethod
    def apply_margins(cls, layout, margins_preset) -> None:
        """Применяет пресет отступов к макету.
        
        Args:
            layout: Макет для применения отступов (должен иметь метод setContentsMargins)
            margins_preset: Кортеж с отступами (left, top, right, bottom)
        """
        if hasattr(layout, 'setContentsMargins'):
            layout.setContentsMargins(*margins_preset)
    
    @classmethod
    def apply_size_policy(cls, widget, policy_preset) -> None:
        """Применяет пресет размерной политики к виджету.
        
        Args:
            widget: Виджет для применения политики (должен иметь метод setSizePolicy)
            policy_preset: Кортеж с политиками (horizontal_policy, vertical_policy)
        """
        if hasattr(widget, 'setSizePolicy'):
            widget.setSizePolicy(*policy_preset)


class StyleSheets:
    """Централизованные определения таблиц стилей.
    
    Генерирует CSS-стили для различных частей приложения,
    используя константы из других классов стилей.
    """
    
    @classmethod
    def get_main_stylesheet(cls) -> str:
        """Возвращает основную таблицу стилей приложения.
        
        Генерирует CSS для панели управления, кнопок действий,
        групп параметров и других основных элементов интерфейса.
        
        Returns:
            str: CSS-код основных стилей приложения
        """
        return f"""
            /* Global typography + control sizing */
            QWidget {{
                font-family: "{AppStyles.DEFAULT_FONT_FAMILY}";
                font-size: {AppStyles.DEFAULT_FONT_SIZE}px;
            }}
            QPushButton {{
                min-height: {AppStyles.BUTTON_HEIGHT}px;
                padding: {ButtonStyles.DEFAULT_PADDING};
                font-size: {AppStyles.BUTTON_FONT_SIZE}px;
                font-weight: {AppStyles.BUTTON_FONT_WEIGHT};
            }}
            QLineEdit, QComboBox, QSpinBox {{
                min-height: {AppStyles.CONTROL_HEIGHT}px;
                font-size: {AppStyles.DEFAULT_FONT_SIZE}px;
            }}

            #topBar {{
                border: 1px solid palette(Mid);
                border-radius: {AppStyles.DEFAULT_BORDER_RADIUS}px;
                background: palette(Base);
            }}
            #topBar QComboBox {{
                padding: {ButtonStyles.COMPACT_PADDING};
                min-width: 160px;
                font-size: {AppStyles.DEFAULT_FONT_SIZE}px;
            }}
            #topBar QLabel {{
                font-weight: {AppStyles.BUTTON_FONT_WEIGHT};
            }}
            /* Action panel buttons */
            #actionsPanel QPushButton {{
                /* use global QPushButton settings */
            }}
            /* Parameter group styling */
            QGroupBox#paramsBox {{
                border: none;
                margin: 0px;
                padding: 0px;
                font-weight: {AppStyles.BUTTON_FONT_WEIGHT};
            }}
            QGroupBox#paramsBox::title {{
                subcontrol-origin: content;
                subcontrol-position: top left;
                left: 0px;
                top: -2px;
                padding: 0px;
                margin: 0px;
            }}
            /* Compact elements inside parameters */
            #paramsFrame QCheckBox, #paramsFrame QLabel {{
                padding: 6px 2px;
                margin: 4px 8px 4px 0px;
            }}
            #paramsFrame QCheckBox::indicator {{
                width: 14px;
                height: 14px;
            }}
            #paramsFrame QSpinBox {{
                min-height: {AppStyles.SPINBOX_MIN_HEIGHT}px;
                padding: 2px 4px;
            }}
            /* Remove margins from parameters container */
            #paramsFrame {{
                border: none;
                padding: 0px;
                margin: 0px;
            }}
        """


class ComponentStyles:
    """Утилиты стилизации отдельных компонентов.
    
    Предоставляет функции для применения стандартных стилей
    к различным элементам интерфейса: меткам, группам файлов,
    параметрическим группам и другим компонентам.
    """
    
    @classmethod
    def style_caption_label(cls, label) -> None:
        """Применяет жирное начертание шрифта к меткам-заголовкам.
        
        Args:
            label: QLabel для стилизации (должен иметь методы font() и setFont())
        """
        if hasattr(label, 'font') and hasattr(label, 'setFont'):
            font = label.font()
            font.setBold(AppStyles.CAPTION_FONT_WEIGHT)
            label.setFont(font)
    
    @classmethod
    def style_file_group(cls, group_box) -> None:
        """Применяет стандартную стилизацию к группам выбора файлов.

        На Windows при увеличенном DPI фиксированная вертикальная политика
        может приводить к «сжатию» содержимого. Поэтому делаем группу
        растягиваемой по высоте внутри scroll area.
        """
        LayoutStyles.apply_size_policy(group_box, LayoutStyles.EXPANDING_EXPANDING)
        group_box.setMinimumHeight(AppStyles.FILE_GROUP_MIN_HEIGHT)
    
    @classmethod
    def style_excel_group(cls, group_box) -> None:
        """Применяет стилизацию к группам выбора Excel-файлов.

        Аналогично style_file_group: не фиксируем по высоте, чтобы
        не «сминалось» при DPI scaling.
        """
        LayoutStyles.apply_size_policy(group_box, LayoutStyles.EXPANDING_EXPANDING)
        group_box.setMinimumHeight(AppStyles.EXCEL_GROUP_MIN_HEIGHT)
    
    @classmethod
    def style_params_group(cls, group_box) -> None:
        """Применяет стилизацию к группам параметров.
        
        Настраивает группу как сворачиваемую, устанавливает идентификатор,
        размерную политику и максимальную высоту для компактного отображения.
        
        Args:
            group_box: QGroupBox для стилизации как группы параметров
        """
        group_box.setObjectName("paramsBox")
        group_box.setCheckable(True)
        group_box.setChecked(False)
        LayoutStyles.apply_size_policy(group_box, LayoutStyles.EXPANDING_FIXED)
        group_box.setMaximumHeight(AppStyles.PARAMS_GROUP_MAX_HEIGHT)


class ThemeAwareStyles:
    """Стили, адаптирующиеся к светлой/темной теме.
    
    Предоставляет функции для получения цветов и стилей,
    которые автоматически адаптируются к текущей теме интерфейса.
    """
    
    @classmethod
    def get_border_color(cls, is_dark: bool) -> str:
        """Возвращает подходящий цвет границы для текущей темы.
        
        Args:
            is_dark: True для темной темы, False для светлой
            
        Returns:
            str: Hex-код цвета границы (#3a3a3a для темной, #c0c0c0 для светлой)
        """
        return "#3a3a3a" if is_dark else "#c0c0c0"
    
    @classmethod
    def get_tooltip_border_color(cls, is_dark: bool) -> str:
        """Возвращает подходящий цвет границы подсказок для текущей темы.
        
        Args:
            is_dark: True для темной темы, False для светлой
            
        Returns:
            str: Hex-код цвета границы подсказок (#5a5a5a для темной, #a0a0a0 для светлой)
        """
        return "#5a5a5a" if is_dark else "#a0a0a0"
    
    @classmethod
    def get_theme_stylesheet(cls, is_dark: bool) -> str:
        """Возвращает тематические дополнения к таблице стилей.
        
        Генерирует CSS для элементов, которые должны адаптироваться
        к текущей теме: границы групп, подсказки и другие элементы.
        
        Args:
            is_dark: True для темной темы, False для светлой
            
        Returns:
            str: CSS-код тематических стилей
        """
        border = cls.get_border_color(is_dark)
        tooltip_border = cls.get_tooltip_border_color(is_dark)

        # Календарь: в светлой теме — бежевый, в тёмной — тёмный фон + бежевое выделение.
        # Важно: выделение делаем бежевым, чтобы не было "синего" (акцент Windows).
        if is_dark:
            cal_bg = "#2b2d30"
            cal_text = "#e6e6e6"
            cal_sel_bg = "#c2a77d"   # бежевое выделение
            cal_sel_text = "#1b1b1b" # тёмный текст на бежевом
            cal_hover_bg = "#35373c"
            cal_disabled_text = "#808080"
        else:
            cal_bg = "#f5f0e6"
            cal_text = "#1b1b1b"
            cal_sel_bg = "#e2c9a7"   # бежевое выделение
            cal_sel_text = "#1b1b1b"
            cal_hover_bg = "#eadcc8"
            cal_disabled_text = "#9a9a9a"

        return f"""
        QGroupBox {{
            border: 1px solid {border};
            border-radius: {AppStyles.SMALL_BORDER_RADIUS}px;
            margin-top: 10px;
            padding-top: 6px;
        }}
        QGroupBox::title {{
            subcontrol-origin: margin;
            subcontrol-position: top left;
            padding: 0 6px;
            font-weight: bold;
        }}
        QToolTip {{
            border: 1px solid {tooltip_border};
            padding: 4px;
            border-radius: {AppStyles.TINY_BORDER_RADIUS}px;
        }}

        QCalendarWidget {{
            border: 1px solid {border};
            border-radius: {AppStyles.SMALL_BORDER_RADIUS}px;
        }}
        QCalendarWidget QWidget {{
            background-color: {cal_bg};
            color: {cal_text};
        }}
        QCalendarWidget QAbstractItemView {{
            background-color: {cal_bg};
            color: {cal_text};
            selection-background-color: {cal_sel_bg};
            selection-color: {cal_sel_text};
        }}
        QCalendarWidget QAbstractItemView:disabled {{
            color: {cal_disabled_text};
        }}
        QCalendarWidget QToolButton {{
            background-color: {cal_bg};
            color: {cal_text};
            border: none;
            padding: 4px 8px;
        }}
        QCalendarWidget QToolButton:hover {{
            background-color: {cal_hover_bg};
        }}
        QCalendarWidget QMenu {{
            background-color: {cal_bg};
            color: {cal_text};
        }}
        QCalendarWidget QSpinBox {{
            background-color: {cal_bg};
            color: {cal_text};
        }}
        """


class StyleManager:
    """Основной класс управления стилями.
    
    Предоставляет высокоуровневые функции для стилизации окон и кнопок.
    Объединяет функциональность других классов стилей в удобном API.
    """
    
    @classmethod
    def setup_main_window(cls, window) -> None:
        """Применяет всю стилизацию к главному окну."""
        if hasattr(window, 'setWindowIcon'):
            window.setWindowIcon(AppStyles.create_app_icon())

        if hasattr(window, 'resize'):
            window.resize(*AppStyles.WINDOW_DEFAULT_SIZE)

        # Минимальный размер, чтобы интерфейс не «сминался»
        if hasattr(window, 'setMinimumSize'):
            window.setMinimumSize(*AppStyles.WINDOW_MIN_SIZE)

        if hasattr(window, 'setStyleSheet'):
            window.setStyleSheet(StyleSheets.get_main_stylesheet())
    
    @classmethod
    def style_toolbar_button(cls, button) -> None:
        """Применяет стиль кнопки панели инструментов.
        
        Args:
            button: QPushButton для стилизации как кнопка панели инструментов
        """
        ButtonStyles.apply_button_style(button, ButtonStyles.TOOLBAR_BUTTON)
    
    @classmethod
    def style_action_button(cls, button) -> None:
        """Применяет стиль кнопки действия.
        
        Args:
            button: QPushButton для стилизации как кнопка действия
        """
        ButtonStyles.apply_button_style(button, ButtonStyles.ACTION_BUTTON)

    @classmethod
    def style_doc_button(cls, button) -> None:
        """Применяет компактный стиль кнопки для раздела "Документы"."""
        ButtonStyles.apply_button_style(button, ButtonStyles.DOC_BUTTON)
    
    @classmethod
    def style_browse_button(cls, button) -> None:
        """Применяет стиль кнопки обзора файлов.
        
        Args:
            button: QPushButton для стилизации как кнопка обзора
        """
        ButtonStyles.apply_button_style(button, ButtonStyles.BROWSE_BUTTON)
