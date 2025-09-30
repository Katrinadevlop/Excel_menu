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
    """Main application styling constants and presets."""
    
    # === WINDOW SETTINGS ===
    WINDOW_DEFAULT_SIZE = (1000, 760)
    WINDOW_MIN_SIZE = (800, 600)
    WINDOW_ICON_SIZE = 256
    
    # === SPACING AND MARGINS ===
    DEFAULT_MARGIN = 12
    DEFAULT_SPACING = 10
    CONTENT_SPACING = 8
    COMPACT_SPACING = 4
    
    # === FONT SETTINGS ===
    DEFAULT_FONT_SIZE = 14
    CAPTION_FONT_WEIGHT = True  # Bold
    BUTTON_FONT_SIZE = 14
    BUTTON_FONT_WEIGHT = 600
    
    # === BORDER RADIUS ===
    DEFAULT_BORDER_RADIUS = 8
    SMALL_BORDER_RADIUS = 6
    TINY_BORDER_RADIUS = 4
    
    # === COMPONENT HEIGHTS ===
    FILE_GROUP_MIN_HEIGHT = 45
    EXCEL_GROUP_MIN_HEIGHT = 95
    PARAMS_GROUP_MAX_HEIGHT = 55
    SPINBOX_MIN_HEIGHT = 20
    
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
        """Creates the application icon with gradient background."""
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
    """Button styling presets and utilities."""
    
    # === BUTTON PADDING ===
    DEFAULT_PADDING = "6px 12px"
    COMPACT_PADDING = "4px 8px"
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
    
    @classmethod
    def apply_button_style(cls, button, style_preset: Dict[str, Any]) -> None:
        """Apply a button style preset to a QPushButton."""
        if hasattr(button, 'setStyleSheet'):
            css_rules = []
            for prop, value in style_preset.items():
                css_rules.append(f"{prop}: {value}")
            
            css = "QPushButton { " + "; ".join(css_rules) + " }"
            button.setStyleSheet(css)


class LayoutStyles:
    """Layout and container styling constants."""
    
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
        """Apply margin preset to a layout."""
        if hasattr(layout, 'setContentsMargins'):
            layout.setContentsMargins(*margins_preset)
    
    @classmethod
    def apply_size_policy(cls, widget, policy_preset) -> None:
        """Apply size policy preset to a widget."""
        if hasattr(widget, 'setSizePolicy'):
            widget.setSizePolicy(*policy_preset)


class StyleSheets:
    """Centralized stylesheet definitions."""
    
    @classmethod
    def get_main_stylesheet(cls) -> str:
        """Returns the main application stylesheet."""
        return f"""
            #topBar {{
                border: 1px solid palette(Mid);
                border-radius: {AppStyles.DEFAULT_BORDER_RADIUS}px;
                background: palette(Base);
            }}
            #topBar QPushButton {{
                padding: {ButtonStyles.DEFAULT_PADDING};
                font-size: {AppStyles.BUTTON_FONT_SIZE}px;
                font-weight: {AppStyles.BUTTON_FONT_WEIGHT};
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
                padding: {ButtonStyles.DEFAULT_PADDING};
                font-size: {AppStyles.BUTTON_FONT_SIZE}px;
                font-weight: {AppStyles.BUTTON_FONT_WEIGHT};
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
    """Individual component styling utilities."""
    
    @classmethod
    def style_caption_label(cls, label) -> None:
        """Apply bold font styling to caption labels."""
        if hasattr(label, 'font') and hasattr(label, 'setFont'):
            font = label.font()
            font.setBold(AppStyles.CAPTION_FONT_WEIGHT)
            label.setFont(font)
    
    @classmethod
    def style_file_group(cls, group_box) -> None:
        """Apply standard styling to file selection groups."""
        LayoutStyles.apply_size_policy(group_box, LayoutStyles.EXPANDING_FIXED)
        group_box.setMinimumHeight(AppStyles.FILE_GROUP_MIN_HEIGHT)
    
    @classmethod
    def style_excel_group(cls, group_box) -> None:
        """Apply styling to Excel file selection groups."""
        LayoutStyles.apply_size_policy(group_box, LayoutStyles.EXPANDING_FIXED)
        group_box.setMinimumHeight(AppStyles.EXCEL_GROUP_MIN_HEIGHT)
    
    @classmethod
    def style_params_group(cls, group_box) -> None:
        """Apply styling to parameter groups."""
        group_box.setObjectName("paramsBox")
        group_box.setCheckable(True)
        group_box.setChecked(False)
        LayoutStyles.apply_size_policy(group_box, LayoutStyles.EXPANDING_FIXED)
        group_box.setMaximumHeight(AppStyles.PARAMS_GROUP_MAX_HEIGHT)


class ThemeAwareStyles:
    """Theme-aware styling that adapts to light/dark themes."""
    
    @classmethod
    def get_border_color(cls, is_dark: bool) -> str:
        """Get appropriate border color for current theme."""
        return "#3a3a3a" if is_dark else "#c0c0c0"
    
    @classmethod
    def get_tooltip_border_color(cls, is_dark: bool) -> str:
        """Get appropriate tooltip border color for current theme."""
        return "#5a5a5a" if is_dark else "#a0a0a0"
    
    @classmethod
    def get_theme_stylesheet(cls, is_dark: bool) -> str:
        """Get theme-specific stylesheet additions."""
        border = cls.get_border_color(is_dark)
        tooltip_border = cls.get_tooltip_border_color(is_dark)
        
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
        """


class StyleManager:
    """Main style management utility class."""
    
    @classmethod
    def setup_main_window(cls, window) -> None:
        """Apply all styling to the main window."""
        # Set window properties
        if hasattr(window, 'setWindowIcon'):
            window.setWindowIcon(AppStyles.create_app_icon())
        
        if hasattr(window, 'resize'):
            window.resize(*AppStyles.WINDOW_DEFAULT_SIZE)
        
        # Apply main stylesheet
        if hasattr(window, 'setStyleSheet'):
            window.setStyleSheet(StyleSheets.get_main_stylesheet())
    
    @classmethod
    def style_toolbar_button(cls, button) -> None:
        """Apply toolbar button styling."""
        ButtonStyles.apply_button_style(button, ButtonStyles.TOOLBAR_BUTTON)
    
    @classmethod
    def style_action_button(cls, button) -> None:
        """Apply action button styling."""
        ButtonStyles.apply_button_style(button, ButtonStyles.ACTION_BUTTON)
    
    @classmethod
    def style_browse_button(cls, button) -> None:
        """Apply browse button styling."""
        ButtonStyles.apply_button_style(button, ButtonStyles.BROWSE_BUTTON)
