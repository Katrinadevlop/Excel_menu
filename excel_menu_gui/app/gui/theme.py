import enum
from typing import Callable, Optional
import winreg

from PySide6.QtCore import QTimer
from PySide6.QtGui import QPalette, QColor
from PySide6.QtWidgets import QApplication
from ui_styles import ThemeAwareStyles


class ThemeMode(enum.Enum):
    SYSTEM = "system"
    LIGHT = "light"
    DARK = "dark"


def _read_reg_dword(root, path: str, name: str) -> Optional[int]:
    try:
        with winreg.OpenKey(root, path) as key:
            val, _ = winreg.QueryValueEx(key, name)
            if isinstance(val, int):
                return val
    except OSError:
        return None
    return None


def windows_apps_use_light_theme() -> bool:
    # HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme (1=light, 0=dark)
    val = _read_reg_dword(winreg.HKEY_CURRENT_USER,
                          r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
                          "AppsUseLightTheme")
    if val is None:
        return True
    return val != 0


def windows_accent_color() -> QColor:
    # Try DWM colorization color first, fallback to a blue accent
    val = _read_reg_dword(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\DWM", "ColorizationColor")
    if val is None:
        val = _read_reg_dword(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\DWM", "AccentColor")
    if val is None:
        return QColor("#3d6ff5")
    # Interpret as BGR (common for these DWORDs). Alpha may be in high byte.
    r = (val & 0x000000FF)
    g = (val & 0x0000FF00) >> 8
    b = (val & 0x00FF0000) >> 16
    # Clamp and return
    r = max(0, min(255, r)); g = max(0, min(255, g)); b = max(0, min(255, b))
    return QColor(r, g, b)


def build_palette(dark: bool) -> QPalette:
    accent = windows_accent_color()
    p = QPalette()
    if dark:
        window = QColor("#1e1f22")
        base = QColor("#2b2d30")
        alt = QColor("#242529")
        text = QColor("#e6e6e6")
        button = QColor("#2f3136")
        button_text = QColor("#e6e6e6")
        tooltip_base = base
        tooltip_text = text
        link = QColor(accent)
        highlight = QColor(accent)
        highlighted_text = QColor("#ffffff")
        disabled_text = QColor("#808080")
    else:
        window = QColor("#fafafa")
        base = QColor("#ffffff")
        alt = QColor("#f2f2f2")
        text = QColor("#202020")
        button = QColor("#f3f3f3")
        button_text = QColor("#202020")
        tooltip_base = QColor("#ffffdc")
        tooltip_text = QColor("#202020")
        link = QColor(accent)
        highlight = QColor(accent)
        highlighted_text = QColor("#ffffff")
        disabled_text = QColor("#9a9a9a")

    p.setColor(QPalette.Window, window)
    p.setColor(QPalette.WindowText, text)
    p.setColor(QPalette.Base, base)
    p.setColor(QPalette.AlternateBase, alt)
    p.setColor(QPalette.ToolTipBase, tooltip_base)
    p.setColor(QPalette.ToolTipText, tooltip_text)
    p.setColor(QPalette.Text, text)
    p.setColor(QPalette.Button, button)
    p.setColor(QPalette.ButtonText, button_text)
    p.setColor(QPalette.BrightText, QColor("#ff3333"))
    p.setColor(QPalette.Link, link)
    p.setColor(QPalette.Highlight, highlight)
    p.setColor(QPalette.HighlightedText, highlighted_text)

    # Disabled state tuning
    p.setColor(QPalette.Disabled, QPalette.Text, disabled_text)
    p.setColor(QPalette.Disabled, QPalette.ButtonText, disabled_text)
    p.setColor(QPalette.Disabled, QPalette.WindowText, disabled_text)

    return p


def apply_theme(app: QApplication, mode: ThemeMode) -> None:
    if mode == ThemeMode.SYSTEM:
        dark = not windows_apps_use_light_theme()
    elif mode == ThemeMode.DARK:
        dark = True
    else:
        dark = False

    # Сбрасываем старую палитру и стили перед применением новой темы
    app.setStyle("Fusion")
    p = build_palette(dark)
    app.setPalette(p)
    
    # Очищаем и заново устанавливаем stylesheet чтобы избежать смешения
    app.setStyleSheet("")

    # Apply theme-aware styling using centralized styles
    theme_stylesheet = ThemeAwareStyles.get_theme_stylesheet(dark)
    app.setStyleSheet(theme_stylesheet)


def start_system_theme_watcher(on_change: Callable[[bool], None], interval_ms: int = 1500) -> QTimer:
    """
    Периодически проверяет реестр Windows на смену темы (светлая/тёмная) и вызывает on_change(light: bool).
    Возвращает запущенный QTimer.
    """
    timer = QTimer()
    state = {"light": windows_apps_use_light_theme()}

    def tick():
        cur = windows_apps_use_light_theme()
        if cur != state["light"]:
            state["light"] = cur
            on_change(cur)

    timer.timeout.connect(tick)
    timer.start(interval_ms)
    return timer

