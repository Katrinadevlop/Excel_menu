import os
import sys
import shutil
import logging
import hashlib
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple

from PySide6.QtCore import Qt, QMimeData, QSize, QUrl, QSettings
from PySide6.QtGui import QPalette, QColor, QIcon, QPixmap, QPainter, QPen, QBrush, QLinearGradient, QFont, QDesktopServices
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QBoxLayout,
    QLabel, QPushButton, QFileDialog, QTextEdit, QComboBox, QLineEdit,
    QGroupBox, QCheckBox, QSpinBox, QRadioButton, QButtonGroup, QMessageBox, QFrame, QSizePolicy, QScrollArea,
    QListWidget, QListWidgetItem, QInputDialog,
)

from app.services.comparator import compare_and_highlight, get_sheet_names, ColumnParseError
# Временная алиас-совместимость для старых импортов внутри процесса
import importlib as _importlib
sys.modules.setdefault('comparator', _importlib.import_module('app.services.comparator'))
sys.modules.setdefault('brokerage_journal', _importlib.import_module('app.reports.brokerage_journal'))
from app.services.template_linker import default_template_path
from app.gui.theme import ThemeMode, apply_theme, start_system_theme_watcher
from app.reports.presentation_handler import create_presentation_with_excel_data
from app.reports.brokerage_journal import create_brokerage_journal_from_menu
from app.reports.pricelist_excel import create_pricelist_xlsx
from app.services.dish_extractor import extract_all_dishes_with_details, DishItem
from app.integrations.iiko_rms_client import IikoRmsClient, IikoApiError
from app.services.menu_template_filler import MenuTemplateFiller
from app.gui.ui_styles import (
    AppStyles, ButtonStyles, LayoutStyles, StyleSheets, ComponentStyles,
    StyleManager, ThemeAwareStyles
)


class DropLineEdit(QLineEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setPlaceholderText("Перетащите файл сюда или нажмите Обзор…")

    def dragEnterEvent(self, event):
        md: QMimeData = event.mimeData()
        if md.hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            local = urls[0].toLocalFile()
            if local:
                self.setText(local)
        event.acceptProposedAction()


def label_caption(text: str) -> QLabel:
    lbl = QLabel(text)
    ComponentStyles.style_caption_label(lbl)
    return lbl


def nice_group(title: str, content: QWidget) -> QGroupBox:
    gb = QGroupBox(title)
    lay = QVBoxLayout(gb)
    lay.addWidget(content)
    return gb


class FileDropGroup(QGroupBox):
    def __init__(self, title: str, target_line_edit: QLineEdit, content: QWidget, parent=None):
        super().__init__(title, parent)
        self._target = target_line_edit
        self.setAcceptDrops(True)
        lay = QVBoxLayout(self)
        lay.addWidget(content)

    def dragEnterEvent(self, event):
        md: QMimeData = event.mimeData()
        if md.hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            local = urls[0].toLocalFile()
            if local:
                self._target.setText(local)
        event.acceptProposedAction()


def create_app_icon() -> QIcon:
    """Legacy function for compatibility. Use AppStyles.create_app_icon() instead."""
    return AppStyles.create_app_icon()


def find_template(filename: str) -> Optional[str]:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
    candidates = [
        base / "excel_menu_gui" / "templates" / filename,
        base / "templates" / filename,
        Path(__file__).parent / "templates" / filename,
    ]
    for p in candidates:
        if p.exists():
            return str(p)
    return None


@dataclass
class FileConfig:
    path: str = ""
    sheet: str = ""
    col: str = "A"
    header_row_1based: int = 1


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # Локальные настройки (Windows Registry) для запоминания параметров iiko
        self._settings = QSettings("excel_menu_gui", "excel_menu_gui")
        self.setWindowTitle("Работа с меню")
        # Apply centralized styling
        StyleManager.setup_main_window(self)

        central = QWidget()
        self.setCentralWidget(central)
        self.rootLayout = QVBoxLayout(central)
        LayoutStyles.apply_margins(self.rootLayout, LayoutStyles.DEFAULT_MARGINS)
        self.rootLayout.setSpacing(AppStyles.DEFAULT_SPACING)

        # Панель управления (сверху)
        self.topBar = QFrame(); self.topBar.setObjectName("topBar")
        self.layTop = QHBoxLayout(self.topBar)
        LayoutStyles.apply_margins(self.layTop, LayoutStyles.TOPBAR_MARGINS)
        self.layTop.setSpacing(AppStyles.DEFAULT_SPACING)

        lblTheme = QLabel("Тема:")
        self.cmbTheme = QComboBox()
        self.cmbTheme.addItems(["Системная", "Светлая", "Тёмная"])
        self.cmbTheme.setCurrentIndex(0)
        self.cmbTheme.currentIndexChanged.connect(self.on_theme_changed)

        self.btnDownloadTemplate = QPushButton("Скачать шаблон меню")
        self.btnDownloadTemplate.clicked.connect(self.do_download_template)
        self.btnMakePresentation = QPushButton("Сделать презентацию")
        self.btnMakePresentation.clicked.connect(self.do_make_presentation)
        self.btnBrokerage = QPushButton("Бракеражный журнал")
        self.btnBrokerage.clicked.connect(self.do_brokerage_journal)
        self.btnOpenMenu = QPushButton("Открыть меню")
        self.btnOpenMenu.clicked.connect(self.do_open_menu)
        self.btnDownloadPricelists = QPushButton("Скачать ценники")
        self.btnDownloadPricelists.clicked.connect(self.do_download_pricelists)
        self.btnShowCompare = QPushButton("Сравнить меню")
        self.btnShowCompare.clicked.connect(self.show_compare_sections)

        self.layTop.addWidget(lblTheme)
        self.layTop.addWidget(self.cmbTheme)
        self.layTop.addStretch(1)
        self.layTop.addWidget(self.btnShowCompare)
        self.layTop.addWidget(self.btnDownloadTemplate)
        self.layTop.addWidget(self.btnMakePresentation)
        self.layTop.addWidget(self.btnBrokerage)
        self.layTop.addWidget(self.btnOpenMenu)
        self.layTop.addWidget(self.btnDownloadPricelists)

        # Apply centralized stylesheet (already set in StyleManager.setup_main_window)

        self.topGroup = nice_group("Панель управления", self.topBar)
        self.topGroup.setObjectName("topGroup")
        LayoutStyles.apply_size_policy(self.topGroup, LayoutStyles.EXPANDING_FIXED)
        self.rootLayout.addWidget(self.topGroup)

        # Область прокрутки для остального содержимого, чтобы панель всегда была наверху
        self.scrollArea = QScrollArea()
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setFrameShape(QFrame.NoFrame)
        self.contentContainer = QWidget()
        self.contentLayout = QVBoxLayout(self.contentContainer)
        LayoutStyles.apply_margins(self.contentLayout, LayoutStyles.NO_MARGINS)
        self.contentLayout.setSpacing(AppStyles.CONTENT_SPACING)  # фиксированный вертикальный интервал между компонентами
        self.scrollArea.setWidget(self.contentContainer)
        self.scrollArea.setAlignment(Qt.AlignTop)  # прижимаем контент к верху, если он ниже области
        self.rootLayout.addWidget(self.scrollArea, 1)

        # Excel файл для презентации (используем тот же стиль, что и для файлов сравнения)
        self.edExcelPath = DropLineEdit()
        self.edExcelPath.setPlaceholderText("Выберите Excel файл с меню...")
        self.btnBrowseExcel = QPushButton("Обзор…")
        self.btnBrowseExcel.clicked.connect(lambda: self.browse_excel_file())

        excel_row = QWidget(); excel_layout = QHBoxLayout(excel_row)
        excel_layout.addWidget(self.edExcelPath, 1)
        excel_layout.addWidget(self.btnBrowseExcel)

        excel_group = QWidget(); excel_group_layout = QVBoxLayout(excel_group)
        excel_group_layout.addWidget(label_caption("Выберите Excel файл с меню..."))
        excel_group_layout.addWidget(excel_row)
        
        self.grpExcelFile = FileDropGroup("Выберите Excel файл с меню для презентации...", self.edExcelPath, excel_group)
        # Apply centralized styling for Excel file groups
        ComponentStyles.style_excel_group(self.grpExcelFile)
        self.contentLayout.addWidget(self.grpExcelFile)
        self.grpExcelFile.setVisible(False)
        
        # Панель действий внизу для сравнения (фиксированная)
        self.actionsPanel = QWidget(); self.actionsPanel.setObjectName("actionsPanel")
        self.actionsLayout = QHBoxLayout(self.actionsPanel)
        LayoutStyles.apply_margins(self.actionsLayout, LayoutStyles.MINIMAL_TOP_MARGIN)  # минимальный отступ сверху
        self.btnCompare = QPushButton("Сравнить и подсветить")
        self.btnCompare.clicked.connect(self.do_compare)
        self.actionsLayout.addStretch(1)
        self.actionsLayout.addWidget(self.btnCompare)
        self.rootLayout.addWidget(self.actionsPanel)
        self.actionsPanel.setVisible(False)
        
        # Панель действий внизу для презентаций (фиксированная)
        self.presentationActionsPanel = QWidget(); self.presentationActionsPanel.setObjectName("actionsPanel")
        self.presentationActionsLayout = QHBoxLayout(self.presentationActionsPanel)
        LayoutStyles.apply_margins(self.presentationActionsLayout, LayoutStyles.CONTENT_TOP_MARGIN)  # небольшой отступ сверху
        self.btnCreatePresentationWithData = QPushButton("Скачать презентацию с меню")
        self.btnCreatePresentationWithData.clicked.connect(self.do_create_presentation_with_data)
        self.presentationActionsLayout.addStretch(1)
        self.presentationActionsLayout.addWidget(self.btnCreatePresentationWithData)
        self.rootLayout.addWidget(self.presentationActionsPanel)
        self.presentationActionsPanel.setVisible(False)
        
        # Панель действий внизу для бракеражного журнала (фиксированная)
        self.brokerageActionsPanel = QWidget(); self.brokerageActionsPanel.setObjectName("actionsPanel")
        self.brokerageActionsLayout = QHBoxLayout(self.brokerageActionsPanel)
        LayoutStyles.apply_margins(self.brokerageActionsLayout, LayoutStyles.CONTENT_TOP_MARGIN)  # небольшой отступ сверху
        self.btnCreateBrokerageJournal = QPushButton("Скачать бракеражный журнал")
        self.btnCreateBrokerageJournal.clicked.connect(self.do_create_brokerage_journal_with_data)
        self.brokerageActionsLayout.addStretch(1)
        self.brokerageActionsLayout.addWidget(self.btnCreateBrokerageJournal)
        self.rootLayout.addWidget(self.brokerageActionsPanel)
        self.brokerageActionsPanel.setVisible(False)
        
        # Панель действий внизу для шаблона меню (фиксированная)
        self.templateActionsPanel = QWidget(); self.templateActionsPanel.setObjectName("actionsPanel")
        self.templateActionsLayout = QHBoxLayout(self.templateActionsPanel)
        LayoutStyles.apply_margins(self.templateActionsLayout, LayoutStyles.CONTENT_TOP_MARGIN)  # небольшой отступ сверху
        self.btnFillTemplate = QPushButton("Заполнить и скачать шаблон меню")
        self.btnFillTemplate.clicked.connect(self.do_fill_template_with_data)
        self.templateActionsLayout.addStretch(1)
        self.templateActionsLayout.addWidget(self.btnFillTemplate)
        self.rootLayout.addWidget(self.templateActionsPanel)
        self.templateActionsPanel.setVisible(False)

        # File 1
        self.edPath1 = DropLineEdit()
        self.btnBrowse1 = QPushButton("Обзор…")
        self.btnBrowse1.clicked.connect(lambda: self.browse_file(self.edPath1, which=1))

        row1 = QWidget(); r1 = QHBoxLayout(row1)
        r1.addWidget(self.edPath1, 1)
        r1.addWidget(self.btnBrowse1)

        g1 = QWidget(); l1 = QVBoxLayout(g1)
        l1.addWidget(label_caption("Файл 1"))
        l1.addWidget(row1)
        self.grpFirst = FileDropGroup("Первый файл", self.edPath1, g1)
        # Apply centralized styling for file groups
        ComponentStyles.style_file_group(self.grpFirst)
        self.contentLayout.addWidget(self.grpFirst)
        self.grpFirst.setVisible(False)

        # File 2
        self.edPath2 = DropLineEdit()
        self.btnBrowse2 = QPushButton("Обзор…")
        self.btnBrowse2.clicked.connect(lambda: self.browse_file(self.edPath2, which=2))

        row2 = QWidget(); r2 = QHBoxLayout(row2)
        r2.addWidget(self.edPath2, 1)
        r2.addWidget(self.btnBrowse2)

        g2 = QWidget(); l2 = QVBoxLayout(g2)
        l2.addWidget(label_caption("Файл 2"))
        l2.addWidget(row2)
        self.grpSecond = FileDropGroup("Второй файл", self.edPath2, g2)
        # Apply centralized styling for file groups
        ComponentStyles.style_file_group(self.grpSecond)
        self.contentLayout.addWidget(self.grpSecond)
        self.grpSecond.setVisible(False)

        # (дополнительно) — сворачиваемая группа
        opts = QWidget(); lo = QHBoxLayout(opts)
        LayoutStyles.apply_margins(lo, LayoutStyles.NO_MARGINS)
        lo.setSpacing(AppStyles.CONTENT_SPACING)  # немного больше для удобства чтения
        self.chkIgnoreCase = QCheckBox("Игнорировать регистр")
        self.chkIgnoreCase.setChecked(True)
        self.chkFuzzy = QCheckBox("Похожесть")
        self.spinFuzzy = QSpinBox(); self.spinFuzzy.setRange(0, 100); self.spinFuzzy.setValue(85)
        self.spinFuzzy.setEnabled(False)
        self.chkFuzzy.toggled.connect(self.spinFuzzy.setEnabled)

        lo.addWidget(self.chkIgnoreCase)
        lo.addWidget(self.chkFuzzy)
        lo.addWidget(QLabel("Порог:"))
        lo.addWidget(self.spinFuzzy)
        lo.addStretch(1)

        self.paramsBox = QGroupBox("Параметры (дополнительно)")
        # Apply centralized styling for parameter groups
        ComponentStyles.style_params_group(self.paramsBox)
        lparams = QVBoxLayout(self.paramsBox)
        lparams.setContentsMargins(0, 10, 0, 0)  # полностью убираем отступы
        lparams.setSpacing(AppStyles.CONTENT_SPACING)  # добавляем промежуток между заголовком и содержимым
        # Контейнер параметров без дополнительной рамки
        self._paramsFrame = QFrame(); self._paramsFrame.setObjectName("paramsFrame")
        lf = QHBoxLayout(self._paramsFrame)
        LayoutStyles.apply_margins(lf, LayoutStyles.NO_MARGINS)  # убираем отступы
        lf.setSpacing(0)
        lf.addWidget(opts)
        lparams.addWidget(self._paramsFrame)
        self._paramsFrame.setVisible(False)
        self.paramsBox.toggled.connect(self.on_params_toggled)
        self.contentLayout.addWidget(self.paramsBox)
        self.paramsBox.setVisible(False)

        # ===== ЦЕННИКИ: поиск + выбор =====
        self._pricelist_dishes: List[DishItem] = []
        self._pricelist_selected_keys: set[str] = set()

        # Источник блюд: только iiko (без экселя и без UI-настроек подключения)
        self._iiko_base_url = str(self._settings.value("iiko/base_url", "https://287-772-687.iiko.it/resto"))
        self._iiko_login = str(self._settings.value("iiko/login", "user"))

        # Храним только sha1-хэш пароля (как требует iikoRMS resto API).
        self._iiko_pass_sha1_cached = str(self._settings.value("iiko/pass_sha1", ""))

        # Миграция со старой версии: если вдруг сохранён plaintext-пароль — конвертируем и удаляем.
        try:
            legacy_pwd = str(self._settings.value("iiko/password", ""))
            if (not self._iiko_pass_sha1_cached) and legacy_pwd:
                self._iiko_pass_sha1_cached = hashlib.sha1(legacy_pwd.encode("utf-8")).hexdigest()
                self._settings.setValue("iiko/pass_sha1", self._iiko_pass_sha1_cached)
                self._settings.remove("iiko/password")
        except Exception:
            pass

        src_row = QWidget(); src_layout = QHBoxLayout(src_row)
        LayoutStyles.apply_margins(src_layout, LayoutStyles.NO_MARGINS)
        src_layout.addWidget(QLabel("Источник: iiko"))
        src_layout.addStretch(1)

        self.edDishSearch = QLineEdit()
        self.edDishSearch.setPlaceholderText("Начните вводить название блюда… (Enter — добавить)")
        self.edDishSearch.textChanged.connect(self._update_pricelist_suggestions)
        self.edDishSearch.returnPressed.connect(self._add_pricelist_from_enter)

        self.lblPricelistInfo = QLabel("1) Нажмите 'Загрузить блюда'  2) Введите название")

        self.btnLoadDishes = QPushButton("Загрузить блюда")
        self.btnLoadDishes.clicked.connect(self._load_pricelist_dishes)

        self.btnShowAllDishes = QPushButton("Показать все блюда")
        self.btnShowAllDishes.clicked.connect(self._show_all_pricelist_dishes)

        btns_row = QWidget(); btns_layout = QHBoxLayout(btns_row)
        LayoutStyles.apply_margins(btns_layout, LayoutStyles.NO_MARGINS)
        btns_layout.addWidget(self.btnLoadDishes)
        btns_layout.addWidget(self.btnShowAllDishes)
        btns_layout.addStretch(1)

        self.lstDishSuggestions = QListWidget()
        self.lstDishSuggestions.setMinimumHeight(220)
        self.lstDishSuggestions.itemClicked.connect(self._on_pricelist_suggestion_clicked)
        self.lstDishSuggestions.itemDoubleClicked.connect(self._on_pricelist_suggestion_clicked)

        self.lstSelectedDishes = QListWidget()
        self.lstSelectedDishes.setMinimumHeight(160)

        self.btnClearSelectedDishes = QPushButton("Очистить выбор")
        self.btnClearSelectedDishes.clicked.connect(self._clear_pricelist_selection)

        pricelist_box = QWidget(); pricelist_layout = QVBoxLayout(pricelist_box)
        pricelist_layout.addWidget(src_row)
        pricelist_layout.addWidget(label_caption("Поиск блюда"))
        pricelist_layout.addWidget(self.edDishSearch)
        pricelist_layout.addWidget(self.lblPricelistInfo)
        pricelist_layout.addWidget(btns_row)
        pricelist_layout.addWidget(label_caption("Подсказки (кликните, чтобы добавить)"))
        pricelist_layout.addWidget(self.lstDishSuggestions)
        pricelist_layout.addWidget(label_caption("Выбранные блюда (с галочками)"))
        pricelist_layout.addWidget(self.lstSelectedDishes)
        pricelist_layout.addWidget(self.btnClearSelectedDishes)

        self.grpPricelist = nice_group("Ценники: выбрать блюда", pricelist_box)
        self.contentLayout.addWidget(self.grpPricelist)
        self.grpPricelist.setVisible(False)


        # Панель действий внизу для ценников (фиксированная)
        self.pricelistActionsPanel = QWidget(); self.pricelistActionsPanel.setObjectName("actionsPanel")
        self.pricelistActionsLayout = QHBoxLayout(self.pricelistActionsPanel)
        LayoutStyles.apply_margins(self.pricelistActionsLayout, LayoutStyles.CONTENT_TOP_MARGIN)
        self.btnCreatePricelist = QPushButton("Сформировать ценники (Excel)")
        self.btnCreatePricelist.clicked.connect(self.do_create_pricelist_excel)
        self.pricelistActionsLayout.addStretch(1)
        self.pricelistActionsLayout.addWidget(self.btnCreatePricelist)
        self.rootLayout.addWidget(self.pricelistActionsPanel)
        self.pricelistActionsPanel.setVisible(False)

        # Добавляем нижний растягивающий элемент, чтобы контент не растягивался равномерно, а был прижат кверху
        self.contentLayout.addStretch(1)

        # Theming initialization
        self._theme_mode = ThemeMode.SYSTEM  # По умолчанию используем системную тему
        apply_theme(QApplication.instance(), self._theme_mode)
        
        # Запускаем таймер для отслеживания изменений системной темы
        self._theme_timer = start_system_theme_watcher(
            lambda is_light: self.handle_system_theme_change(is_light),
            interval_ms=1000  # Проверка каждую секунду
        )

        # Применяем компактный режим при узкой ширине окна (мобильный превью)
        try:
            self._apply_compact_mode(self.width() <= 480)
        except Exception:
            pass

    def log(self, msg: str):
        # Лог отключён по запросу — ничего не делаем
        pass

    def on_theme_changed(self, idx: int):
        # Получаем режим темы из выбора пользователя
        if idx == 0:
            self._theme_mode = ThemeMode.SYSTEM  # Системная тема
        elif idx == 1:
            self._theme_mode = ThemeMode.LIGHT   # Светлая тема
        else:
            self._theme_mode = ThemeMode.DARK    # Тёмная тема
            
        # Применяем выбранную тему
        apply_theme(QApplication.instance(), self._theme_mode)
        
    def handle_system_theme_change(self, is_light: bool):
        """Обработчик изменения системной темы Windows"""
        # Обновляем тему только если выбрана "Системная"
        if self._theme_mode == ThemeMode.SYSTEM and self.cmbTheme.currentIndex() == 0:
            # Применяем соответствующую системную тему
            theme = ThemeMode.LIGHT if is_light else ThemeMode.DARK
            apply_theme(QApplication.instance(), theme)


    def closeEvent(self, event):
        try:
            if hasattr(self, "_theme_timer") and self._theme_timer is not None:
                self._theme_timer.stop()
        except Exception:
            pass
        super().closeEvent(event)

    def _apply_compact_mode(self, compact: bool) -> None:
        if getattr(self, "_compact", None) == compact:
            return
        self._compact = compact
        try:
            # Корневой layout
            if hasattr(self, "rootLayout"):
                LayoutStyles.apply_margins(self.rootLayout, (8, 8, 8, 8) if compact else LayoutStyles.DEFAULT_MARGINS)
                self.rootLayout.setSpacing(AppStyles.COMPACT_SPACING if compact else AppStyles.DEFAULT_SPACING)
            # Верхняя панель: горизонтально на десктопе, вертикально в компактном режиме
            if hasattr(self, "layTop"):
                self.layTop.setDirection(QBoxLayout.TopToBottom if compact else QBoxLayout.LeftToRight)
                LayoutStyles.apply_margins(self.layTop, (8, 8, 8, 8) if compact else LayoutStyles.TOPBAR_MARGINS)
                self.layTop.setSpacing(AppStyles.COMPACT_SPACING if compact else AppStyles.DEFAULT_SPACING)
            # Контентные отступы
            if hasattr(self, "contentLayout"):
                self.contentLayout.setSpacing(AppStyles.COMPACT_SPACING if compact else AppStyles.CONTENT_SPACING)
        except Exception:
            pass

    def resizeEvent(self, event):
        super().resizeEvent(event)
        try:
            self._apply_compact_mode(self.width() <= 480)
        except Exception:
            pass

    def browse_file(self, target: QLineEdit, which: int):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите файл", str(Path.cwd()), "Excel (*.xls *.xlsx *.xlsm);;Все файлы (*.*)")
        if path:
            target.setText(path)

    def select_default_sheet(self, path: str) -> Optional[str]:
        try:
            names = get_sheet_names(path)
            if not names:
                return None
            for nm in names:
                low = nm.strip().lower()
                if "касс" in low or "kass" in low:
                    return nm
            return names[0]
        except Exception:
            return None


    def do_compare(self):
        try:
            p1 = self.edPath1.text().strip()
            p2 = self.edPath2.text().strip()
            s1 = self.select_default_sheet(p1) if p1 else None
            s2 = self.select_default_sheet(p2) if p2 else None
            if not (p1 and p2):
                QMessageBox.warning(self, "Внимание", "Укажите оба файла.")
                return
            if not (s1 and s2):
                QMessageBox.warning(self, "Листы", "Не удалось определить листы для сравнения.")
                return
            
            # Выбираем место сохранения результата сравнения
            # Получаем дату из Excel файлов для правильного названия
            from comparator import _extract_best_date_from_file
            from datetime import date
            
            # Извлекаем даты из обоих файлов
            d1 = _extract_best_date_from_file(p1, s1)
            d2 = _extract_best_date_from_file(p2, s2)
            
            # Определяем самую позднюю дату для названия файла
            latest_date = None
            if d1 and d2:
                latest_date = max(d1, d2)
            elif d1:
                latest_date = d1
            elif d2:
                latest_date = d2
            
            # Формируем предлагаемое имя с правильной датой
            if latest_date:
                date_str = latest_date.strftime("%d.%m.%Y")
                suggested_name = f"сравнение_меню_{date_str}.xlsx"
            else:
                # Если даты не найдены, используем текущую дату как fallback
                today_str = date.today().strftime("%d.%m.%Y")
                suggested_name = f"сравнение_меню_{today_str}.xlsx"
            
            desktop = Path.home() / "Desktop"
            suggested_path = str(desktop / suggested_name)
            
            save_path, _ = QFileDialog.getSaveFileName(
                self, 
                "Сохранить результат сравнения", 
                suggested_path, 
                "Excel (*.xlsx);;Excel (*.xls);;Все файлы (*.*)"
            )
            
            if not save_path:
                return  # Пользователь отменил сохранение
            
            try:
                # Всегда авто по дате
                temp_out_path, matches = compare_and_highlight(
                    path1=p1, sheet1=s1,
                    path2=p2, sheet2=s2,
                    col1="A",
                    col2="E",
                    header_row1=1,
                    header_row2=1,
                    ignore_case=self.chkIgnoreCase.isChecked(),
                    use_fuzzy=self.chkFuzzy.isChecked(),
                    fuzzy_threshold=int(self.spinFuzzy.value()),
                    final_choice=0,  # авто: финальным будет файл с более поздней датой
                )
                
                # Перемещаем файл в выбранное место
                if temp_out_path != save_path:
                    import shutil
                    shutil.move(temp_out_path, save_path)
                
                self.log(f"Готово. Совпадений: {matches}. Итоговый файл: {save_path}")
                QMessageBox.information(self, "Готово", f"Совпадений: {matches}\nФайл сохранён: {save_path}")
            except ColumnParseError as e:
                QMessageBox.warning(self, "Колонка", str(e))
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def show_compare_sections(self):
        try:
            # Скрываем панели, не относящиеся к сравнению
            if hasattr(self, "grpExcelFile"):
                self.grpExcelFile.setVisible(False)
            if hasattr(self, "presentationActionsPanel"):
                self.presentationActionsPanel.setVisible(False)
            if hasattr(self, "brokerageActionsPanel"):
                self.brokerageActionsPanel.setVisible(False)
            if hasattr(self, "templateActionsPanel"):
                self.templateActionsPanel.setVisible(False)
            if hasattr(self, "grpPricelist"):
                self.grpPricelist.setVisible(False)
            if hasattr(self, "pricelistActionsPanel"):
                self.pricelistActionsPanel.setVisible(False)

            # Показать формы сравнения и панель действий
            if hasattr(self, "grpFirst"):
                self.grpFirst.setVisible(True)
            if hasattr(self, "grpSecond"):
                self.grpSecond.setVisible(True)
            if hasattr(self, "paramsBox"):
                self.paramsBox.setVisible(True)
                self.paramsBox.setChecked(False)
            if hasattr(self, "actionsPanel"):
                self.actionsPanel.setVisible(True)
        except Exception:
            pass

    def on_params_toggled(self, checked: bool):
        try:
            if hasattr(self, "_paramsFrame"):
                self._paramsFrame.setVisible(checked)
            if checked and hasattr(self, "scrollArea") and hasattr(self, "_paramsFrame"):
                self.scrollArea.ensureWidgetVisible(self._paramsFrame)
        except Exception:
            pass

    def do_download_template(self):
        """Показывает панель для работы с шаблоном меню"""
        try:
            # Скрываем другие панели
            if hasattr(self, "grpFirst"):
                self.grpFirst.setVisible(False)
            if hasattr(self, "grpSecond"):
                self.grpSecond.setVisible(False)
            if hasattr(self, "paramsBox"):
                self.paramsBox.setVisible(False)
            if hasattr(self, "actionsPanel"):
                self.actionsPanel.setVisible(False)
            if hasattr(self, "presentationActionsPanel"):
                self.presentationActionsPanel.setVisible(False)
            if hasattr(self, "brokerageActionsPanel"):
                self.brokerageActionsPanel.setVisible(False)
            if hasattr(self, "grpPricelist"):
                self.grpPricelist.setVisible(False)
            if hasattr(self, "pricelistActionsPanel"):
                self.pricelistActionsPanel.setVisible(False)
            
            # Показываем панель для работы с шаблоном меню и её панель действий
            if hasattr(self, "grpExcelFile"):
                self.grpExcelFile.setVisible(True)
                self.grpExcelFile.setTitle("Файл меню для заполнения шаблона")
                self.edExcelPath.setPlaceholderText("Выберите Excel файл с меню для заполнения шаблона...")
            if hasattr(self, "templateActionsPanel"):
                self.templateActionsPanel.setVisible(True)
                
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def do_make_presentation(self):
        """Показывает панель для работы с презентациями"""
        try:
            # Скрываем другие панели
            if hasattr(self, "grpFirst"):
                self.grpFirst.setVisible(False)
            if hasattr(self, "grpSecond"):
                self.grpSecond.setVisible(False)
            if hasattr(self, "paramsBox"):
                self.paramsBox.setVisible(False)
            if hasattr(self, "actionsPanel"):
                self.actionsPanel.setVisible(False)
            if hasattr(self, "brokerageActionsPanel"):
                self.brokerageActionsPanel.setVisible(False)
            if hasattr(self, "templateActionsPanel"):
                self.templateActionsPanel.setVisible(False)
            if hasattr(self, "grpPricelist"):
                self.grpPricelist.setVisible(False)
            if hasattr(self, "pricelistActionsPanel"):
                self.pricelistActionsPanel.setVisible(False)
            
            # Показываем панель для работы с презентациями и её панель действий
            if hasattr(self, "grpExcelFile"):
                self.grpExcelFile.setVisible(True)
                self.grpExcelFile.setTitle("Файл меню для презентации")
                self.edExcelPath.setPlaceholderText("Выберите Excel файл с меню (салаты, первые блюда, мясо, птица, рыба, гарниры)...")
            if hasattr(self, "presentationActionsPanel"):
                self.presentationActionsPanel.setVisible(True)
                
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))
    
    def browse_excel_file(self):
        """Выбор Excel файла для презентации"""
        path, _ = QFileDialog.getOpenFileName(
            self, 
            "Выберите Excel файл с меню", 
            str(Path.cwd()), 
            "Excel (*.xls *.xlsx *.xlsm);;Все файлы (*.*)"
        )
        if path:
            self.edExcelPath.setText(path)

    def do_open_menu(self):
        """Открывает файл меню (Excel) в приложении по умолчанию (например, Excel)."""
        try:
            path, _ = QFileDialog.getOpenFileName(
                self,
                "Открыть меню",
                str(Path.cwd()),
                "Excel (*.xls *.xlsx *.xlsm);;Все файлы (*.*)",
            )
            if not path:
                return

            ok = QDesktopServices.openUrl(QUrl.fromLocalFile(path))
            if not ok:
                QMessageBox.warning(self, "Не удалось открыть", "Не удалось открыть файл выбранной программой.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def do_download_pricelists(self):
        """Показывает раздел для формирования ценников по выбранным блюдам."""
        try:
            # Скрываем другие панели
            if hasattr(self, "grpFirst"):
                self.grpFirst.setVisible(False)
            if hasattr(self, "grpSecond"):
                self.grpSecond.setVisible(False)
            if hasattr(self, "paramsBox"):
                self.paramsBox.setVisible(False)
            if hasattr(self, "actionsPanel"):
                self.actionsPanel.setVisible(False)
            if hasattr(self, "presentationActionsPanel"):
                self.presentationActionsPanel.setVisible(False)
            if hasattr(self, "brokerageActionsPanel"):
                self.brokerageActionsPanel.setVisible(False)
            if hasattr(self, "templateActionsPanel"):
                self.templateActionsPanel.setVisible(False)

            # Для ценников НЕ показываем общий блок выбора Excel-файла (он мешает и не нужен при iiko)
            if hasattr(self, "grpExcelFile"):
                self.grpExcelFile.setVisible(False)

            # Показываем панель ценников
            if hasattr(self, "grpPricelist"):
                self.grpPricelist.setVisible(True)
            if hasattr(self, "pricelistActionsPanel"):
                self.pricelistActionsPanel.setVisible(True)

            # Сброс состояния
            self._pricelist_dishes = []
            self._pricelist_selected_keys = set()
            self.lblPricelistInfo.setText("1) Нажмите 'Загрузить блюда'  2) Начните ввод")
            self.edDishSearch.clear()
            self.lstDishSuggestions.clear()
            self.lstSelectedDishes.clear()


            try:
                self.edDishSearch.setFocus()
            except Exception:
                pass

            # Автозагрузка блюд из iiko: только если пароль уже сохранён.
            # Иначе пользователь нажмёт кнопку «Загрузить блюда» и введёт пароль один раз.
            try:
                if (not self._pricelist_dishes) and bool(self._iiko_pass_sha1_cached):
                    self._load_pricelist_dishes()
            except Exception:
                pass

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))
    
    def do_create_presentation_with_data(self):
        """Создает презентацию с данными из Excel файла"""
        try:
            # Получаем путь к Excel файлу
            excel_path = self.edExcelPath.text().strip()
            if not excel_path:
                QMessageBox.warning(self, "Внимание", "Выберите Excel файл с меню.")
                return
            
            # Проверяем существование Excel файла
            if not Path(excel_path).exists():
                QMessageBox.warning(self, "Ошибка", "Указанный Excel файл не найден.")
                return
                
            # Находим шаблон презентации
            template_path = find_template("presentation_template.pptx")
            if not template_path:
                QMessageBox.warning(self, "Шаблон", "Шаблон презентации не найден. Положите файл presentation_template.pptx в папку templates.")
                return
            
            # Выбираем место сохранения презентации
            suggested_name = "menu — копия.pptx"
            desktop = Path.home() / "Desktop"
            suggested_path = str(desktop / suggested_name)
            
            save_path, _ = QFileDialog.getSaveFileName(
                self, 
                "Сохранить презентацию с меню", 
                suggested_path, 
                "PowerPoint (*.pptx);;PowerPoint (*.ppt);;Все файлы (*.*)"
            )
            
            if not save_path:
                return  # Пользователь отменил сохранение
                
            # Создаем презентацию с данными
            success, message = create_presentation_with_excel_data(
                template_path, 
                excel_path, 
                save_path
            )
            
            if success:
                # Убрано информационное окно при успешном сохранении
                pass
            else:
                QMessageBox.warning(self, "Ошибка", f"Не удалось создать презентацию:\n{message}")
                
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

    def do_brokerage_journal(self):
        """Показывает панель для работы с бракеражным журналом"""
        try:
            # Скрываем другие панели
            if hasattr(self, "grpFirst"):
                self.grpFirst.setVisible(False)
            if hasattr(self, "grpSecond"):
                self.grpSecond.setVisible(False)
            if hasattr(self, "paramsBox"):
                self.paramsBox.setVisible(False)
            if hasattr(self, "actionsPanel"):
                self.actionsPanel.setVisible(False)
            if hasattr(self, "presentationActionsPanel"):
                self.presentationActionsPanel.setVisible(False)
            if hasattr(self, "templateActionsPanel"):
                self.templateActionsPanel.setVisible(False)
            if hasattr(self, "grpPricelist"):
                self.grpPricelist.setVisible(False)
            if hasattr(self, "pricelistActionsPanel"):
                self.pricelistActionsPanel.setVisible(False)
            
            # Показываем панель для работы с бракеражным журналом и её панель действий
            if hasattr(self, "grpExcelFile"):
                self.grpExcelFile.setVisible(True)
                self.grpExcelFile.setTitle("Файл меню для бракеражного журнала")
                self.edExcelPath.setPlaceholderText("Выберите Excel файл с меню для бракеражного журнала...")
            if hasattr(self, "brokerageActionsPanel"):
                self.brokerageActionsPanel.setVisible(True)
                
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))
    
    def do_create_brokerage_journal_with_data(self):
        """Создает бракеражный журнал с данными из Excel файла"""
        try:
            # Получаем путь к Excel файлу с меню
            excel_path = self.edExcelPath.text().strip()
            if not excel_path:
                QMessageBox.warning(self, "Внимание", "Выберите Excel файл с меню.")
                return
            
            # Проверяем существование Excel файла
            if not Path(excel_path).exists():
                QMessageBox.warning(self, "Ошибка", "Указанный Excel файл не найден.")
                return
                
            # Находим шаблон бракеражного журнала
            template_path = find_template("Бракеражный журнал шаблон.xlsx")
            if not template_path:
                QMessageBox.warning(self, "Шаблон", "Шаблон бракеражного журнала не найден. Положите файл 'Бракеражный журнал шаблон.xlsx' в папку templates.")
                return
            
            # Получаем дату для названия файла
            from app.reports.brokerage_journal import BrokerageJournalGenerator
            from datetime import date
            
            generator = BrokerageJournalGenerator()
            menu_date = generator.extract_date_from_menu(excel_path)
            
            if menu_date:
                date_str = menu_date.strftime("%d.%m.%Y")
                suggested_name = f"бракеражный_журнал_{date_str}.xlsx"
            else:
                today_str = date.today().strftime("%d.%m.%Y")
                suggested_name = f"бракеражный_журнал_{today_str}.xlsx"
            
            # Выбираем место сохранения бракеражного журнала
            desktop = Path.home() / "Desktop"
            suggested_path = str(desktop / suggested_name)
            
            save_path, _ = QFileDialog.getSaveFileName(
                self,
                "Сохранить бракеражный журнал",
                suggested_path,
                "Excel (*.xlsx);;Excel (*.xls);;Все файлы (*.*)"
            )
            
            if not save_path:
                return  # Пользователь отменил сохранение
                
            # Создаем бракеражный журнал с данными
            success, message = create_brokerage_journal_from_menu(
                excel_path, 
                template_path, 
                save_path
            )
            
            if success:
                # Убрано информационное окно при успешном сохранении
                pass
            else:
                QMessageBox.warning(self, "Ошибка", f"Не удалось создать бракеражный журнал:\n{message}")
                
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

    def do_fill_template_with_data(self):
        """Заполняет шаблон меню данными из Excel файла"""
        try:
            # Получаем путь к Excel файлу с меню
            excel_path = self.edExcelPath.text().strip()
            if not excel_path:
                QMessageBox.warning(self, "Внимание", "Выберите Excel файл с меню.")
                return
            
            # Проверяем существование Excel файла
            if not Path(excel_path).exists():
                QMessageBox.warning(self, "Ошибка", "Указанный Excel файл не найден.")
                return
                
            # Находим шаблон меню
            template_path = default_template_path()
            if not template_path or not Path(template_path).exists():
                QMessageBox.warning(self, "Шаблон", "Шаблон меню не найден. Положите файл шаблона в папку templates/.")
                return
            
            # Получаем дату для названия файла
            from datetime import date
            
            # Создаём экземпляр класса заполнения шаблона
            filler = MenuTemplateFiller()
            
            # Получаем дату из меню
            menu_date = filler.extract_date_from_menu(excel_path)
            
            # Формируем имя "<день> <месяц> - <день недели>.xlsx"
            russian_months = {
                1: 'января', 2: 'февраля', 3: 'марта', 4: 'апреля', 5: 'мая', 6: 'июня',
                7: 'июля', 8: 'августа', 9: 'сентября', 10: 'октября', 11: 'ноября', 12: 'декабря'
            }
            weekday_names = {
                0: 'понедельник', 1: 'вторник', 2: 'среда', 3: 'четверг', 4: 'пятница', 5: 'суббота', 6: 'воскресенье'
            }
            if menu_date:
                d, m, wd = menu_date.day, menu_date.month, menu_date.weekday()
            else:
                today = date.today()
                d, m, wd = today.day, today.month, today.weekday()
            suggested_name = f"{d} {russian_months.get(m, '')} - {weekday_names.get(wd, '')}.xlsx"
            
            # Выбираем место сохранения заполненного шаблона
            desktop = Path.home() / "Desktop"
            suggested_path = str(desktop / suggested_name)
            
            save_path, _ = QFileDialog.getSaveFileName(
                self,
                "Сохранить заполненный шаблон меню",
                suggested_path,
                "Excel (*.xlsx);;Excel (*.xls);;Все файлы (*.*)"
            )
            
            if not save_path:
                return  # Пользователь отменил сохранение
                
            # Копирование прямоугольника A6..F42 с листа «Касса» источника в лист «Касса» шаблона
            success, message = filler.copy_kassa_rect_A6_F42(
                template_path=template_path,
                source_menu_path=excel_path,
                output_path=save_path,
            )
            
            if not success:
                QMessageBox.warning(self, "Ошибка", f"Не удалось выполнить операцию:\n{message}")
                
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

    # ===== ЦЕННИКИ: логика =====
    def _ensure_iiko_pass_sha1(self) -> bool:
        """Если sha1(pass) не сохранён — спрашиваем пароль один раз и сохраняем только sha1."""
        if self._iiko_pass_sha1_cached:
            return True

        pwd, ok = QInputDialog.getText(
            self,
            "iiko",
            "Введите пароль iiko (сохранится только SHA1-хэш на этом компьютере):",
            QLineEdit.Password,
        )
        if not ok:
            return False

        pwd = (pwd or "").strip()
        if not pwd:
            return False

        self._iiko_pass_sha1_cached = hashlib.sha1(pwd.encode("utf-8")).hexdigest()
        try:
            self._settings.setValue("iiko/pass_sha1", self._iiko_pass_sha1_cached)
            # На всякий случай удалим возможный legacy plaintext
            self._settings.remove("iiko/password")
        except Exception:
            pass

        return True

    def _pl_key(self, name: str) -> str:
        return " ".join(str(name).strip().lower().replace('ё', 'е').split())

    def _format_dish_line(self, d: DishItem) -> str:
        parts = [d.name]
        if d.weight:
            parts.append(str(d.weight))
        if d.price:
            parts.append(str(d.price))
        return " — ".join(parts)

    def _load_pricelist_dishes(self):
        try:
            if not self._ensure_iiko_pass_sha1():
                return

            base_url = (self._iiko_base_url or "").strip()
            login = (self._iiko_login or "").strip()
            pass_sha1 = (self._iiko_pass_sha1_cached or "").strip()

            if not base_url or not login or not pass_sha1:
                QMessageBox.warning(self, "Внимание", "Не заданы параметры подключения iiko.")
                return

            client = IikoRmsClient(base_url=base_url, login=login, pass_sha1=pass_sha1)
            products = client.get_products()

            dishes = [DishItem(name=p.name, weight=p.weight, price=p.price) for p in products]
            self._pricelist_dishes = dishes
            self.lblPricelistInfo.setText(f"Загружено из iiko: {len(dishes)}")
            self._update_pricelist_suggestions(self.edDishSearch.text())
        except IikoApiError as e:
            msg = str(e)
            # Если пароль неверный/просрочен — сбросим сохранённый и попросим ввести заново
            if ("401" in msg) or ("Unauthorized" in msg):
                try:
                    self._settings.remove("iiko/pass_sha1")
                    self._settings.remove("iiko/password")
                except Exception:
                    pass
                self._iiko_pass_sha1_cached = ""
                QMessageBox.warning(self, "iiko", "Доступ не получен (401). Данные сброшены — введите пароль ещё раз.")
                return
            QMessageBox.critical(self, "iiko", msg)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def _show_all_pricelist_dishes(self):
        """Показывает список блюд без фильтра (с ограничением по количеству)."""
        try:
            if not self._pricelist_dishes:
                if not self._ensure_iiko_pass_sha1():
                    return
                self._load_pricelist_dishes()

            self.lstDishSuggestions.clear()
            if not self._pricelist_dishes:
                return

            # Чтобы UI не зависал на огромной номенклатуре
            limit = 500
            for d in self._pricelist_dishes[:limit]:
                item = QListWidgetItem(self._format_dish_line(d))
                item.setData(Qt.UserRole, d)
                self.lstDishSuggestions.addItem(item)

            if len(self._pricelist_dishes) > limit:
                self.lblPricelistInfo.setText(
                    f"Загружено из iiko: {len(self._pricelist_dishes)} (показаны первые {limit})"
                )
            else:
                self.lblPricelistInfo.setText(f"Загружено из iiko: {len(self._pricelist_dishes)}")

        except IikoApiError as e:
            QMessageBox.critical(self, "iiko", str(e))
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def _update_pricelist_suggestions(self, text: str):
        self.lstDishSuggestions.clear()

        q = (text or "").strip().lower().replace('ё', 'е')
        if len(q) < 2:
            return

        # Автозагрузка подсказок: только если sha1(pass) уже сохранён (чтобы не всплывало окно ввода пароля при наборе текста)
        if (not self._pricelist_dishes) and bool(self._iiko_pass_sha1_cached):
            try:
                self._load_pricelist_dishes()
            except Exception:
                pass

        if not self._pricelist_dishes:
            return

        shown = 0
        for d in self._pricelist_dishes:
            name_norm = d.name.lower().replace('ё', 'е')
            if q in name_norm:
                item = QListWidgetItem(self._format_dish_line(d))
                item.setData(Qt.UserRole, d)
                self.lstDishSuggestions.addItem(item)
                shown += 1
                if shown >= 30:
                    break

    def _add_pricelist_selected(self, d: DishItem):
        key = self._pl_key(d.name)
        if not key:
            return
        if key in self._pricelist_selected_keys:
            return

        it = QListWidgetItem(self._format_dish_line(d))
        it.setData(Qt.UserRole, d)
        it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
        it.setCheckState(Qt.Checked)
        self.lstSelectedDishes.addItem(it)
        self._pricelist_selected_keys.add(key)

    def _on_pricelist_suggestion_clicked(self, item: QListWidgetItem):
        try:
            d = item.data(Qt.UserRole)
            if isinstance(d, DishItem):
                self._add_pricelist_selected(d)
        except Exception:
            pass

    def _add_pricelist_from_enter(self):
        """Enter в поле поиска: добавляем точное совпадение или первый пункт из подсказок."""
        try:
            if not self._pricelist_dishes:
                QMessageBox.warning(self, "Внимание", "Сначала нажмите 'Загрузить блюда'.")
                return

            q_raw = (self.edDishSearch.text() or "").strip()
            if not q_raw:
                return
            q = self._pl_key(q_raw)

            # 1) точное совпадение по названию
            for d in self._pricelist_dishes:
                if self._pl_key(d.name) == q:
                    self._add_pricelist_selected(d)
                    return

            # 2) иначе — первый элемент подсказок
            if self.lstDishSuggestions.count() > 0:
                it = self.lstDishSuggestions.item(0)
                d = it.data(Qt.UserRole)
                if isinstance(d, DishItem):
                    self._add_pricelist_selected(d)
                    return

            QMessageBox.information(self, "Не найдено", "По вашему запросу нет совпадений в загруженном меню.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def _clear_pricelist_selection(self):
        self.lstSelectedDishes.clear()
        self._pricelist_selected_keys = set()

    def do_create_pricelist_excel(self):
        try:
            # берем только отмеченные галочкой
            selected: List[DishItem] = []
            for i in range(self.lstSelectedDishes.count()):
                it = self.lstSelectedDishes.item(i)
                if it.checkState() != Qt.Checked:
                    continue
                d = it.data(Qt.UserRole)
                if isinstance(d, DishItem):
                    selected.append(d)

            if not selected:
                QMessageBox.warning(self, "Внимание", "Выберите хотя бы одно блюдо (поставьте галочку).")
                return

            # Предлагаем имя файла
            desktop = Path.home() / "Desktop"
            stamp = datetime.now().strftime("%d.%m.%Y")
            suggested_path = str(desktop / f"ценники_{stamp}.xlsx")

            save_path, _ = QFileDialog.getSaveFileName(
                self,
                "Сохранить ценники",
                suggested_path,
                "Excel (*.xlsx);;Все файлы (*.*)",
            )
            if not save_path:
                return

            create_pricelist_xlsx(selected, save_path)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

