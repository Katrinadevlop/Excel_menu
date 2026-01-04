import os
import sys
import shutil
import logging
import hashlib
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, List, Optional, Tuple
from urllib.parse import urlsplit, parse_qsl

from PySide6.QtCore import Qt, QMimeData, QSize, QUrl, QSettings
from PySide6.QtGui import QPalette, QColor, QIcon, QPixmap, QPainter, QPen, QBrush, QLinearGradient, QFont, QDesktopServices
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QBoxLayout,
    QLabel, QPushButton, QFileDialog, QTextEdit, QComboBox, QLineEdit,
    QGroupBox, QCheckBox, QSpinBox, QRadioButton, QButtonGroup, QMessageBox, QFrame, QSizePolicy, QScrollArea,
    QListWidget, QListWidgetItem, QInputDialog, QDialog, QDialogButtonBox,
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
from app.integrations.iiko_cloud_v1_client import IikoCloudV1Client, IikoOrganization
from app.services.menu_template_filler import MenuTemplateFiller
from tools.fill_dynamic_menu import fill_dynamic_menu
from app.gui.ui_styles import (
    AppStyles, ButtonStyles, LayoutStyles, StyleSheets, ComponentStyles,
    StyleManager, ThemeAwareStyles
)


class DropLineEdit(QLineEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setPlaceholderText("Перетащите файл сюда или нажмите Обзор…")


class PasteOnDoubleClickLineEdit(QLineEdit):
    """QLineEdit, который вставляет буфер обмена по двойному клику."""

    def mouseDoubleClickEvent(self, event):
        try:
            self.paste()
        except Exception:
            pass
        super().mouseDoubleClickEvent(event)

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
    def _show_access_token_dialog(self, token: str) -> None:
        """Показывает access_token и даёт кнопку 'Скопировать'."""
        try:
            dlg = QDialog(self)
            dlg.setWindowTitle("iiko — access_token")
            lay = QVBoxLayout(dlg)
            lay.addWidget(QLabel("Токен (access_token):"))

            ed = QLineEdit(dlg)
            ed.setText((token or "").strip())
            ed.setReadOnly(True)
            try:
                ed.setCursorPosition(0)
                ed.selectAll()
            except Exception:
                pass
            lay.addWidget(ed)

            btns = QHBoxLayout()
            btn_copy = QPushButton("Скопировать")
            btn_close = QPushButton("Закрыть")

            def _copy():
                try:
                    QApplication.clipboard().setText(ed.text())
                except Exception:
                    pass

            btn_copy.clicked.connect(_copy)
            btn_close.clicked.connect(dlg.accept)
            btns.addWidget(btn_copy)
            btns.addStretch(1)
            btns.addWidget(btn_close)
            lay.addLayout(btns)

            dlg.exec()
        except Exception:
            # если диалог не удалось показать — просто игнорируем
            pass
    def _reset_iiko_auth_settings(self) -> None:
        """Сбрасывает сохранённые данные авторизации iiko (чтобы не подставлялись старые секреты/хэши)."""
        keys = [
            # iikoCloud v1
            "iiko/cloud/api_url",
            "iiko/cloud/api_login",
            "iiko/cloud/access_token",
            "iiko/cloud/org_id",
            "iiko/cloud/org_name",

            # legacy iiko.biz
            "iiko/biz/user_secret",
            "iiko/biz/user_id",
            "iiko/biz/access_token",
            "iiko/biz/org_id",
            "iiko/biz/org_name",
            "iiko/biz/api_url",

            # REST
            "iiko/pass_sha1",
            "iiko/password",
        ]
        for k in keys:
            try:
                self._settings.remove(k)
            except Exception:
                pass

        # сбрасываем кэш-поля в рантайме
        try:
            self._iiko_cloud_api_url = "https://api-ru.iiko.services"
            self._iiko_cloud_api_login = ""
            self._iiko_cloud_access_token = ""
            self._iiko_cloud_org_id = ""
            self._iiko_cloud_org_name = ""

            # legacy iiko.biz
            self._iiko_biz_user_secret = ""
            self._iiko_biz_user_id = "pos_login_f13591ea"
            self._iiko_biz_org_id = ""
            self._iiko_biz_org_name = ""
            self._iiko_biz_api_url = "https://iiko.biz:9900"
            self._iiko_biz_access_token = ""
            self._iiko_pass_sha1_cached = ""
        except Exception:
            pass

        try:
            self._iiko_products_by_key = {}
            self._pricelist_dishes = []
        except Exception:
            pass

    def _prompt_text(self, title: str, label: str, echo_mode=QLineEdit.Normal, default_text: str = "") -> Optional[str]:
        """Диалог ввода текста, где двойной клик в поле вставляет буфер обмена."""
        dlg = QDialog(self)
        dlg.setWindowTitle(title)
        lay = QVBoxLayout(dlg)
        lay.addWidget(QLabel(label))

        ed = PasteOnDoubleClickLineEdit(dlg)
        ed.setEchoMode(echo_mode)
        if default_text:
            ed.setText(default_text)
            try:
                ed.selectAll()
            except Exception:
                pass
        lay.addWidget(ed)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        lay.addWidget(bb)

        ed.setFocus()
        if dlg.exec() != QDialog.Accepted:
            return None
        return (ed.text() or "").strip()

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
        self.btnOpenMenu = QPushButton("Авторизация точки")
        self.btnOpenMenu.clicked.connect(self.do_authorize_point)

        self.btnOpenTomorrowDishes = QPushButton("Открыть блюда")
        self.btnOpenTomorrowDishes.clicked.connect(self.do_open_tomorrow_dishes)

        self.btnDocuments = QPushButton("Документы")
        self.btnDocuments.clicked.connect(self.show_documents_tab)

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
        # Временно скрываем кнопки iiko (авторизация/открыть/ценники) с панели
        # self.layTop.addWidget(self.btnOpenMenu)
        # self.layTop.addWidget(self.btnOpenTomorrowDishes)
        # self.layTop.addWidget(self.btnDownloadPricelists)
        self.layTop.addWidget(self.btnDocuments)

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

        # Источник блюд: iiko (REST / iikoCloud)
        self._iiko_mode = str(self._settings.value("iiko/mode", "cloud"))  # cloud | rest

        # REST (resto)
        self._iiko_base_url = str(self._settings.value("iiko/base_url", "https://287-772-687.iiko.it/resto"))
        self._iiko_login = str(self._settings.value("iiko/login", "user"))
        # Храним только sha1-хэш пароля (как требует iikoRMS resto API).
        self._iiko_pass_sha1_cached = str(self._settings.value("iiko/pass_sha1", ""))

        # iikoCloud v1 (api-ru.iiko.services): apiLogin -> access_token
        self._iiko_cloud_api_url = str(self._settings.value("iiko/cloud/api_url", "https://api-ru.iiko.services"))
        self._iiko_cloud_api_login = str(self._settings.value("iiko/cloud/api_login", ""))
        self._iiko_cloud_access_token = str(self._settings.value("iiko/cloud/access_token", ""))
        self._iiko_cloud_org_id = str(self._settings.value("iiko/cloud/org_id", ""))
        self._iiko_cloud_org_name = str(self._settings.value("iiko/cloud/org_name", ""))

        # legacy iiko.biz (оставлено на случай старых настроек)
        self._iiko_biz_api_url = str(self._settings.value("iiko/biz/api_url", "https://iiko.biz:9900"))
        self._iiko_biz_user_id = str(self._settings.value("iiko/biz/user_id", "pos_login_f13591ea"))
        self._iiko_biz_user_secret = str(self._settings.value("iiko/biz/user_secret", ""))
        self._iiko_biz_access_token = str(self._settings.value("iiko/biz/access_token", ""))
        self._iiko_biz_org_id = str(self._settings.value("iiko/biz/org_id", ""))
        self._iiko_biz_org_name = str(self._settings.value("iiko/biz/org_name", ""))

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

        # ===== ОТКРЫТИЕ БЛЮД НА ЗАВТРА (из Excel -> снять со стоп-листа в iiko) =====
        self._tomorrow_menu_dishes: List[DishItem] = []
        self._iiko_products_by_key: dict[str, str] = {}
        self._suppress_tomorrow_item_changed = False

        self.lblTomorrowInfo = QLabel(
            "1) Выберите Excel меню на завтра  2) Нажмите 'Загрузить из Excel'  3) Поставьте галочку — блюдо откроется (снимем со стоп-листа)"
        )

        self.btnLoadTomorrowFromExcel = QPushButton("Загрузить из Excel")
        self.btnLoadTomorrowFromExcel.clicked.connect(self._load_tomorrow_dishes_from_excel)

        self.edTomorrowSearch = QLineEdit()
        self.edTomorrowSearch.setPlaceholderText("Поиск по списку…")
        self.edTomorrowSearch.textChanged.connect(self._filter_tomorrow_dishes)

        self.lstTomorrowDishes = QListWidget()
        self.lstTomorrowDishes.setMinimumHeight(320)
        self.lstTomorrowDishes.itemChanged.connect(self._on_tomorrow_item_changed)

        tomorrow_box = QWidget(); tomorrow_layout = QVBoxLayout(tomorrow_box)
        tomorrow_layout.addWidget(self.lblTomorrowInfo)
        tomorrow_layout.addWidget(self.btnLoadTomorrowFromExcel)
        tomorrow_layout.addWidget(self.edTomorrowSearch)
        tomorrow_layout.addWidget(label_caption("Блюда на завтра (галочка = открыть)"))
        tomorrow_layout.addWidget(self.lstTomorrowDishes)

        self.grpTomorrowOpen = nice_group("Открыть блюда на завтра (iiko стоп-лист)", tomorrow_box)
        self.contentLayout.addWidget(self.grpTomorrowOpen)
        self.grpTomorrowOpen.setVisible(False)

        # ===== ДОКУМЕНТЫ: быстрый доступ к файлам =====
        docs_box = QWidget(); docs_layout = QVBoxLayout(docs_box)
        docs_layout.setSpacing(AppStyles.CONTENT_SPACING)

        self.btnVacationStatement = QPushButton("Заявление на отпуск")
        self.btnVacationStatement.clicked.connect(self.open_vacation_statement)
        StyleManager.style_action_button(self.btnVacationStatement)

        self.btnMedicalBooks = QPushButton("Открыть медкнижки (Excel)")
        self.btnMedicalBooks.clicked.connect(self.open_med_books)
        StyleManager.style_action_button(self.btnMedicalBooks)

        self.btnBirthdayFile = QPushButton("Открыть файл \"День рождения\"")
        self.btnBirthdayFile.clicked.connect(self.open_birthday_file)
        StyleManager.style_action_button(self.btnBirthdayFile)

        self.btnHygieneJournal = QPushButton("Открыть гигиенический журнал")
        self.btnHygieneJournal.clicked.connect(self.open_hygiene_journal)
        StyleManager.style_action_button(self.btnHygieneJournal)

        self.btnDirection = QPushButton("Направление")
        self.btnDirection.clicked.connect(self.open_direction_document)
        StyleManager.style_action_button(self.btnDirection)

        self.btnLockerDoc = QPushButton("Раздевалка")
        self.btnLockerDoc.clicked.connect(self.open_locker_document)
        StyleManager.style_action_button(self.btnLockerDoc)

        self.btnFridgeTemp = QPushButton("Температура холодильник")
        self.btnFridgeTemp.clicked.connect(self.open_fridge_temperature)
        StyleManager.style_action_button(self.btnFridgeTemp)

        self.btnFreezerTemp = QPushButton("Температура морозилка")
        self.btnFreezerTemp.clicked.connect(self.open_freezer_temperature)
        StyleManager.style_action_button(self.btnFreezerTemp)

        self.btnBuffetSheet = QPushButton("Буфет бумажка")
        self.btnBuffetSheet.clicked.connect(self.open_buffet_sheet)
        StyleManager.style_action_button(self.btnBuffetSheet)

        self.btnBakerSheet = QPushButton("Пекарь бумажка")
        self.btnBakerSheet.clicked.connect(self.open_baker_sheet)
        StyleManager.style_action_button(self.btnBakerSheet)

        self.btnFryerJournal = QPushButton("Журнал фритюрного масла")
        self.btnFryerJournal.clicked.connect(self.open_fryer_oil_journal)
        StyleManager.style_action_button(self.btnFryerJournal)

        docs_layout.addWidget(self.btnVacationStatement)
        docs_layout.addWidget(self.btnMedicalBooks)
        docs_layout.addWidget(self.btnBirthdayFile)
        docs_layout.addWidget(self.btnHygieneJournal)
        docs_layout.addWidget(self.btnDirection)
        docs_layout.addWidget(self.btnLockerDoc)
        docs_layout.addWidget(self.btnFridgeTemp)
        docs_layout.addWidget(self.btnFreezerTemp)
        docs_layout.addWidget(self.btnBuffetSheet)
        docs_layout.addWidget(self.btnBakerSheet)
        docs_layout.addWidget(self.btnFryerJournal)
        docs_layout.addStretch(1)

        self.grpDocuments = nice_group("Документы", docs_box)
        self.contentLayout.addWidget(self.grpDocuments)
        self.grpDocuments.setVisible(False)

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

        # Панель действий внизу для "Открыть блюда" (фиксированная)
        self.tomorrowOpenActionsPanel = QWidget(); self.tomorrowOpenActionsPanel.setObjectName("actionsPanel")
        self.tomorrowOpenActionsLayout = QHBoxLayout(self.tomorrowOpenActionsPanel)
        LayoutStyles.apply_margins(self.tomorrowOpenActionsLayout, LayoutStyles.CONTENT_TOP_MARGIN)
        self.btnOpenTomorrowChecked = QPushButton("Открыть отмеченные")
        self.btnOpenTomorrowChecked.clicked.connect(self._open_tomorrow_checked)
        self.tomorrowOpenActionsLayout.addStretch(1)
        self.tomorrowOpenActionsLayout.addWidget(self.btnOpenTomorrowChecked)
        self.rootLayout.addWidget(self.tomorrowOpenActionsPanel)
        self.tomorrowOpenActionsPanel.setVisible(False)

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
            if hasattr(self, "grpTomorrowOpen"):
                self.grpTomorrowOpen.setVisible(False)
            if hasattr(self, "tomorrowOpenActionsPanel"):
                self.tomorrowOpenActionsPanel.setVisible(False)
            if hasattr(self, "grpDocuments"):
                self.grpDocuments.setVisible(False)

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

    def show_documents_tab(self) -> None:
        """Показывает вкладку "Документы" с кнопками для открытия файлов."""
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
            if hasattr(self, "grpTomorrowOpen"):
                self.grpTomorrowOpen.setVisible(False)
            if hasattr(self, "tomorrowOpenActionsPanel"):
                self.tomorrowOpenActionsPanel.setVisible(False)

            # Показываем блок с документами
            if hasattr(self, "grpDocuments"):
                self.grpDocuments.setVisible(True)
                if hasattr(self, "scrollArea"):
                    self.scrollArea.ensureWidgetVisible(self.grpDocuments)
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
            if hasattr(self, "grpTomorrowOpen"):
                self.grpTomorrowOpen.setVisible(False)
            if hasattr(self, "tomorrowOpenActionsPanel"):
                self.tomorrowOpenActionsPanel.setVisible(False)
            if hasattr(self, "grpDocuments"):
                self.grpDocuments.setVisible(False)
            
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
            if hasattr(self, "grpTomorrowOpen"):
                self.grpTomorrowOpen.setVisible(False)
            if hasattr(self, "tomorrowOpenActionsPanel"):
                self.tomorrowOpenActionsPanel.setVisible(False)
            if hasattr(self, "grpDocuments"):
                self.grpDocuments.setVisible(False)
            
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

    def do_authorize_point(self):
        """Авторизация iikoCloud API v1 (api-ru.iiko.services) через apiLogin."""
        try:
            # Явно предлагаем "Сбросить" перед вводом, чтобы точно убрать старые данные.
            msg = QMessageBox(self)
            msg.setWindowTitle("iiko")
            msg.setText("Перед авторизацией:")
            msg.setInformativeText("Можно сбросить сохранённый логин/пароль/токен, чтобы приложение не использовало старые данные.")
            btn_reset = msg.addButton("Сбросить", QMessageBox.DestructiveRole)
            btn_continue = msg.addButton("Продолжить", QMessageBox.AcceptRole)
            msg.addButton("Отмена", QMessageBox.RejectRole)
            msg.exec()

            clicked = msg.clickedButton()
            if clicked is None:
                return
            if clicked == btn_reset:
                self._reset_iiko_auth_settings()
            elif clicked != btn_continue:
                return

            api_url = (self._iiko_cloud_api_url or "").strip() or "https://api-ru.iiko.services"

            api_login_default = (self._iiko_cloud_api_login or "").strip()
            api_login = self._prompt_text(
                "iiko — iikoCloud",
                "apiLogin\r\n(двойной клик в поле = вставить из буфера обмена):",
                QLineEdit.Normal,
                default_text=api_login_default,
            )
            if not api_login:
                return

            client = IikoCloudV1Client(api_url=api_url, api_login=api_login)
            access_token = client.access_token()

            # показываем токен (можно скопировать)
            self._show_access_token_dialog(access_token)

            # Организации
            orgs = client.organizations()
            org_id = ""
            org_name = ""

            if len(orgs) == 1:
                org_id = orgs[0].id
                org_name = orgs[0].name
            else:
                labels = [f"{i+1}. {o.name} ({o.id})" for i, o in enumerate(orgs)]
                chosen, ok = QInputDialog.getItem(
                    self,
                    "iiko — iikoCloud",
                    "Выберите организацию:",
                    labels,
                    0,
                    False,
                )
                if not ok:
                    return
                idx = labels.index(chosen)
                org_id = orgs[idx].id
                org_name = orgs[idx].name

            QMessageBox.information(self, "iiko", f"Подключено к организации: {org_name}")

            self._iiko_mode = "cloud"
            self._iiko_cloud_api_url = api_url
            self._iiko_cloud_api_login = api_login
            self._iiko_cloud_access_token = access_token
            self._iiko_cloud_org_id = org_id
            self._iiko_cloud_org_name = org_name

            # Важно: если меняли пароль — сбросим сохранённые кэши номенклатуры.

            self._iiko_products_by_key = {}
            self._pricelist_dishes = []

            try:
                self._settings.setValue("iiko/mode", "cloud")
                self._settings.setValue("iiko/cloud/api_url", api_url)
                self._settings.setValue("iiko/cloud/api_login", api_login)
                self._settings.setValue("iiko/cloud/access_token", access_token)
                self._settings.setValue("iiko/cloud/org_id", org_id)
                if org_name:
                    self._settings.setValue("iiko/cloud/org_name", org_name)
            except Exception:
                pass

            # Покажем, куда подключились
            if org_name:
                QMessageBox.information(self, "iiko", f"iiko.biz подключено: {org_name} ({org_id})")
            else:
                QMessageBox.information(self, "iiko", f"iiko.biz подключено. organization_id: {org_id}")

        except IikoApiError as e:
            QMessageBox.critical(self, "iiko", str(e))
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

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

    def _open_document(self, settings_key: str, dialog_title: str, file_filter: str) -> None:
        """Открывает произвольный файл документа, запоминая его путь в настройках."""
        try:
            path = ""
            try:
                if hasattr(self, "_settings"):
                    value = self._settings.value(settings_key, "")
                    path = str(value) if value is not None else ""
            except Exception:
                path = ""

            if not path or not Path(path).exists():
                path, _ = QFileDialog.getOpenFileName(
                    self,
                    dialog_title,
                    str(Path.cwd()),
                    file_filter,
                )
                if not path:
                    return
                try:
                    if hasattr(self, "_settings"):
                        self._settings.setValue(settings_key, path)
                except Exception:
                    pass

            ok = QDesktopServices.openUrl(QUrl.fromLocalFile(path))
            if not ok:
                QMessageBox.warning(self, "Не удалось открыть", "Не удалось открыть файл выбранной программой.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def open_vacation_statement(self) -> None:
        """Кнопка "Заявление на отпуск" — просто открывает шаблон в Word, без копирования на рабочий стол."""
        try:
            template_path = find_template("Заявление на отпуск.doc")
            if not template_path:
                QMessageBox.warning(
                    self,
                    "Шаблон",
                    "Файл 'Заявление на отпуск.doc' не найден. Положите его в папку templates.",
                )
                return

            # Сразу открываем шаблон в приложении по умолчанию (обычно Word)
            ok = QDesktopServices.openUrl(QUrl.fromLocalFile(str(template_path)))
            if not ok:
                QMessageBox.warning(
                    self,
                    "Открытие",
                    f"Не удалось автоматически открыть файл:\n{template_path}",
                )
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

    def open_med_books(self) -> None:
        """Кнопка "Открыть медкнижки" — открывает шаблон Медкнижки.xlsx из templates."""
        try:
            template_path = find_template("Медкнижки.xlsx")
            if not template_path:
                QMessageBox.warning(
                    self,
                    "Шаблон",
                    "Файл 'Медкнижки.xlsx' не найден. Положите его в папку templates.",
                )
                return

            ok = QDesktopServices.openUrl(QUrl.fromLocalFile(str(template_path)))
            if not ok:
                QMessageBox.warning(
                    self,
                    "Открытие",
                    f"Не удалось автоматически открыть файл:\n{template_path}",
                )
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

    def open_birthday_file(self) -> None:
        """Кнопка "Открыть файл День рождения" — открывает шаблон из templates."""
        try:
            template_path = find_template("День рождения.xlsx")
            if not template_path:
                QMessageBox.warning(
                    self,
                    "Шаблон",
                    "Файл 'День рождения.xlsx' не найден. Положите его в папку templates.",
                )
                return

            ok = QDesktopServices.openUrl(QUrl.fromLocalFile(str(template_path)))
            if not ok:
                QMessageBox.warning(
                    self,
                    "Открытие",
                    f"Не удалось автоматически открыть файл:\n{template_path}",
                )
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

    def open_hygiene_journal(self) -> None:
        """Открывает шаблон гигиенического журнала из папки templates."""
        try:
            template_path = find_template("Гигиенический журнал.xlsx")
            if not template_path:
                QMessageBox.warning(
                    self,
                    "Шаблон",
                    "Файл 'Гигиенический журнал.xlsx' не найден. Положите его в папку templates.",
                )
                return

            ok = QDesktopServices.openUrl(QUrl.fromLocalFile(str(template_path)))
            if not ok:
                QMessageBox.warning(
                    self,
                    "Открытие",
                    f"Не удалось автоматически открыть файл:\n{template_path}",
                )
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

    def open_direction_document(self) -> None:
        """Открывает шаблон направления из папки templates (Направление.doc)."""
        try:
            template_path = find_template("Направление.doc")
            if not template_path:
                QMessageBox.warning(
                    self,
                    "Шаблон",
                    "Файл 'Направление.doc' не найден. Положите его в папку templates.",
                )
                return

            ok = QDesktopServices.openUrl(QUrl.fromLocalFile(str(template_path)))
            if not ok:
                QMessageBox.warning(
                    self,
                    "Открытие",
                    f"Не удалось автоматически открыть файл:\n{template_path}",
                )
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

    def open_locker_document(self) -> None:
        """Открывает шаблон документа "Раздевалка" из папки templates (Раздевалка.xlsx)."""
        try:
            template_path = find_template("Раздевалка.xlsx")
            if not template_path:
                QMessageBox.warning(
                    self,
                    "Шаблон",
                    "Файл 'Раздевалка.xlsx' не найден. Положите его в папку templates.",
                )
                return

            ok = QDesktopServices.openUrl(QUrl.fromLocalFile(str(template_path)))
            if not ok:
                QMessageBox.warning(
                    self,
                    "Открытие",
                    f"Не удалось автоматически открыть файл:\n{template_path}",
                )
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

    def open_buffet_sheet(self) -> None:
        """Открывает шаблон "Бланк для раздачи кофетерий" из templates."""
        try:
            template_path = find_template("Бланк для раздачи кофетерий.xlsx")
            if not template_path:
                QMessageBox.warning(
                    self,
                    "Шаблон",
                    "Файл 'Бланк для раздачи кофетерий.xlsx' не найден. Положите его в папку templates.",
                )
                return

            ok = QDesktopServices.openUrl(QUrl.fromLocalFile(str(template_path)))
            if not ok:
                QMessageBox.warning(
                    self,
                    "Открытие",
                    f"Не удалось автоматически открыть файл:\n{template_path}",
                )
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

    def open_baker_sheet(self) -> None:
        """Открывает шаблон "Акт_приготовления_ВЫПЕЧКАДЕСЕРТЫ" из templates."""
        try:
            template_path = find_template("Акт_приготовления_ВЫПЕЧКАДЕСЕРТЫ.xlsx")
            if not template_path:
                QMessageBox.warning(
                    self,
                    "Шаблон",
                    "Файл 'Акт_приготовления_ВЫПЕЧКАДЕСЕРТЫ.xlsx' не найден. Положите его в папку templates.",
                )
                return

            ok = QDesktopServices.openUrl(QUrl.fromLocalFile(str(template_path)))
            if not ok:
                QMessageBox.warning(
                    self,
                    "Открытие",
                    f"Не удалось автоматически открыть файл:\n{template_path}",
                )
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

    def open_fryer_oil_journal(self) -> None:
        """Открывает шаблон "Журнал замены фритюрного масла" из templates."""
        try:
            template_path = find_template("Журнал замены фритюрного масла.doc")
            if not template_path:
                QMessageBox.warning(
                    self,
                    "Шаблон",
                    "Файл 'Журнал замены фритюрного масла.doc' не найден. Положите его в папку templates.",
                )
                return

            ok = QDesktopServices.openUrl(QUrl.fromLocalFile(str(template_path)))
            if not ok:
                QMessageBox.warning(
                    self,
                    "Открытие",
                    f"Не удалось автоматически открыть файл:\n{template_path}",
                )
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

    def open_fridge_temperature(self) -> None:
        """Открывает шаблон "Температурный режим холодильного оборудования" из templates."""
        try:
            template_path = find_template("Температурный_режим_холодильньного_оборудования.docx")
            if not template_path:
                QMessageBox.warning(
                    self,
                    "Шаблон",
                    "Файл 'Температурный_режим_холодильньного_оборудования.docx' не найден. Положите его в папку templates.",
                )
                return

            ok = QDesktopServices.openUrl(QUrl.fromLocalFile(str(template_path)))
            if not ok:
                QMessageBox.warning(
                    self,
                    "Открытие",
                    f"Не удалось автоматически открыть файл:\n{template_path}",
                )
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

    def open_freezer_temperature(self) -> None:
        """Открывает шаблон "Температурный режим морозильного оборудования" из templates."""
        try:
            template_path = find_template("Температурный_режим_морозильного_оборудования.docx")
            if not template_path:
                QMessageBox.warning(
                    self,
                    "Шаблон",
                    "Файл 'Температурный_режим_морозильного_оборудования.docx' не найден. Положите его в папку templates.",
                )
                return

            ok = QDesktopServices.openUrl(QUrl.fromLocalFile(str(template_path)))
            if not ok:
                QMessageBox.warning(
                    self,
                    "Открытие",
                    f"Не удалось автоматически открыть файл:\n{template_path}",
                )
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

    def do_open_tomorrow_dishes(self):
        """Открытие блюд на завтра: берём Excel-меню и снимаем блюда со стоп-листа в iiko."""
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
            if hasattr(self, "grpPricelist"):
                self.grpPricelist.setVisible(False)
            if hasattr(self, "pricelistActionsPanel"):
                self.pricelistActionsPanel.setVisible(False)
            if hasattr(self, "grpDocuments"):
                self.grpDocuments.setVisible(False)

            # Показываем выбор Excel-файла меню на завтра
            if hasattr(self, "grpExcelFile"):
                self.grpExcelFile.setVisible(True)
                self.grpExcelFile.setTitle("Excel меню на завтра")
                self.edExcelPath.setPlaceholderText("Выберите Excel файл меню на завтра (тот же формат, что и для презентации/журнала)…")

            if hasattr(self, "grpTomorrowOpen"):
                self.grpTomorrowOpen.setVisible(True)
            if hasattr(self, "tomorrowOpenActionsPanel"):
                self.tomorrowOpenActionsPanel.setVisible(True)

            # Сброс состояния
            self._tomorrow_menu_dishes = []
            self.edTomorrowSearch.clear()
            self.lstTomorrowDishes.clear()
            self.lblTomorrowInfo.setText(
                "1) Выберите Excel меню на завтра  2) Нажмите 'Загрузить из Excel'  3) Поставьте галочку — блюдо откроется (снимем со стоп-листа)"
            )

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
            if hasattr(self, "grpDocuments"):
                self.grpDocuments.setVisible(False)

            # Для ценников НЕ показываем общий блок выбора Excel-файла (он мешает и не нужен при iiko)
            if hasattr(self, "grpExcelFile"):
                self.grpExcelFile.setVisible(False)

            # Скрываем "Открыть блюда" если было открыто
            if hasattr(self, "grpTomorrowOpen"):
                self.grpTomorrowOpen.setVisible(False)
            if hasattr(self, "tomorrowOpenActionsPanel"):
                self.tomorrowOpenActionsPanel.setVisible(False)

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

            # Автозагрузка блюд из iiko: если авторизация точки уже сделана
            try:
                if (not self._pricelist_dishes) and self._can_autoload_iiko_products():
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
            if hasattr(self, "grpTomorrowOpen"):
                self.grpTomorrowOpen.setVisible(False)
            if hasattr(self, "tomorrowOpenActionsPanel"):
                self.tomorrowOpenActionsPanel.setVisible(False)
            if hasattr(self, "grpDocuments"):
                self.grpDocuments.setVisible(False)
            
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
        """Заполняет шаблон меню сегментным алгоритмом (A-F)."""
        try:
            excel_path = self.edExcelPath.text().strip()
            if not excel_path:
                QMessageBox.warning(self, "Внимание", "Выберите Excel файл с меню.")
                return
            if not Path(excel_path).exists():
                QMessageBox.warning(self, "Ошибка", "Указанный Excel файл не найден.")
                return

            template_path = default_template_path()
            if not template_path or not Path(template_path).exists():
                QMessageBox.warning(self, "Шаблон", "Шаблон меню не найден. Положите файл шаблона в папку templates/.")
                return

            desktop = Path.home() / "Desktop"

            # Имя файла: "<дата месяц> - <день недели>.xlsx" (берём из исходного файла: B3 и B2).
            suggested_name = "menu_ready.xlsx"

            def _sanitize_filename_part(s) -> str:
                text = str(s or "")
                for ch in '<>:"/\\|?*':
                    text = text.replace(ch, " ")
                return " ".join(text.split()).strip()

            try:
                from openpyxl import load_workbook

                wb_in = load_workbook(excel_path, data_only=True)
                if "Касса" in wb_in.sheetnames:
                    ws_in = wb_in["Касса"]
                else:
                    ws_in = wb_in.active

                weekday = _sanitize_filename_part(ws_in["B2"].value)
                day_month = _sanitize_filename_part(ws_in["B3"].value)

                base = ""
                if day_month and weekday:
                    base = f"{day_month} - {weekday}"
                elif day_month:
                    base = day_month
                else:
                    base = _sanitize_filename_part(Path(excel_path).stem)

                if base:
                    suggested_name = f"{base}.xlsx"
                    # чтобы не перезаписать исходный файл по умолчанию
                    if (desktop / suggested_name).exists():
                        suggested_name = f"{base} - готово.xlsx"
            except Exception:
                pass

            suggested_path = str(desktop / suggested_name)
            save_path, _ = QFileDialog.getSaveFileName(
                self,
                "Сохранить динамически заполненный шаблон",
                suggested_path,
                "Excel (*.xlsx);;Все файлы (*.*)",
            )
            if not save_path:
                return

            fill_dynamic_menu(
                Path(excel_path),
                Path(template_path),
                Path(save_path),
                source_sheet="Касса",
                template_sheet="Касса",
            )
            QMessageBox.information(self, "Готово", "Шаблон успешно заполнен по сегментам.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

    # ===== iiko: общие хелперы =====
    def _ensure_iiko_pass_sha1(self) -> bool:
        """Если sha1(pass) не сохранён — спрашиваем пароль один раз и сохраняем только sha1."""
        if self._iiko_pass_sha1_cached:
            return True

        pwd = self._prompt_text(
            "iiko",
            "Введите пароль iiko (сохранится только SHA1-хэш на этом компьютере)\n(двойной клик в поле = вставить из буфера обмена):",
            QLineEdit.Password,
        )
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

    def _can_autoload_iiko_products(self) -> bool:
        """Есть ли сохранённые данные, чтобы на вкладках не просить ввод заново."""
        mode = (self._iiko_mode or "cloud").strip().lower()
        if mode == "cloud":
            return bool(
                (self._iiko_cloud_api_login or "").strip()
                and (self._iiko_cloud_org_id or "").strip()
            )
        return bool(
            (self._iiko_base_url or "").strip()
            and (self._iiko_login or "").strip()
            and (self._iiko_pass_sha1_cached or "").strip()
        )

    def _get_iiko_products(self) -> List[Any]:
        """Возвращает список продуктов iiko в зависимости от выбранного режима (REST/Cloud).

        Для iikoCloud теперь работаем ТОЛЬКО с уже сохранённым access_token:
        - не ходим повторно на /api/1/access_token;
        - если токена нет — просим нажать «Авторизация точки»;
        - если токен перестал работать (401/403 и т.п.) — показываем ошибку, без автопереавторизации.
        """
        mode = (self._iiko_mode or "cloud").strip().lower()

        if mode == "cloud":
            api_url = (self._iiko_cloud_api_url or "").strip() or "https://api-ru.iiko.services"
            api_login = (self._iiko_cloud_api_login or "").strip()
            org_id = (self._iiko_cloud_org_id or "").strip()

            if not org_id:
                raise IikoApiError("Не выбрана организация. Нажмите 'Авторизация точки'.")

            token = (self._iiko_cloud_access_token or "").strip()
            if not token:
                raise IikoApiError(
                    "Не найден access_token iikoCloud. Сначала нажмите 'Авторизация точки'."
                )

            client = IikoCloudV1Client(
                api_url=api_url,
                api_login=api_login,
                organization_id=org_id,
                access_token=token,
            )
            # Если токен протух — просто покажем ошибку пользователю.
            return client.get_products()

        # REST
        if not self._ensure_iiko_pass_sha1():
            raise IikoApiError("Не задан SHA1-хэш пароля для REST. Нажмите 'Авторизация точки'.")

        base_url = (self._iiko_base_url or "").strip()
        login = (self._iiko_login or "").strip()
        pass_sha1 = (self._iiko_pass_sha1_cached or "").strip()
        if not base_url or not login or not pass_sha1:
            raise IikoApiError("Не заданы параметры подключения iiko.")

        client = IikoRmsClient(base_url=base_url, login=login, pass_sha1=pass_sha1)
        return client.get_products()

    def _ensure_iiko_products_index(self) -> bool:
        """Подгружает номенклатуру iiko и строит индекс name->productId (для сопоставления с Excel)."""
        if self._iiko_products_by_key:
            return True

        products = self._get_iiko_products()

        idx: dict[str, str] = {}
        for p in products:
            key = self._pl_key(getattr(p, 'name', ''))
            if not key:
                continue
            pid = (getattr(p, 'product_id', '') or "").strip()
            if not pid:
                continue
            idx.setdefault(key, pid)

        self._iiko_products_by_key = idx
        return True

    def _pl_key(self, name: str) -> str:
        return " ".join(str(name).strip().lower().replace('ё', 'е').split())

    # ===== Открытие блюд на завтра (Excel -> iiko stoplist) =====
    def _load_tomorrow_dishes_from_excel(self):
        try:
            excel_path = self.edExcelPath.text().strip()
            if not excel_path:
                QMessageBox.warning(self, "Внимание", "Выберите Excel файл меню на завтра.")
                return
            if not Path(excel_path).exists():
                QMessageBox.warning(self, "Ошибка", "Указанный Excel файл не найден.")
                return

            # 1) Забираем блюда из Excel
            dishes = extract_all_dishes_with_details(excel_path)
            if not dishes:
                QMessageBox.warning(self, "Не найдено", "В Excel не удалось найти блюда (проверьте формат меню).")
                return

            self._tomorrow_menu_dishes = dishes

            # 2) Готовим индекс iiko, чтобы сопоставить имена -> productId
            if not self._ensure_iiko_products_index():
                return

            self._populate_tomorrow_dish_list()

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def _populate_tomorrow_dish_list(self):
        self._suppress_tomorrow_item_changed = True
        try:
            self.lstTomorrowDishes.clear()

            not_found = 0
            for d in self._tomorrow_menu_dishes:
                name = (d.name or "").strip()
                if not name:
                    continue

                key = self._pl_key(name)
                pid = self._iiko_products_by_key.get(key, "")

                it = QListWidgetItem(name)
                it.setData(Qt.UserRole, pid)
                it.setData(Qt.UserRole + 1, "ready" if pid else "not_found")
                it.setData(Qt.UserRole + 2, name)

                if pid:
                    it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
                    it.setCheckState(Qt.Unchecked)
                else:
                    not_found += 1
                    # не даём поставить галочку
                    it.setFlags(it.flags() & ~Qt.ItemIsUserCheckable)
                    it.setForeground(QBrush(QColor("#888888")))
                    it.setText(f"{name}  (не найдено в iiko)")

                self.lstTomorrowDishes.addItem(it)

            total = self.lstTomorrowDishes.count()
            self.lblTomorrowInfo.setText(
                f"Загружено из Excel: {total}. Не найдено в iiko: {not_found}. "
                "Поставьте галочку у найденных — блюдо откроется (снимем со стоп-листа)."
            )

        finally:
            self._suppress_tomorrow_item_changed = False

    def _filter_tomorrow_dishes(self, text: str):
        q = self._pl_key(text)
        for i in range(self.lstTomorrowDishes.count()):
            it = self.lstTomorrowDishes.item(i)
            base_name = (it.data(Qt.UserRole + 2) or it.text() or "")
            show = (not q) or (q in self._pl_key(base_name))
            it.setHidden(not show)

    def _open_one_tomorrow_item(self, it: QListWidgetItem) -> bool:
        """Снять со стоп-листа.

        Сейчас в проекте оставлен один способ авторизации (iiko.biz по ссылке).
        Снятие со стоп-листа в этом приложении ранее делалось через /resto, но на вашей стороне
        REST часто блокируется лицензией (REST_API(2000)).

        Поэтому здесь пока показываем понятную ошибку, чтобы приложение не просило пароль REST.
        """
        pid = (it.data(Qt.UserRole) or "").strip()
        if not pid:
            return False

        raise IikoApiError(
            "Открытие блюда (снятие со стоп-листа) сейчас не настроено через iiko.biz. "
            "Нужно либо включить REST (/resto) на сервере, либо реализовать 'приказы' через iikoChain." 
        )

    def _on_tomorrow_item_changed(self, item: QListWidgetItem):
        if self._suppress_tomorrow_item_changed:
            return

        try:
            status = item.data(Qt.UserRole + 1)
            if status == "not_found":
                return

            if item.checkState() != Qt.Checked:
                return

            # Поставили галочку -> пробуем "открыть" (снять со стоп-листа)
            self._suppress_tomorrow_item_changed = True
            try:
                item.setData(Qt.UserRole + 1, "opening")
                base_name = (item.data(Qt.UserRole + 2) or item.text() or "")
                item.setText(f"{base_name}  (открываю…)")
            finally:
                self._suppress_tomorrow_item_changed = False

            try:
                self._open_one_tomorrow_item(item)
                self._suppress_tomorrow_item_changed = True
                try:
                    base_name = (item.data(Qt.UserRole + 2) or item.text() or "")
                    item.setData(Qt.UserRole + 1, "opened")
                    item.setForeground(QBrush(QColor("#2e7d32")))
                    item.setText(f"{base_name}  (ОТКРЫТО)")
                    item.setCheckState(Qt.Checked)
                finally:
                    self._suppress_tomorrow_item_changed = False

            except IikoApiError as e:
                self._suppress_tomorrow_item_changed = True
                try:
                    base_name = (item.data(Qt.UserRole + 2) or item.text() or "")
                    item.setData(Qt.UserRole + 1, "failed")
                    item.setForeground(QBrush(QColor("#b71c1c")))
                    item.setText(f"{base_name}  (ошибка открытия)")
                    item.setCheckState(Qt.Unchecked)
                finally:
                    self._suppress_tomorrow_item_changed = False

                QMessageBox.critical(self, "iiko", str(e))

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def _open_tomorrow_checked(self):
        try:
            any_done = False
            for i in range(self.lstTomorrowDishes.count()):
                it = self.lstTomorrowDishes.item(i)
                if it.isHidden():
                    continue
                if it.checkState() != Qt.Checked:
                    continue
                if it.data(Qt.UserRole + 1) == "opened":
                    continue
                try:
                    self._open_one_tomorrow_item(it)
                    it.setData(Qt.UserRole + 1, "opened")
                    it.setForeground(QBrush(QColor("#2e7d32")))
                    base_name = (it.data(Qt.UserRole + 2) or it.text() or "")
                    it.setText(f"{base_name}  (ОТКРЫТО)")
                    any_done = True
                except IikoApiError as e:
                    it.setData(Qt.UserRole + 1, "failed")
                    it.setForeground(QBrush(QColor("#b71c1c")))
                    base_name = (it.data(Qt.UserRole + 2) or it.text() or "")
                    it.setText(f"{base_name}  (ошибка открытия)")
                    QMessageBox.critical(self, "iiko", str(e))
                    return

            if any_done:
                QMessageBox.information(self, "Готово", "Открытие отмеченных блюд завершено.")
            else:
                QMessageBox.information(self, "Нет выбора", "Отметьте блюда галочками, которые нужно открыть.")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def _format_dish_line(self, d: DishItem) -> str:
        parts = [d.name]
        if d.weight:
            parts.append(str(d.weight))
        if d.price:
            parts.append(str(d.price))
        return " — ".join(parts)

    def _load_pricelist_dishes(self):
        try:
            products = self._get_iiko_products()

            dishes = [DishItem(name=p.name, weight=p.weight, price=p.price) for p in products]
            self._pricelist_dishes = dishes
            self.lblPricelistInfo.setText(f"Загружено из iiko: {len(dishes)}")
            self._update_pricelist_suggestions(self.edDishSearch.text())
        except IikoApiError as e:
            QMessageBox.critical(self, "iiko", str(e))
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def _show_all_pricelist_dishes(self):
        """Показывает список блюд без фильтра (с ограничением по количеству)."""
        try:
            if not self._pricelist_dishes:
                if not self._can_autoload_iiko_products():
                    QMessageBox.warning(self, "iiko", "Сначала нажмите 'Авторизация точки'.")
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

        # Автозагрузка подсказок: только если авторизация точки уже сделана
        if (not self._pricelist_dishes) and self._can_autoload_iiko_products():
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

