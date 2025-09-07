import os
import sys
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

from PySide6.QtCore import Qt, QMimeData, QSize
from PySide6.QtGui import QPalette, QColor, QIcon, QPixmap, QPainter, QPen, QBrush, QLinearGradient, QFont
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QTextEdit, QComboBox, QLineEdit,
    QGroupBox, QCheckBox, QSpinBox, QRadioButton, QButtonGroup, QMessageBox, QFrame, QSizePolicy, QScrollArea,
)

from comparator import compare_and_highlight, get_sheet_names, ColumnParseError
from template_linker import default_template_path
from theme import ThemeMode, apply_theme
from presentation_handler import create_presentation_with_excel_data


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
    font = lbl.font()
    font.setBold(True)
    lbl.setFont(font)
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
    size = 256
    pix = QPixmap(size, size)
    pix.fill(Qt.transparent)
    p = QPainter(pix)
    try:
        p.setRenderHint(QPainter.Antialiasing, True)
        # Фон — круг с градиентом (теплые оттенки)
        grad = QLinearGradient(0, 0, size, size)
        grad.setColorAt(0.0, QColor("#FF7E5F"))
        grad.setColorAt(1.0, QColor("#FD3A69"))
        p.setBrush(QBrush(grad))
        p.setPen(Qt.NoPen)
        margin = 12
        p.drawEllipse(margin, margin, size - 2 * margin, size - 2 * margin)

        # Светлая окантовка
        p.setPen(QPen(QColor(255, 255, 255, 230), 6))
        p.setBrush(Qt.NoBrush)
        p.drawEllipse(margin + 3, margin + 3, size - 2 * (margin + 3), size - 2 * (margin + 3))

        # Буква "М"
        f = QFont()
        f.setFamily("Segoe UI")
        f.setBold(True)
        f.setPointSize(120)
        p.setFont(f)
        p.setPen(QColor(255, 255, 255))
        p.drawText(pix.rect(), Qt.AlignCenter, "М")
    finally:
        p.end()
    return QIcon(pix)


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
        self.setWindowTitle("Работа с меню")
        self.setWindowIcon(create_app_icon())
        self.resize(1000, 760)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # Панель управления (сверху)
        topBar = QFrame(); topBar.setObjectName("topBar")
        layTop = QHBoxLayout(topBar)
        layTop.setContentsMargins(12, 8, 12, 8)
        layTop.setSpacing(10)

        lblTheme = QLabel("Тема:")
        self.cmbTheme = QComboBox()
        self.cmbTheme.addItems(["Системная", "Светлая", "Тёмная"])
        self.cmbTheme.setCurrentIndex(0)
        self.cmbTheme.currentIndexChanged.connect(self.on_theme_changed)

        self.btnDownloadTemplate = QPushButton("Сделать шаблон")
        self.btnDownloadTemplate.clicked.connect(self.do_download_template)
        self.btnMakePresentation = QPushButton("Сделать презентацию")
        self.btnMakePresentation.clicked.connect(self.do_make_presentation)
        self.btnBrokerage = QPushButton("Брокеражный журнал")
        self.btnBrokerage.clicked.connect(self.do_brokerage_journal)
        self.btnShowCompare = QPushButton("Сравнить меню")
        self.btnShowCompare.clicked.connect(self.show_compare_sections)

        layTop.addWidget(lblTheme)
        layTop.addWidget(self.cmbTheme)
        layTop.addStretch(1)
        layTop.addWidget(self.btnShowCompare)
        layTop.addWidget(self.btnDownloadTemplate)
        layTop.addWidget(self.btnMakePresentation)
        layTop.addWidget(self.btnBrokerage)

        # Небольшое оформление панели управления через стили
        self.setStyleSheet(
            """
            #topBar {
                border: 1px solid palette(Mid);
                border-radius: 8px;
                background: palette(Base);
            }
            #topBar QPushButton {
                padding: 6px 12px;
                font-size: 14px;
                font-weight: 600;
            }
            #topBar QComboBox {
                padding: 4px 8px;
                min-width: 160px;
                font-size: 14px;
            }
            #topBar QLabel {
                font-weight: 600;
            }
            /* Кнопка после параметров — стиль как на панели управления */
            #actionsPanel QPushButton {
                padding: 6px 12px;
                font-size: 14px;
                font-weight: 600;
            }
            /* У группы параметров компактный стиль без рамки */
            QGroupBox#paramsBox {
                border: none;
                margin: 0px;
                padding: 0px;
                font-weight: 600;
            }
            QGroupBox#paramsBox::title {
                subcontrol-origin: content;
                subcontrol-position: top left;
                left: 0px;
                top: -2px; /* поднимаем заголовок вплотную к контенту */
                padding: 0px;
                margin: 0px;
            }
            /* Компактные элементы внутри параметров без рамки */
            #paramsFrame QCheckBox, #paramsFrame QLabel {
                padding: 2px;
                margin: 0px 6px 0px 0px; /* небольшой горизонтальный зазор между элементами */
            }
            #paramsFrame QCheckBox::indicator {
                width: 14px;
                height: 14px; /* квадратные галочки в параметрах */
            }
            #paramsFrame QSpinBox {
                min-height: 20px;
                padding: 2px 4px;
            }
            /* Убираем все отступы у контейнера параметров */
            #paramsFrame {
                border: none;
                padding: 0px;
                margin: 0px;
            }
            /* Компактные стили для панели презентаций */
            #presentationPanel {
                padding: 0px;
                margin: 0px;
            }
            #presentationPanel QGroupBox {
                padding: 8px;
                margin: 4px 0px;
                font-weight: 600;
            }
            #presentationPanel QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0px 5px;
                margin: 10px 0px;
                left: 8px;
                top: -8px;
            }
            /* Компактные отступы внутри группы */
            #presentationPanel .QWidget {
                margin: 2px 0px;
            }
            #presentationPanel QLabel {
                margin: 2px 0px;
                padding: 0px;
            }
            #presentationPanel QLineEdit, #presentationPanel QPushButton {
                margin: 2px 0px;
            }
            """
        )

        self.topGroup = nice_group("Панель управления", topBar)
        self.topGroup.setObjectName("topGroup")
        self.topGroup.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        root.addWidget(self.topGroup)

        # Область прокрутки для остального содержимого, чтобы панель всегда была наверху
        self.scrollArea = QScrollArea()
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setFrameShape(QFrame.NoFrame)
        self.contentContainer = QWidget()
        self.contentLayout = QVBoxLayout(self.contentContainer)
        self.contentLayout.setContentsMargins(0, 0, 0, 0)
        self.contentLayout.setSpacing(0)
        self.scrollArea.setWidget(self.contentContainer)
        root.addWidget(self.scrollArea, 1)

        # Панель для работы с презентациями
        self.presentationPanel = QWidget()
        self.presentationPanel.setObjectName("presentationPanel")
        pres_layout = QVBoxLayout(self.presentationPanel)
        pres_layout.setContentsMargins(0, 10, 0, 0)  # отступ 10px сверху
        pres_layout.setSpacing(5)  # маленькое расстояние между элементами
        
        # Excel файл для презентации
        self.edExcelPath = DropLineEdit()
        self.edExcelPath.setPlaceholderText("Выберите Excel файл с меню для презентации...")
        self.btnBrowseExcel = QPushButton("Обзор…")
        self.btnBrowseExcel.clicked.connect(lambda: self.browse_excel_file())
        
        excel_row = QWidget()
        excel_layout = QHBoxLayout(excel_row)
        excel_layout.setContentsMargins(8, 5, 8, 5)  # добавляем отступы чтобы не обрезалось
        excel_layout.setSpacing(8)
        excel_layout.addWidget(self.edExcelPath, 1)
        self.btnBrowseExcel.setFixedSize(80, 30)  # фиксированный размер кнопки
        excel_layout.addWidget(self.btnBrowseExcel)
        
        excel_group = QWidget()
        excel_group_layout = QVBoxLayout(excel_group)
        excel_group_layout.setContentsMargins(8, 8, 8, 8)  # добавляем отступы для группы
        excel_group_layout.setSpacing(8)
        excel_group_layout.addWidget(label_caption("Excel файл с меню"))
        excel_group_layout.addWidget(excel_row)
        
        self.grpExcelFile = FileDropGroup("Файл меню для презентации", self.edExcelPath, excel_group)
        # Установка минимальной высоты для компактности
        self.grpExcelFile.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.grpExcelFile.setMinimumHeight(130)
        pres_layout.addWidget(self.grpExcelFile)
        
        # Добавляем растягивающий элемент, чтобы панель была прижата к верху
        pres_layout.addStretch(1)
        
        self.contentLayout.addWidget(self.presentationPanel)
        self.presentationPanel.setVisible(False)
        
        # Панель действий внизу для сравнения (фиксированная)
        self.actionsPanel = QWidget(); self.actionsPanel.setObjectName("actionsPanel")
        la = QHBoxLayout(self.actionsPanel)
        la.setContentsMargins(0, 2, 0, 0)  # минимальный отступ сверху
        self.btnCompare = QPushButton("Сравнить и подсветить")
        self.btnCompare.clicked.connect(self.do_compare)
        la.addStretch(1)
        la.addWidget(self.btnCompare)
        root.addWidget(self.actionsPanel)
        self.actionsPanel.setVisible(False)
        
        # Панель действий внизу для презентаций (фиксированная)
        self.presentationActionsPanel = QWidget(); self.presentationActionsPanel.setObjectName("actionsPanel")
        pla = QHBoxLayout(self.presentationActionsPanel)
        pla.setContentsMargins(0, 8, 0, 0)  # небольшой отступ сверху
        self.btnCreatePresentationWithData = QPushButton("Скачать презентацию с данными")
        self.btnCreatePresentationWithData.clicked.connect(self.do_create_presentation_with_data)
        pla.addStretch(1)
        pla.addWidget(self.btnCreatePresentationWithData)
        root.addWidget(self.presentationActionsPanel)
        self.presentationActionsPanel.setVisible(False)

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
        # Уменьшаем высоту панели сравнения
        self.grpFirst.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.grpFirst.setMinimumHeight(45)
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
        # Уменьшаем высоту второй панели сравнения
        self.grpSecond.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.grpSecond.setMinimumHeight(45)
        self.contentLayout.addWidget(self.grpSecond)
        self.grpSecond.setVisible(False)

        # (дополнительно) — сворачиваемая группа
        opts = QWidget(); lo = QHBoxLayout(opts)
        lo.setContentsMargins(0, 0, 0, 0)
        lo.setSpacing(8)  # немного больше для удобства чтения
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
        self.paramsBox.setObjectName("paramsBox")
        self.paramsBox.setCheckable(True)
        self.paramsBox.setChecked(False)
        # Устанавливаем компактную высоту для панели параметров
        self.paramsBox.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.paramsBox.setMaximumHeight(40)
        lparams = QVBoxLayout(self.paramsBox)
        lparams.setContentsMargins(0, 0, 0, 0)  # полностью убираем отступы
        lparams.setSpacing(0)  # убираем промежутки между элементами
        # Контейнер параметров без дополнительной рамки
        self._paramsFrame = QFrame(); self._paramsFrame.setObjectName("paramsFrame")
        lf = QHBoxLayout(self._paramsFrame)
        lf.setContentsMargins(0, 0, 0, 0)  # убираем отступы
        lf.setSpacing(0)
        lf.addWidget(opts)
        lparams.addWidget(self._paramsFrame)
        self._paramsFrame.setVisible(False)
        self.paramsBox.toggled.connect(self.on_params_toggled)
        self.contentLayout.addWidget(self.paramsBox)
        self.paramsBox.setVisible(False)


        # Theming (System alias Dark by default)
        self._theme_mode = ThemeMode.DARK  # «Системная» ведёт себя как «Тёмная»
        apply_theme(QApplication.instance(), self._theme_mode)

    def log(self, msg: str):
        # Лог отключён по запросу — ничего не делаем
        pass

    def on_theme_changed(self, idx: int):
        if idx == 0:
            # «Системная» ведёт себя как «Тёмная», чтобы визуально не отличалась
            self._theme_mode = ThemeMode.DARK
        elif idx == 1:
            self._theme_mode = ThemeMode.LIGHT
        else:
            self._theme_mode = ThemeMode.DARK
        apply_theme(QApplication.instance(), self._theme_mode)


    def closeEvent(self, event):
        try:
            if hasattr(self, "_theme_timer") and self._theme_timer is not None:
                self._theme_timer.stop()
        except Exception:
            pass
        super().closeEvent(event)

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
            try:
                # Всегда авто по дате
                out_path, matches = compare_and_highlight(
                    path1=p1, sheet1=s1,
                    path2=p2, sheet2=s2,
                    col1="A",
                    col2="E",
                    header_row1=1,
                    header_row2=1,
                    ignore_case=self.chkIgnoreCase.isChecked(),
                    use_fuzzy=self.chkFuzzy.isChecked(),
                    fuzzy_threshold=int(self.spinFuzzy.value()),
                    final_choice=2,
                )
                self.log(f"Готово. Совпадений: {matches}. Итоговый файл: {out_path}")
                QMessageBox.information(self, "Готово", f"Совпадений: {matches}\nИтоговый файл: {out_path}")
            except ColumnParseError as e:
                QMessageBox.warning(self, "Колонка", str(e))
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def show_compare_sections(self):
        try:
            # Скрываем панель презентаций и её панель действий
            if hasattr(self, "presentationPanel"):
                self.presentationPanel.setVisible(False)
            if hasattr(self, "presentationActionsPanel"):
                self.presentationActionsPanel.setVisible(False)
            
            # Показать формы сравнения и панель действий
            if hasattr(self, "grpFirst"):
                self.grpFirst.setVisible(True)
            if hasattr(self, "grpSecond"):
                self.grpSecond.setVisible(True)
            if hasattr(self, "paramsBox"):
                # показываем группу, но оставляем скрытой по умолчанию
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
        try:
            tpl = default_template_path()
            if not Path(tpl).exists():
                QMessageBox.warning(self, "Шаблон", f"Шаблон не найден: {tpl}\nСначала положите файл в templates/menu_template.xls")
                return
            suggested = str(Path.home() / "menu_template.xls")
            out_path, _ = QFileDialog.getSaveFileName(self, "Сохранить шаблон", suggested, "Excel (*.xls)")
            if not out_path:
                return
            shutil.copy2(tpl, out_path)
            self.log(f"Шаблон сохранён: {out_path}")
            QMessageBox.information(self, "Готово", f"Шаблон сохранён:\n{out_path}")
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
            
            # Показываем панель для работы с презентациями и её панель действий
            if hasattr(self, "presentationPanel"):
                self.presentationPanel.setVisible(True)
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
            
            # Выбираем путь для сохранения итоговой презентации
            suggested = str(Path.home() / "презентация_с_меню.pptx")
            out_path, _ = QFileDialog.getSaveFileName(
                self, 
                "Сохранить презентацию с меню", 
                suggested, 
                "PowerPoint (*.pptx)"
            )
            if not out_path:
                return
                
            # Создаем презентацию с данными
            success, message = create_presentation_with_excel_data(
                template_path, 
                excel_path, 
                out_path
            )
            
            if success:
                QMessageBox.information(self, "Готово", f"Презентация создана успешно!\n{message}\nФайл: {out_path}")
            else:
                QMessageBox.warning(self, "Ошибка", f"Не удалось создать презентацию:\n{message}")
                
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

    def do_brokerage_journal(self):
        try:
            tpl = find_template("brokerage_journal_template.xlsx")
            if not tpl:
                QMessageBox.information(self, "Брокеражный журнал", "Шаблон брокеражного журнала не найден. Положите файл brokerage_journal_template.xlsx в папку templates.")
                return
            suggested = str(Path.home() / "брокеражный_журнал.xlsx")
            out_path, _ = QFileDialog.getSaveFileName(self, "Сохранить брокеражный журнал", suggested, "Excel (*.xlsx)")
            if not out_path:
                return
            shutil.copy2(tpl, out_path)
            QMessageBox.information(self, "Готово", f"Брокеражный журнал сохранён:\n{out_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

