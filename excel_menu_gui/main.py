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
from theme import ThemeMode, apply_theme, start_system_theme_watcher


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

        self.btnDownloadTemplate = QPushButton("Скачать шаблон")
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
            /* Параметры: рамка только при раскрытии */
            #paramsBox {
                border: 1px solid palette(Mid);
                border-radius: 8px;
                margin-top: 8px;
            }
            #paramsBox[checked=\"false\"] {
                border: none;
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
        self.contentLayout.setSpacing(10)
        self.scrollArea.setWidget(self.contentContainer)
        root.addWidget(self.scrollArea, 1)

        # File 1
        self.edPath1 = DropLineEdit()
        self.btnBrowse1 = QPushButton("Обзор…")
        self.btnBrowse1.clicked.connect(lambda: self.browse_file(self.edPath1, which=1))
        self.edCol1 = QLineEdit("A")
        self.edCol1.setMaximumWidth(60)
        self.spinHdr1 = QSpinBox()
        self.spinHdr1.setRange(1, 10000)
        self.spinHdr1.setValue(1)

        row1 = QWidget(); r1 = QHBoxLayout(row1)
        r1.addWidget(self.edPath1, 1)
        r1.addWidget(self.btnBrowse1)
        row1b = QWidget(); r1b = QHBoxLayout(row1b)
        r1b.addWidget(label_caption("Колонка блюд:"))
        r1b.addWidget(self.edCol1)
        r1b.addWidget(label_caption("Строка заголовка:"))
        r1b.addWidget(self.spinHdr1)

        g1 = QWidget(); l1 = QVBoxLayout(g1)
        l1.addWidget(label_caption("Файл 1"))
        l1.addWidget(row1)
        l1.addWidget(row1b)
        self.grpFirst = nice_group("Первый файл", g1)
        self.contentLayout.addWidget(self.grpFirst)
        self.grpFirst.setVisible(False)

        # File 2
        self.edPath2 = DropLineEdit()
        self.btnBrowse2 = QPushButton("Обзор…")
        self.btnBrowse2.clicked.connect(lambda: self.browse_file(self.edPath2, which=2))
        self.edCol2 = QLineEdit("A")
        self.edCol2.setMaximumWidth(60)
        self.spinHdr2 = QSpinBox(); self.spinHdr2.setRange(1, 10000); self.spinHdr2.setValue(1)

        row2 = QWidget(); r2 = QHBoxLayout(row2)
        r2.addWidget(self.edPath2, 1)
        r2.addWidget(self.btnBrowse2)
        row2b = QWidget(); r2b = QHBoxLayout(row2b)
        r2b.addWidget(label_caption("Колонка блюд:"))
        r2b.addWidget(self.edCol2)
        r2b.addWidget(label_caption("Строка заголовка:"))
        r2b.addWidget(self.spinHdr2)

        g2 = QWidget(); l2 = QVBoxLayout(g2)
        l2.addWidget(label_caption("Файл 2"))
        l2.addWidget(row2)
        l2.addWidget(row2b)
        self.grpSecond = nice_group("Второй файл", g2)
        self.contentLayout.addWidget(self.grpSecond)
        self.grpSecond.setVisible(False)

        # Параметры (дополнительно) — сворачиваемая группа
        opts = QWidget(); lo = QHBoxLayout(opts)
        self.chkIgnoreCase = QCheckBox("Игнорировать регистр")
        self.chkIgnoreCase.setChecked(True)
        self.chkFuzzy = QCheckBox("Похожесть")
        self.spinFuzzy = QSpinBox(); self.spinFuzzy.setRange(0, 100); self.spinFuzzy.setValue(85)
        self.spinFuzzy.setEnabled(False)
        self.chkFuzzy.toggled.connect(self.spinFuzzy.setEnabled)

        self.rbAuto = QRadioButton("Итоговый: авто (по дате)")
        self.rbAuto.setChecked(True)
        self.finalGroup = QButtonGroup(self)
        self.finalGroup.addButton(self.rbAuto)

        lo.addWidget(self.chkIgnoreCase)
        lo.addWidget(self.chkFuzzy)
        lo.addWidget(QLabel("Порог:"))
        lo.addWidget(self.spinFuzzy)
        lo.addStretch(1)
        lo.addWidget(self.rbAuto)

        self.paramsBox = QGroupBox("Параметры (дополнительно)")
        self.paramsBox.setObjectName("paramsBox")
        self.paramsBox.setCheckable(True)
        self.paramsBox.setChecked(False)
        lparams = QVBoxLayout(self.paramsBox)
        lparams.addWidget(opts)
        opts.setVisible(False)
        self.paramsBox.toggled.connect(opts.setVisible)
        self.contentLayout.addWidget(self.paramsBox)
        self.paramsBox.setVisible(False)

        # Действия (показываются после нажатия "Сравнить меню")
        self.actionsPanel = QWidget(); la = QHBoxLayout(self.actionsPanel)
        self.btnCompare = QPushButton("Сравнить и подсветить")
        self.btnCompare.clicked.connect(self.do_compare)

        la.addStretch(1)
        la.addWidget(self.btnCompare)
        self.contentLayout.addWidget(self.actionsPanel)
        self.actionsPanel.setVisible(False)


        # Theming (follow system by default)
        self._theme_mode = ThemeMode.SYSTEM
        apply_theme(QApplication.instance(), self._theme_mode)
        # Watch for system theme changes and apply automatically when in SYSTEM mode
        self._theme_timer = start_system_theme_watcher(lambda light: self._on_system_theme_change(light))

    def log(self, msg: str):
        # Лог отключён по запросу — ничего не делаем
        pass

    def on_theme_changed(self, idx: int):
        if idx == 0:
            self._theme_mode = ThemeMode.SYSTEM
        elif idx == 1:
            self._theme_mode = ThemeMode.LIGHT
        else:
            self._theme_mode = ThemeMode.DARK
        apply_theme(QApplication.instance(), self._theme_mode)

    def _on_system_theme_change(self, light: bool):
        # React only if following system
        if getattr(self, "_theme_mode", ThemeMode.SYSTEM) == ThemeMode.SYSTEM:
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
                    col1=self.edCol1.text().strip() or "A",
                    col2=self.edCol2.text().strip() or "A",
                    header_row1=self.spinHdr1.value(),
                    header_row2=self.spinHdr2.value(),
                    ignore_case=self.chkIgnoreCase.isChecked(),
                    use_fuzzy=self.chkFuzzy.isChecked(),
                    fuzzy_threshold=int(self.spinFuzzy.value()),
                    final_choice=0,
                )
                self.log(f"Готово. Совпадений: {matches}. Итоговый файл: {out_path}")
                QMessageBox.information(self, "Готово", f"Совпадений: {matches}\nИтоговый файл: {out_path}")
            except ColumnParseError as e:
                QMessageBox.warning(self, "Колонка", str(e))
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def show_compare_sections(self):
        try:
            # Показать формы сравнения и панель действий
            if hasattr(self, "grpFirst"):
                self.grpFirst.setVisible(True)
            if hasattr(self, "grpSecond"):
                self.grpSecond.setVisible(True)
            if hasattr(self, "paramsBox"):
                self.paramsBox.setVisible(True)
            if hasattr(self, "actionsPanel"):
                self.actionsPanel.setVisible(True)
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
        try:
            tpl = find_template("presentation_template.pptx")
            if not tpl:
                QMessageBox.information(self, "Презентация", "Шаблон презентации не найден. Положите файл presentation_template.pptx в папку templates.")
                return
            suggested = str(Path.home() / "презентация.pptx")
            out_path, _ = QFileDialog.getSaveFileName(self, "Сохранить презентацию", suggested, "PowerPoint (*.pptx)")
            if not out_path:
                return
            shutil.copy2(tpl, out_path)
            QMessageBox.information(self, "Готово", f"Презентация сохранена:\n{out_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

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

