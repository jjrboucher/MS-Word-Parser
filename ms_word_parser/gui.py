import os
import sys
import logging
import subprocess
from datetime import datetime as dt, timedelta
from pathlib import Path

from PyQt6.QtCore import (
    QCoreApplication,
    QMetaObject,
    QObject,
    QRect,
    Qt,
    QThread,
    QUrl,
    pyqtSignal,
)
from PyQt6.QtGui import (
    QAction,
    QActionGroup,
    QColor,
    QDesktopServices,
    QFont,
)
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QGroupBox,
    QLabel,
    QMainWindow,
    QMenu,
    QMenuBar,
    QMessageBox,
    QFileDialog,
    QGridLayout,
    QPlainTextEdit,
    QPushButton,
    QRadioButton,
    QStyle,
    QStyleFactory,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

try:
    from classes.datastore import DataStore
    from tips import (
        tip_scriptOverview,
        tip_excelWorksheets,
        tip_processingOptions,
        tip_guiWorkFlow,
    )
    from parse_docx import (
        __version__,
        __appname__,
        __source__,
        __date__,
        __author__,
        __dtfmt__,
        write_to_excel,
        write_to_sqlite,
        run_parallel,
        get_files,
    )
except ModuleNotFoundError:
    from ms_word_parser.classes.datastore import DataStore
    from ms_word_parser.tips import (
        tip_scriptOverview,
        tip_excelWorksheets,
        tip_processingOptions,
        tip_guiWorkFlow,
    )
    from ms_word_parser.parse_docx import (
        __version__,
        __appname__,
        __source__,
        __date__,
        __author__,
        __dtfmt__,
        write_to_excel,
        write_to_sqlite,
        run_parallel,
        get_files,
    )

green = QColor(86, 208, 50)
red = QColor(204, 0, 0)
black = QColor(0, 0, 0)


class AboutWindow(QWidget):
    """Sets the structure for the About window"""

    __slots__ = ("text_font", "aboutLabel", "urlLabel", "logoLabel")

    def __init__(self):
        super().__init__()
        layout = QGridLayout()
        self.text_font = QFont()
        if os.sys.platform == "win32":
            self.text_font.setPointSize(9)
        elif os.sys.platform == "linux":
            self.text_font.setPointSize(8)
        elif os.sys.platform == "darwin":
            self.text_font.setPointSize(12)
        self.aboutLabel = QLabel()
        self.urlLabel = QLabel()
        self.logoLabel = QLabel()
        spacer = QLabel()
        layout.addWidget(self.aboutLabel, 0, 0)
        layout.addWidget(spacer, 0, 1)
        layout.addWidget(self.urlLabel, 1, 0)
        layout.addWidget(self.logoLabel, 0, 2)
        self.setStyleSheet("background-color: white; color: black;")
        self.setFixedSize(350, 140)
        screen = QApplication.primaryScreen()
        screen_geometry = screen.geometry()
        x = (screen_geometry.width() - self.width()) // 2
        y = (screen_geometry.height() - self.height()) // 2
        self.move(x, y)
        self.setLayout(layout)
        style = self.style()
        dialog_icon = style.standardIcon(
            QStyle.StandardPixmap.SP_FileDialogDetailedView
        )
        self.setWindowIcon(dialog_icon)


class ContentsWindow(QWidget):
    """Sets the structure for the Contents window"""

    __slots__ = ("text_font", "text_edit")

    def __init__(self):
        super().__init__()
        self.text_font = QFont()
        if os.sys.platform == "win32":
            self.text_font.setPointSize(9)
        elif os.sys.platform == "linux":
            self.text_font.setPointSize(8)
        elif os.sys.platform == "darwin":
            self.text_font.setPointSize(12)
        self.setWindowTitle("Contents")
        self.setFixedSize(700, 800)
        window_text = (
            f"{tip_scriptOverview['Title']}\n{tip_scriptOverview['Text']}\n"
            f"{tip_excelWorksheets['Title']}\n{tip_excelWorksheets['Text']}\n"
            f"{tip_processingOptions['Title']}\n{tip_processingOptions['Text']}\n"
            f"{tip_guiWorkFlow['Title']}\n{tip_guiWorkFlow['Text']}"
        )
        self.text_edit = QPlainTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setPlainText(window_text)
        self.text_edit.setFont(self.text_font)
        self.text_edit.setStyleSheet("padding: 0px;")
        layout = QVBoxLayout()
        layout.addWidget(self.text_edit)
        screen_geometry = QApplication.primaryScreen().availableGeometry()
        x = (screen_geometry.width() - self.width()) // 2
        y = (screen_geometry.height() - self.height()) // 2
        self.move(x, y)
        self.setLayout(layout)
        style = self.style()
        dialog_icon = style.standardIcon(
            QStyle.StandardPixmap.SP_FileDialogDetailedView
        )
        self.setWindowIcon(dialog_icon)


class _WriteLogProxy:

    def __init__(self, signal):
        self._signal = signal

    def update_status(self, msg, level="info", color=None):
        self._signal.emit(msg, level, color)


class ExcelSqliteWriter(QObject):
    """
    Puts write_to_* on a background thread to not lock the GUI.
    """

    log = pyqtSignal(str, str, object)
    finished = pyqtSignal(str)

    def __init__(self, store):
        super().__init__()
        self.store = store

    def run(self):
        real_gui = self.store.ms_word_gui
        self.store.ms_word_gui = _WriteLogProxy(self.log)
        try:
            if self.store.excel:
                write_to_excel(
                    self.store.excel_file, self.store.triage_files, store=self.store
                )
            if self.store.sqlite:
                write_to_sqlite(self.store)
            self.finished.emit("")
        except Exception as e:
            self.finished.emit(str(e))
        finally:
            self.store.ms_word_gui = real_gui


class UiMainWindow:

    def __init__(self, store: DataStore):
        super().__init__()
        self.store = store
        self.d_width = 1142
        self.d_height = 350
        self.files = []
        self.excel_path = ""
        self.excel_full_path = ""
        self.log_path = ""
        self.log_handler = None
        self.can_process = False
        self.logger = logging.getLogger("ms-word-parser")
        self.logger.setLevel(logging.DEBUG)
        self.log_fmt = logging.Formatter(
            "%(asctime)s | %(levelname)-8s | %(message)s",
            datefmt=__dtfmt__,
        )
        self.text_font = QFont()
        if os.sys.platform == "win32":
            self.text_font.setPointSize(9)
        elif os.sys.platform == "linux":
            self.text_font.setPointSize(8)
        elif os.sys.platform == "darwin":
            self.text_font.setPointSize(12)

        self.running = False

    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName("MainWindow")
        MainWindow.resize(self.d_width, self.d_height)
        MainWindow.setFixedWidth(self.d_width)
        MainWindow.setFixedHeight(self.d_height)
        MainWindow.setStyleSheet(self.stylesheet)
        screen_geometry = QApplication.primaryScreen().availableGeometry()
        x = (screen_geometry.width() - self.width()) // 2
        y = (screen_geometry.height() - self.height()) // 2
        self.move(x, y)

        # Menu Actions
        self.actionSelect_Output = QAction(MainWindow)
        self.actionSelect_Output.setObjectName("actionSelect_Output")
        self.actionSelect_Output.triggered.connect(self.select_output)
        self.actionAdd_Ingest = QAction(MainWindow)
        self.actionAdd_Ingest.setObjectName("actionAdd_Ingest")
        self.actionAdd_Ingest.triggered.connect(self.add_ingest)
        self.actionAdd_Ingest.setVisible(False)
        self.actionAdd_Files = QAction(MainWindow)
        self.actionAdd_Files.setObjectName("actionAdd_Files")
        self.actionAdd_Files.triggered.connect(self.add_files)
        self.actionAdd_Files.setVisible(False)
        self.actionAdd_Directory = QAction(MainWindow)
        self.actionAdd_Directory.setObjectName("actionAdd_Directory")
        self.actionAdd_Directory.triggered.connect(self.add_directory)
        self.actionAdd_Directory.setVisible(False)
        self.actionExit = QAction(MainWindow)
        self.actionExit.setObjectName("actionExit")
        self.actionExit.triggered.connect(self.close)
        self.actionAbout = QAction(MainWindow)
        self.actionAbout.setObjectName("actionAbout")
        self.actionAbout.triggered.connect(self._about)
        self.actionContents = QAction(MainWindow)
        self.actionContents.setObjectName("actionContents")
        self.actionContents.triggered.connect(self._contents)

        # Central Widget
        self.centralWidget = QWidget(MainWindow)
        self.centralWidget.setObjectName("centralWidget")

        # Processing Options
        self.processOptions = QGroupBox(self.centralWidget)
        self.processOptions.setObjectName("processOptions")
        self.processOptions.setGeometry(QRect(10, 10, 340, 100))
        self.processOptions.setStyleSheet(self.panel_style)
        self.processOptions.setFont(self.text_font)
        self.triageButton = QRadioButton(self.processOptions)
        self.triageButton.setObjectName("triageButton")
        self.triageButton.setGeometry(QRect(10, 30, 89, 20))
        self.triageButton.setStyleSheet(self.stylesheet)
        self.triageButton.setChecked(True)
        self.triageButton.setFont(self.text_font)
        self.fullButton = QRadioButton(self.processOptions)
        self.fullButton.setObjectName("fullButton")
        self.fullButton.setGeometry(QRect(10, 60, 60, 20))
        self.fullButton.setStyleSheet(self.stylesheet)
        self.fullButton.setFont(self.text_font)
        self.excelCheck = QCheckBox(self.processOptions)
        self.excelCheck.setObjectName("excelCheck")
        self.excelCheck.setGeometry(QRect(100, 30, 89, 20))
        self.excelCheck.setStyleSheet(self.stylesheet)
        self.excelCheck.setChecked(True)
        self.excelCheck.setFont(self.text_font)
        self.excelCheck.stateChanged.connect(lambda: self.toggle_process())
        self.sqliteButton = QCheckBox(self.processOptions)
        self.sqliteButton.setObjectName("sqliteButton")
        self.sqliteButton.setGeometry(QRect(190, 30, 89, 20))
        self.sqliteButton.setStyleSheet(self.stylesheet)
        self.sqliteButton.setFont(self.text_font)
        self.sqliteButton.stateChanged.connect(lambda: self.toggle_process())
        self.hashFiles = QCheckBox(self.processOptions)
        self.hashFiles.setObjectName("hashFiles")
        self.hashFiles.setGeometry(QRect(100, 60, 89, 20))
        self.hashFiles.setStyleSheet(self.stylesheet)
        self.hashFiles.setFont(self.text_font)
        self.timelineButton = QCheckBox(self.processOptions)
        self.timelineButton.setObjectName("timelineButton")
        self.timelineButton.setGeometry(QRect(190, 60, 89, 20))
        self.timelineButton.setStyleSheet(self.stylesheet)
        self.timelineButton.setFont(self.text_font)

        # Operation Options
        self.operationOptions = QGroupBox(self.centralWidget)
        self.operationOptions.setObjectName("operationOptions")
        self.operationOptions.setGeometry(QRect(10, 116, 340, 100))
        self.operationOptions.setStyleSheet(self.panel_style)
        self.operationOptions.setFont(self.text_font)
        self.outputButton = QPushButton(self.operationOptions)
        self.outputButton.setObjectName("outputButton")
        self.outputButton.setGeometry(QRect(10, 28, 86, 24))
        self.outputButton.setStyleSheet(self.stylesheet)
        self.outputButton.clicked.connect(self.select_output)
        self.outputButton.setFont(self.text_font)
        self.addFilesButton = QPushButton(self.operationOptions)
        self.addFilesButton.setObjectName("addFilesButton")
        self.addFilesButton.setGeometry(QRect(112, 28, 86, 24))
        self.addFilesButton.setEnabled(False)
        self.addFilesButton.setStyleSheet(self.disabled)
        self.addFilesButton.clicked.connect(self.add_files)
        self.addFilesButton.setFont(self.text_font)
        self.addDirectoryButton = QPushButton(self.operationOptions)
        self.addDirectoryButton.setObjectName("addDirectoryButton")
        self.addDirectoryButton.setGeometry(QRect(214, 28, 86, 24))
        self.addDirectoryButton.setEnabled(False)
        self.addDirectoryButton.setStyleSheet(self.disabled)
        self.addDirectoryButton.clicked.connect(self.add_directory)
        self.addDirectoryButton.setFont(self.text_font)
        self.addIngestButton = QPushButton(self.operationOptions)
        self.addIngestButton.setObjectName("addIngestButton")
        self.addIngestButton.setGeometry(QRect(10, 58, 86, 24))
        self.addIngestButton.setEnabled(False)
        self.addIngestButton.setStyleSheet(self.disabled)
        self.addIngestButton.setFont(self.text_font)
        self.addIngestButton.clicked.connect(self.add_ingest)
        self.processButton = QPushButton(self.operationOptions)
        self.processButton.setObjectName("processButton")
        self.processButton.setGeometry(QRect(112, 58, 86, 24))
        self.processButton.setEnabled(False)
        self.processButton.setStyleSheet(self.disabled)
        self.processButton.clicked.connect(
            lambda: self.analyze_docs(
                self.files,
                self.triageButton.isChecked(),
                self.hashFiles.isChecked(),
                self.timelineButton.isChecked(),
                self.excelCheck.isChecked(),
                self.sqliteButton.isChecked(),
            )
        )
        self.processButton.setFont(self.text_font)
        self.resetButton = QPushButton(self.operationOptions)
        self.resetButton.setObjectName("resetButton")
        self.resetButton.setGeometry(QRect(214, 58, 86, 24))
        self.resetButton.clicked.connect(self._reset)
        self.resetButton.setStyleSheet(self.stylesheet)
        self.resetButton.setFont(self.text_font)

        # Output Files
        self.outputFiles = QGroupBox(self.centralWidget)
        self.outputFiles.setObjectName("outputFiles")
        self.outputFiles.setGeometry(QRect(10, 220, 340, 90))
        self.outputFiles.setStyleSheet(self.panel_style)
        self.outputFiles.setFont(self.text_font)
        self.outputPathLabel = QLabel(self.outputFiles)
        self.outputPathLabel.setObjectName("outputPathLabel")
        self.outputPathLabel.setGeometry(QRect(10, 30, 80, 16))
        self.outputPathLabel.setStyleSheet(self.panel_style)
        self.outputPathLabel.setFont(self.text_font)
        self.outputPath = QTextEdit(self.outputFiles)
        self.outputPath.setAlignment(
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft
        )
        self.outputPath.setObjectName("outputPath")
        self.outputPath.setGeometry(QRect(90, 26, 240, 26))
        self.outputPath.setStyleSheet(self.panel_style)
        self.outputPath.setReadOnly(True)
        self.outputPath.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.outputPath.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.outputPath.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.outputPath.setFont(self.text_font)
        self.generalLog = QLabel(self.outputFiles)
        self.generalLog.setObjectName("generalLog")
        self.generalLog.setGeometry(QRect(10, 61, 80, 16))
        self.generalLog.setStyleSheet(self.panel_style)
        self.generalLog.setFont(self.text_font)
        self.generalLogFile = QTextEdit(self.outputFiles)
        self.generalLogFile.setAlignment(
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft
        )
        self.generalLogFile.setObjectName("generalLogFile")
        self.generalLogFile.setGeometry(QRect(90, 58, 240, 26))
        self.generalLogFile.setStyleSheet(self.panel_style)
        self.generalLogFile.setReadOnly(True)
        self.generalLogFile.setFont(self.text_font)
        self.generalLogFile.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.generalLogFile.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.generalLogFile.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)

        # Processing Status
        self.processStatus = QGroupBox(self.centralWidget)
        self.processStatus.setObjectName("processStatus")
        self.processStatus.setGeometry(QRect(360, 10, 768, 300))
        self.processStatus.setStyleSheet(self.panel_style)
        self.processStatus.setFont(self.text_font)
        self.docxOutput = QTextEdit(self.processStatus)
        self.docxOutput.setObjectName("docxOutput")
        self.docxOutput.setGeometry(QRect(16, 60, 737, 230))
        self.docxOutput.setStyleSheet(self.scrollbar_sheet)
        self.docxOutput.setReadOnly(True)
        self.docxOutput.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )
        self.docxOutput.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.docxOutput.setFont(self.text_font)
        self.numOfFilesLabel = QLabel(self.processStatus)
        self.numOfFilesLabel.setObjectName("numOfFilesLabel")
        self.numOfFilesLabel.setGeometry(QRect(18, 28, 120, 26))
        self.numOfFilesLabel.setStyleSheet(self.panel_style)
        self.numOfFilesLabel.setFont(self.text_font)
        self.numOfFiles = QTextEdit(self.processStatus)
        self.numOfFiles.setObjectName("numOfFiles")
        self.numOfFiles.setGeometry(QRect(85, 28, 40, 26))
        self.numOfFiles.setAlignment(
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft
        )
        self.numOfFiles.setReadOnly(True)
        self.numOfFiles.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.numOfFiles.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.numOfFiles.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.numOfFiles.setFont(self.text_font)
        self.numOfFiles.setStyleSheet(self.panel_style)
        self.numOfErrorsLabel = QLabel(self.processStatus)
        self.numOfErrorsLabel.setObjectName("numOfErrorsLabel")
        self.numOfErrorsLabel.setGeometry(QRect(135, 28, 80, 26))
        self.numOfErrorsLabel.setStyleSheet(self.panel_style)
        self.numOfErrorsLabel.setFont(self.text_font)
        self.numOfErrors = QTextEdit(self.processStatus)
        self.numOfErrors.setObjectName("numOfErrors")
        self.numOfErrors.setGeometry(QRect(207, 28, 40, 26))
        self.numOfErrors.setAlignment(
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft
        )
        self.numOfErrors.setReadOnly(True)
        self.numOfErrors.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.numOfErrors.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.numOfErrors.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.numOfErrors.setFont(self.text_font)
        self.numOfErrors.setStyleSheet(self.panel_style)
        self.numRemainingLabel = QLabel(self.processStatus)
        self.numRemainingLabel.setObjectName("numRemainingLabel")
        self.numRemainingLabel.setGeometry(QRect(257, 28, 120, 26))
        self.numRemainingLabel.setStyleSheet(self.panel_style)
        self.numRemainingLabel.setFont(self.text_font)
        self.numRemaining = QTextEdit(self.processStatus)
        self.numRemaining.setObjectName("numRemaining")
        self.numRemaining.setGeometry(QRect(339, 28, 40, 26))
        self.numRemaining.setAlignment(
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft
        )
        self.numRemaining.setReadOnly(True)
        self.numRemaining.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.numRemaining.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.numRemaining.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.numRemaining.setFont(self.text_font)
        self.numRemaining.setStyleSheet(self.panel_style)
        self.openLogButton = QPushButton(self.processStatus)
        self.openLogButton.setObjectName("openLogButton")
        self.openLogButton.setGeometry(QRect(522, 29, 110, 24))
        self.openLogButton.setFont(self.text_font)
        self.openLogButton.setStyleSheet(self.disabled)
        self.openLogButton.setEnabled(False)
        self.openLogButton.clicked.connect(lambda: self.open_file(self.log_path))
        self.stopButton = QPushButton(self.processStatus)
        self.stopButton.setObjectName("stopButton")
        self.stopButton.setGeometry(QRect(402, 29, 110, 24))
        self.stopButton.setEnabled(False)
        self.stopButton.setStyleSheet(self.disabled)
        self.stopButton.clicked.connect(self._stop)
        self.stopButton.setFont(self.text_font)
        self.openButton = QPushButton(self.processStatus)
        self.openButton.setObjectName("openButton")
        self.openButton.setGeometry(QRect(642, 29, 110, 24))
        self.openButton.setFont(self.text_font)
        self.openButton.setStyleSheet(self.disabled)
        self.openButton.setEnabled(False)
        self.openButton.clicked.connect(self.open_path)
        MainWindow.setCentralWidget(self.centralWidget)
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setObjectName("menubar")
        self.menubar.setGeometry(QRect(0, 0, 1192, 22))
        self.menubar.setFont(self.text_font)
        self.menuFile = QMenu(self.menubar)
        self.menuFile.setObjectName("menuFile")
        self.menuFile.setFont(self.text_font)
        self.menuStyle = QMenu(self.menubar)
        self.menuStyle.setObjectName("menuStyle")
        self.menuStyle.setFont(self.text_font)
        self.menuHelp = QMenu(self.menubar)
        self.menuHelp.setObjectName("menuHelp")
        self.menuHelp.setFont(self.text_font)
        MainWindow.setMenuBar(self.menubar)

        # Menu Bar
        self.menubar.addAction(self.menuFile.menuAction())
        self.menubar.addAction(self.menuStyle.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())
        self.menuFile.addAction(self.actionSelect_Output)
        self.menuFile.addSeparator()
        self.menuFile.addAction(self.actionAdd_Ingest)
        self.menuFile.addAction(self.actionAdd_Files)
        self.menuFile.addAction(self.actionAdd_Directory)
        self.menuFile.addSeparator()
        self.menuFile.addAction(self.actionExit)
        current_style = QApplication.instance().style().objectName().lower()
        self.styleActionGroup = QActionGroup(MainWindow)
        self.styleActionGroup.setExclusive(True)
        for style_name in sorted(QStyleFactory.keys()):
            action = QAction(style_name, MainWindow)
            action.setCheckable(True)
            action.setChecked(style_name.lower() == current_style)
            action.triggered.connect(
                lambda checked, name=style_name: self._set_widget_style(name)
            )
            self.styleActionGroup.addAction(action)
            self.menuStyle.addAction(action)
        self.menuHelp.addAction(self.actionContents)
        self.menuHelp.addSeparator()
        self.menuHelp.addAction(self.actionAbout)
        self.retranslateUi(MainWindow)

        QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(
            QCoreApplication.translate("MainWindow", __appname__, None)
        )
        self.actionSelect_Output.setText(
            QCoreApplication.translate("MainWindow", "Select &Output Path ...", None)
        )
        self.actionAdd_Ingest.setText(
            QCoreApplication.translate("MainWindow", "Add &Ingest File ...", None)
        )
        self.actionAdd_Files.setText(
            QCoreApplication.translate("MainWindow", "Add &Files ...", None)
        )
        self.actionAdd_Directory.setText(
            QCoreApplication.translate("MainWindow", "Add &Directory ...", None)
        )
        self.actionExit.setText(QCoreApplication.translate("MainWindow", "&Exit", None))
        self.actionAbout.setText(
            QCoreApplication.translate("MainWindow", "&About", None)
        )
        self.actionContents.setText(
            QCoreApplication.translate("MainWindow", "Contents", None)
        )
        self.processOptions.setTitle(
            QCoreApplication.translate("MainWindow", "Processing Options", None)
        )
        self.triageButton.setText(
            QCoreApplication.translate("MainWindow", "Triage", None)
        )
        self.fullButton.setText(QCoreApplication.translate("MainWindow", "Full", None))
        self.excelCheck.setText(
            QCoreApplication.translate("MainWindow", "To Excel", None)
        )
        self.hashFiles.setText(
            QCoreApplication.translate("MainWindow", "Hash Files", None)
        )
        self.sqliteButton.setText(
            QCoreApplication.translate("MainWindow", "To SQLite", None)
        )
        self.timelineButton.setText(
            QCoreApplication.translate("MainWindow", "Timeline", None)
        )
        self.outputFiles.setTitle(
            QCoreApplication.translate("MainWindow", "Output Files", None)
        )
        self.outputPathLabel.setText(
            QCoreApplication.translate("MainWindow", "Output Path:", None)
        )
        self.processStatus.setTitle(
            QCoreApplication.translate("MainWindow", "Processing Status", None)
        )
        self.processButton.setText(
            QCoreApplication.translate("MainWindow", "Process", None)
        )
        self.stopButton.setText(QCoreApplication.translate("MainWindow", "Stop", None))
        self.addIngestButton.setText(
            QCoreApplication.translate("MainWindow", "Add Ingest", None)
        )
        self.resetButton.setText(
            QCoreApplication.translate("MainWindow", "Reset", None)
        )
        self.outputButton.setText(
            QCoreApplication.translate("MainWindow", "Output Path", None)
        )
        self.addFilesButton.setText(
            QCoreApplication.translate("MainWindow", "Add Files", None)
        )
        self.addDirectoryButton.setText(
            QCoreApplication.translate("MainWindow", "Add Directory", None)
        )
        self.openLogButton.setText(
            QCoreApplication.translate("MainWindow", "Open Log File", None)
        )
        self.openButton.setText(
            QCoreApplication.translate("MainWindow", "Open Output Path", None)
        )
        self.numOfFilesLabel.setText(
            QCoreApplication.translate("MainWindow", "# of Files", None)
        )
        self.numOfFiles.setText(QCoreApplication.translate("MainWindow", "0", None))
        self.numOfErrorsLabel.setText(
            QCoreApplication.translate("MainWindow", "# of Errors", None)
        )
        self.numOfErrors.setText(QCoreApplication.translate("MainWindow", "0", None))
        self.numRemainingLabel.setText(
            QCoreApplication.translate("MainWindow", "# Remaining", None)
        )
        self.numRemaining.setText(QCoreApplication.translate("MainWindow", "0", None))
        self.generalLog.setText(
            QCoreApplication.translate("MainWindow", "Log File:", None)
        )
        self.generalLogFile.setText(
            QCoreApplication.translate("MainWindow", self.store.log_file, None)
        )
        self.operationOptions.setTitle(
            QCoreApplication.translate("MainWindow", "Operation Options", None)
        )
        self.menuFile.setTitle(QCoreApplication.translate("MainWindow", "File", None))
        self.menuStyle.setTitle(QCoreApplication.translate("MainWindow", "Style", None))
        self.menuHelp.setTitle(QCoreApplication.translate("MainWindow", "Help", None))

    def select_output(self):
        update_status = self.update_status
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Select a directory for output ...",
            "",
            QFileDialog.Option.ShowDirsOnly,
        )
        if folder_path:
            folder_path = os.path.normpath(folder_path)
            self.store.output_path = folder_path
            self.excel_path = self.store.output_path
            self.log_path = os.path.normpath(
                f"{self.excel_path}{os.sep}{self.store.log_file}"
            )
            self.log_handler = logging.FileHandler(self.log_path, "w", "utf-8")
            self.log_handler.setFormatter(self.log_fmt)
            self.logger.addHandler(self.log_handler)
            update_status = self.update_status
            update_status(f"{__appname__}")
            excel_full_path = f"{folder_path}{os.sep}{self.store.basename}.xlsx"
            excel_full_path = os.path.normpath(excel_full_path)
            self.excel_full_path = excel_full_path
            self.store.excel_file = excel_full_path
            sqlite_full_path = os.path.normpath(
                f"{folder_path}{os.sep}{self.store.basename}.db"
            )
            self.store.sqlite_file = sqlite_full_path
            update_status(f"Output File Path: {folder_path}")
            update_status(f"Log file: {self.log_path}")
            if self.numOfFiles.toPlainText() != "0":
                self.processButton.setEnabled(True)
                self.processButton.setStyleSheet(self.stylesheet)
            self.actionAdd_Files.setVisible(True)
            self.actionAdd_Directory.setVisible(True)
            self.actionAdd_Ingest.setVisible(True)
            self.generalLogFile.setText(self.store.log_file)
            self.outputPath.setText(folder_path)
            self.openButton.setEnabled(True)
            self.openButton.setStyleSheet(self.stylesheet)
            self.addFilesButton.setEnabled(True)
            self.addFilesButton.setStyleSheet(self.stylesheet)
            self.addDirectoryButton.setEnabled(True)
            self.addDirectoryButton.setStyleSheet(self.stylesheet)
            self.addIngestButton.setEnabled(True)
            self.addIngestButton.setStyleSheet(self.stylesheet)

    def add_ingest(self):
        update_status = self.update_status
        all_files = []
        exists = []
        no_files = []
        file, _ = QFileDialog.getOpenFileName(
            self,
            "Select ingestion file ...",
            "",
            "txt Files (*.txt)",
        )
        if file:
            with open(file, "r", encoding="utf-8-sig") as content:
                ingest_data = content.readlines()
                for line in ingest_data:
                    all_files.append(os.path.normpath(line.strip()))
                for f in all_files:
                    if os.path.exists(f):
                        exists.append(f)
                    else:
                        no_files.append(f)
                all_files = exists
                self.numOfFiles.setText(str(len(all_files)))
                self.numRemaining.setText(str(len(all_files)))
                if no_files:
                    update_status(
                        f"The following {len(no_files)} file(s) do not exist:",
                        color=red,
                    )
                    joiner = f"\n{dt.now().strftime(__dtfmt__)} -     "
                    update_status("    " + joiner.join(no_files), color=red)
                if len(all_files) > 1:
                    update_status(
                        f"The following {len(all_files)} files have been loaded:"
                    )
                else:
                    update_status(
                        f"The following {len(all_files)} file has been loaded:"
                    )
                joiner = f"\n{dt.now().strftime(__dtfmt__)} -     "
                update_status("    " + joiner.join(all_files))
                self.files = all_files
                self.can_process = True
                self.toggle_process()

    def add_directory(self):
        update_status = self.update_status
        folder_path = QFileDialog.getExistingDirectory(
            self, "Select a directory ...", "", QFileDialog.Option.ShowDirsOnly
        )
        if folder_path:
            folder_path = Path(folder_path)
            app = QApplication.instance()
            style = app.style()
            msg_box = QMessageBox(None)
            msg_box.setIcon(QMessageBox.Icon.Question)
            dialog_icon = style.standardIcon(
                QStyle.StandardPixmap.SP_FileDialogDetailedView
            )
            msg_box.setWindowIcon(dialog_icon)
            msg_box.setWindowTitle("Load recursively?")
            msg_box.setText(
                "Do you want to recursively load all files in this directory?"
            )
            msg_box.setStandardButtons(
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            msg_box.setDefaultButton(QMessageBox.StandardButton.Yes)
            response = msg_box.exec()
            if response == QMessageBox.StandardButton.Yes:
                recursive_list = get_files(folder_path, True)
                files = [str(file) for file in recursive_list]
            else:
                non_recursive_list = get_files(folder_path, False)
                files = [str(file) for file in non_recursive_list]
            self.numOfFiles.setText(str(len(files)))
            self.numRemaining.setText(str(len(files)))
            if files:
                if len(files) > 1:
                    update_status(f"The following {len(files)} files have been loaded:")
                else:
                    update_status(f"The following {len(files)} file has been loaded:")
                joiner = f"\n{dt.now().strftime(__dtfmt__)} -     "
                update_status("    " + joiner.join(files))
                self.files = files
                self.can_process = True
                self.toggle_process()
            else:
                update_status("No files found. Please check your path and try again.")

    def add_files(self):
        update_status = self.update_status
        all_files = []
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select files ...",
            "",
            "docx, dotx, dotm, docm Files (*.docx *.dotx *.dotm *.docm)",
        )
        if files:
            for file in files:
                all_files.append(os.path.normpath(file))
            self.numOfFiles.setText(str(len(all_files)))
            self.numRemaining.setText(str(len(all_files)))
            if len(all_files) > 1:
                update_status(f"The following {len(all_files)} files have been loaded:")
            else:
                update_status(f"The following {len(all_files)} file has been loaded:")
            joiner = f"\n{dt.now().strftime(__dtfmt__)} -     "
            update_status("    " + joiner.join(all_files))
            self.files = all_files
            self.can_process = True
            self.toggle_process()

    def open_path(self):
        out_path = self.outputPath.toPlainText().strip()
        if out_path:
            QDesktopServices.openUrl(QUrl.fromLocalFile(out_path))

    def open_file(self, file):
        this_os = sys.platform
        cmd = {
            "win32": "start",
            "darwin": "open",
            "linux": "xdg-open",
        }
        launch = cmd[this_os]
        try:
            if this_os == "win32":
                os.startfile(file)
            else:
                subprocess.Popen([launch, file], start_new_session=True)
        except Exception as e:
            self.update_status(f"Unable to open {file}: {e}", level="error")

    def toggle_process(self):
        if (
            self.excelCheck.isChecked() or self.sqliteButton.isChecked()
        ) and self.can_process:
            self.processButton.setEnabled(True)
            self.processButton.setStyleSheet(self.stylesheet)
        else:
            self.processButton.setEnabled(False)
            self.processButton.setStyleSheet(self.disabled)

    def _reset(self):
        self.store.reset_vars()
        self.can_process = False
        self.generalLogFile.setText(self.store.log_file)
        self.outputPath.clear()
        self.numOfFiles.setText("0")
        self.numOfErrors.setText("0")
        self.numRemaining.setText("0")
        self.docxOutput.setTextColor(self.default_text_color)
        self.docxOutput.clear()
        self.processButton.setEnabled(False)
        self.processButton.setStyleSheet(self.disabled)
        self.openLogButton.setEnabled(False)
        self.openLogButton.setStyleSheet(self.disabled)
        self.openButton.setEnabled(False)
        self.openButton.setStyleSheet(self.disabled)
        self.actionAdd_Files.setVisible(False)
        self.actionAdd_Directory.setVisible(False)
        self.actionAdd_Ingest.setVisible(False)
        self.triageButton.setChecked(True)
        self.addFilesButton.setEnabled(False)
        self.addFilesButton.setStyleSheet(self.disabled)
        self.addDirectoryButton.setEnabled(False)
        self.addDirectoryButton.setStyleSheet(self.disabled)
        self.addIngestButton.setEnabled(False)
        self.addIngestButton.setStyleSheet(self.disabled)
        self.hashFiles.setChecked(False)
        self.sqliteButton.setChecked(False)
        self.timelineButton.setChecked(False)
        self.stopButton.setEnabled(False)
        self.stopButton.setStyleSheet(self.disabled)
        self.excelCheck.setChecked(True)

    def _stop(self):
        self.running = False
        self.stopButton.setStyleSheet(self.disabled)
        self.stopButton.setEnabled(False)
        self.addFilesButton.setEnabled(False)
        self.addFilesButton.setStyleSheet(self.disabled)
        self.addDirectoryButton.setEnabled(False)
        self.addDirectoryButton.setStyleSheet(self.disabled)

    def _about(self):
        self.aboutWindow = AboutWindow()
        self.aboutWindow.setWindowFlags(
            self.aboutWindow.windowFlags() & ~Qt.WindowType.WindowMinMaxButtonsHint
        )
        sqliteUrl = (
            "https://github.com/jjrboucher/MS-Word-Parser/tree/master/sqlite_queries"
        )
        githubLink = f'<a href="{__source__}">View the source on GitHub</a>'
        sqliteLink = f'<a href="{sqliteUrl}">Sample SQLite queries</a>'
        self.aboutWindow.setWindowTitle("About")
        self.aboutWindow.aboutLabel.setText(
            f"Version: {__appname__}\nLast Updated: {__date__}\n\nAuthors:\n{__author__}"
        )
        self.aboutWindow.urlLabel.setOpenExternalLinks(True)
        self.aboutWindow.urlLabel.setText(
            f"{githubLink}&nbsp;&nbsp;&nbsp;&nbsp;{sqliteLink}"
        )
        self.aboutWindow.aboutLabel.setFont(self.text_font)
        self.aboutWindow.urlLabel.setFont(self.text_font)
        self.aboutWindow.show()

    def _contents(self):
        self.contentsWindow = ContentsWindow()
        self.contentsWindow.setWindowFlags(
            self.contentsWindow.windowFlags() & ~Qt.WindowType.WindowMinMaxButtonsHint
        )
        self.contentsWindow.show()

    def update_status(self, msg, level="info", color=None):
        levels = {
            "info": logging.INFO,
            "warning": logging.WARNING,
            "error": logging.ERROR,
            "debug": logging.DEBUG,
        }
        log_level = levels[level]
        if level in {"info", "error"}:
            if self.store.ms_word_gui:
                self.docxOutput.setTextColor(color or self.default_text_color)
                self.docxOutput.append(f"{dt.now().strftime(__dtfmt__)} - {msg}")
                self.docxOutput.setTextColor(self.default_text_color)
        try:
            self.logger.log(log_level, msg)
        except (UnicodeDecodeError, UnicodeEncodeError):
            self.logger.log(log_level, msg.encode("utf-8", errors="surrogatepass"))
        QApplication.processEvents()

    def analyze_docs(self, files, triage_files, hash_files, timeline, excel, sqlite):
        if not self.running:
            self.running = True
        start_time = dt.now().strftime(__dtfmt__)
        self.store.start_time = start_time
        self.store.sqlite = sqlite
        self.store.filenames = files
        self.store.timeline = timeline
        self.store.excel = excel
        self.store.triage_files = triage_files
        self.store.hash_files = hash_files
        self.stopButton.setEnabled(True)
        self.stopButton.setStyleSheet(self.stylesheet)
        self.resetButton.setEnabled(False)
        self.resetButton.setStyleSheet(self.disabled)
        self.processButton.setEnabled(False)
        self.processButton.setStyleSheet(self.disabled)
        update_status = self.update_status
        script_start = dt.now().strftime(__dtfmt__)
        update_status(f"Script executed: {script_start}")
        update_status("Summary of files parsed:")
        update_status(f'{"="*36}')
        remaining = int(self.numRemaining.toPlainText())

        def on_file_done(error_count):
            nonlocal remaining
            self.numOfErrors.setText(str(error_count))
            if remaining != 0:
                remaining -= 1
            self.numRemaining.setText(str(remaining))

        docxErrorCount = run_parallel(
            files,
            triage_files,
            hash_files,
            self.store,
            log_fn=lambda msg, level="info", color=None: update_status(
                msg, level=level, color=color
            ),
            is_stopped=lambda: not self.running,
            on_file_done=on_file_done,
            error_color=red,
        )
        stopped = not self.running
        if stopped:
            update_status("Processing stopped")
            self.stopButton.setEnabled(False)
            self.resetButton.setEnabled(True)
            self.resetButton.setStyleSheet(self.stylesheet)
            update_status("Attempting to write current results to Excel")
        self._start_write_thread(docxErrorCount, stopped)

    def _start_write_thread(self, docxErrorCount, stopped):
        update_status = self.update_status
        self._write_thread = QThread()
        self._write_worker = ExcelSqliteWriter(self.store)
        self._write_worker.moveToThread(self._write_thread)
        self._write_thread.started.connect(self._write_worker.run)
        self._write_worker.log.connect(
            lambda msg, level, color: update_status(
                msg, level=level, **({} if color is None else {"color": color})
            )
        )
        self._write_worker.finished.connect(
            lambda error: self._on_write_finished(error, docxErrorCount, stopped)
        )
        self._write_worker.finished.connect(self._write_thread.quit)
        self._write_worker.finished.connect(self._write_worker.deleteLater)
        self._write_thread.finished.connect(self._write_thread.deleteLater)
        self._write_thread.start()

    def _on_write_finished(self, error, docxErrorCount, stopped):
        update_status = self.update_status
        if error:
            update_status(f"Unable to write results to Excel: {error}")
        else:
            clr = red if docxErrorCount > 0 else self.default_text_color
            if not stopped:
                update_status(f'{"="*24}')
            summary = (
                f"Finished writing to Excel. Errors detected: {docxErrorCount}"
                if stopped
                else f"Processing finished for all files. Errors detected: {docxErrorCount}"
            )
            update_status(summary, color=clr)
            if docxErrorCount > 0:
                update_status("The following files had errors:", "error", color=clr)
                for each_file in self.store.errors_worksheet["File Name"]:
                    update_status(f"  {each_file}", "error", color=clr)
            end_time = dt.now().strftime(__dtfmt__)
            update_status(f"Script finished execution: {end_time}", color=green)
            run_time = str(
                timedelta(
                    seconds=(
                        dt.strptime(end_time, __dtfmt__)
                        - dt.strptime(self.store.start_time, __dtfmt__)
                    ).seconds
                )
            )
            update_status(f"Total processing time: {run_time}", color=green)
            self.openLogButton.setEnabled(True)
            self.openLogButton.setStyleSheet(self.stylesheet)
        if stopped:
            return
        self.resetButton.setEnabled(True)
        self.resetButton.setStyleSheet(self.stylesheet)
        self.stopButton.setEnabled(False)
        self.stopButton.setStyleSheet(self.disabled)
        if not error:
            self.store.reset_vars()


class MsWordGui(QMainWindow, UiMainWindow):
    """MS Word Parser GUI Class"""

    disabled = """
        QPushButton {
            background-color: white; border: 1px solid black; color: grey;
        }
        """

    stylesheet = """
        QMainWindow {
            background-color: white; color: black;
        }
        QLineEdit {
            background-color: white; color: black;
        }
        QDateTimeEdit {
            background-color: white; color: black;
        }
        QCheckBox {
            background: white; color:black;
        }
        QMenu {
            background-color: white; border: 1px solid black; color: black;
        }
        QMenu::item {
            padding: 4px 20px; background-color: transparent; color: black;
        }
        QMenu::item:selected {
            background-color: #d9ebfb; color: black;
        }
        QMenuBar {
            background-color: white; color: black;
        }
        QMenuBar::item {
            background-color: white; color: black;
        }
        QMenuBar::item:selected {
            background-color: #d9ebfb; color: black;
        }
        QPushButton {
            background-color: white; border: 1px solid black; color: black;
        }
        QPushButton:hover {
            background-color: #d9ebfb; border: 1px solid black;
        }
        QRadioButton {
            background: white; color:black;
        }
        """
    scrollbar_sheet = """
        QScrollBar:vertical {
            border: 0px;
            background:white;
            width:7px;
            margin: 0px 0px 0px 0px;
        }
        QScrollBar::handle:vertical {
            min-height: 30px;
            border: 0px;
            border-radius: 3px;
            background-color: #a0a0a0;
        }
        QScrollBar::handle:vertical:hover {
            background: #808080;
        }
        QScrollBar::add-line:vertical {
            height: 0px;
            subcontrol-position: bottom;
            subcontrol-origin: margin;
        }
        QScrollBar::sub-line:vertical {
            height: 0 px;
            subcontrol-position: top;
            subcontrol-origin: margin;
        }
        QScrollBar:horizontal {
            border: 0px;
            background: white;
            height: 7px;
            margin: 0px 0px 0px 0px;
        }
        QScrollBar::handle:horizontal {
            background-color: #a0a0a0;
            min-width: 5px;
            border: 0px;
            border-radius: 3px;
        }
        QScrollBar::handle:horizontal:hover {
            background: #808080;
        }
        QScrollBar::sub-line:horizontal, QScrollBar::add-line:horizontal {
            background: none;
            border: none;
            width: 7px;
            subcontrol-origin: margin;
        }
        """
    separator_sheet = """
        QFrame {
            border-top: white;
            border-bottom: white;
            border-left: 1px solid #e4e4e4;
            border-right: 1px solid #e4e4e4;
        }
        """
    panel_style = "background-color: #ffffff; color: black;"

    def _set_widget_style(self, name):
        QApplication.instance().setStyle(name)

    def __init__(self, store: DataStore):
        """Call and setup the UI"""
        self.store = store
        super().__init__(store=store)
        self.default_text_color = QColor("black")
        style = self.style()
        dialog_icon = style.standardIcon(
            QStyle.StandardPixmap.SP_FileDialogDetailedView
        )
        self.setWindowIcon(dialog_icon)
        self.setupUi(self)


def gui():
    store = DataStore()
    try:
        from ctypes import windll

        app_id = f"jjrboucher.ms-word-parser.gui.v{__version__.replace('.','-')}"
        windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except ImportError:
        pass
    ms_word_app = QApplication([__appname__, "windows:darkmode=2"])
    theme = "windowsvista"
    if os.sys.platform in {"linux", "darwin"}:
        theme = "Fusion"
    ms_word_app.setStyle(theme)
    ms_word_app.setApplicationName(__appname__)
    ms_word_app.setApplicationDisplayName(__appname__)
    style = ms_word_app.style()
    icon = style.standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView)
    ms_word_app.setWindowIcon(icon)
    ms_word_gui = MsWordGui(store)
    store.ms_word_gui = ms_word_gui
    ms_word_gui.show()
    ms_word_app.exec()
