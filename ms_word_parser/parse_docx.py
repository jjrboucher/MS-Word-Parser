#!/usr/bin/env python3

import os
import sys
import json
import math
import sqlite3
import re
import logging
import subprocess
import argparse
import threading
from datetime import datetime as dt, timedelta
from pathlib import Path
import warnings
import pandas as pd

from PyQt6.QtCore import (
    QCoreApplication,
    QMetaObject,
    QRect,
    Qt,
    QUrl,
)
from PyQt6.QtGui import (
    QAction,
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
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

try:
    from classes.datastore import DataStore
    from classes.docx import Docx
    from tips import (
        tip_sameRsidRoot,
        tip_numDocumentsEachRsidRoot,
        tip_docsCreatedBySameWindowsUser,
        tip_scriptOverview,
        tip_excelWorksheets,
        tip_processingOptions,
        tip_guiWorkFlow,
    )
except ModuleNotFoundError:
    from ms_word_parser.classes.datastore import DataStore
    from ms_word_parser.classes.docx import Docx
    from ms_word_parser.tips import (
        tip_sameRsidRoot,
        tip_numDocumentsEachRsidRoot,
        tip_docsCreatedBySameWindowsUser,
        tip_scriptOverview,
        tip_excelWorksheets,
        tip_processingOptions,
        tip_guiWorkFlow,
    )

if sys.platform == "win32":
    import msvcrt
    def _read_key():
        return msvcrt.getwch()
else:
    import tty, termios
    def _read_key():
        fd = sys.stdin.fileno()
        old = termios.tcgetattr(fd)
        try:
            tty.setraw(fd)
            return sys.stdin.read(1)
        finally:
            termios.tcsetattr(fd, termios.TCSADRAIN, old)        


warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings("ignore", category=FutureWarning)
green = QColor(86, 208, 50)
red = QColor(204, 0, 0)
black = QColor(0, 0, 0)
__red__ = "\033[1;31m"
__green__ = "\033[1;32m"
__clr__ = "\033[1;m"
__version__ = "3.0.0"
__appname__ = f"MS Word Parser v{__version__}"
__source__ = "https://github.com/jjrboucher/MS-Word-Parser"
__date__ = "01 Mar 2026"
__author__ = (
    "Jacques Boucher - jjrboucher@gmail.com\nCorey Forman - corey@digitalsleuth.ca"
)
__dtfmt__ = "%Y-%m-%d %H:%M:%S"


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
        self.logger.setLevel(logging.INFO)
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
        self.processOptions.setStyleSheet("background: #ffffff; color: black;")
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
        self.operationOptions.setStyleSheet("background-color: #ffffff; color:black;")
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
        self.outputFiles.setStyleSheet("background-color: #ffffff; color: black;")
        self.outputFiles.setFont(self.text_font)
        self.outputPathLabel = QLabel(self.outputFiles)
        self.outputPathLabel.setObjectName("outputPathLabel")
        self.outputPathLabel.setGeometry(QRect(10, 30, 80, 16))
        self.outputPathLabel.setStyleSheet("background: #ffffff; color: black;")
        self.outputPathLabel.setFont(self.text_font)
        self.outputPath = QTextEdit(self.outputFiles)
        self.outputPath.setAlignment(
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft
        )
        self.outputPath.setObjectName("outputPath")
        self.outputPath.setGeometry(QRect(90, 26, 240, 26))
        self.outputPath.setStyleSheet("background: #ffffff; color: black;")
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
        self.generalLog.setStyleSheet("background: #ffffff; color: black;")
        self.generalLog.setFont(self.text_font)
        self.generalLogFile = QTextEdit(self.outputFiles)
        self.generalLogFile.setAlignment(
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft
        )
        self.generalLogFile.setObjectName("generalLogFile")
        self.generalLogFile.setGeometry(QRect(90, 58, 240, 26))
        self.generalLogFile.setStyleSheet("background: #ffffff; color: black;")
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
        self.processStatus.setStyleSheet("background: #ffffff; color: black;")
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
        self.numOfFilesLabel.setStyleSheet("background: #ffffff; color: black;")
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
        self.numOfErrorsLabel = QLabel(self.processStatus)
        self.numOfErrorsLabel.setObjectName("numOfErrorsLabel")
        self.numOfErrorsLabel.setGeometry(QRect(135, 28, 80, 26))
        self.numOfErrorsLabel.setStyleSheet("background: #ffffff; color: black;")
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
        self.numRemainingLabel = QLabel(self.processStatus)
        self.numRemainingLabel.setObjectName("numRemainingLabel")
        self.numRemainingLabel.setGeometry(QRect(257, 28, 120, 26))
        self.numRemainingLabel.setStyleSheet("background: #ffffff; color: black;")
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
        self.menuHelp = QMenu(self.menubar)
        self.menuHelp.setObjectName("menuHelp")
        self.menuHelp.setFont(self.text_font)
        MainWindow.setMenuBar(self.menubar)

        # Menu Bar
        self.menubar.addAction(self.menuFile.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())
        self.menuFile.addAction(self.actionSelect_Output)
        self.menuFile.addSeparator()
        self.menuFile.addAction(self.actionAdd_Ingest)
        self.menuFile.addAction(self.actionAdd_Files)
        self.menuFile.addAction(self.actionAdd_Directory)
        self.menuFile.addSeparator()
        self.menuFile.addAction(self.actionExit)
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
        self.docxOutput.setTextColor(black)
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
        githubLink = f'<a href="{__source__}">View the source on GitHub</a>'
        self.aboutWindow.setWindowTitle("About")
        self.aboutWindow.aboutLabel.setText(
            f"Version: {__appname__}\nLast Updated: {__date__}\n\nAuthors:\n{__author__}"
        )
        self.aboutWindow.urlLabel.setOpenExternalLinks(True)
        self.aboutWindow.urlLabel.setText(githubLink)
        self.aboutWindow.aboutLabel.setFont(self.text_font)
        self.aboutWindow.urlLabel.setFont(self.text_font)
        self.aboutWindow.show()

    def _contents(self):
        self.contentsWindow = ContentsWindow()
        self.contentsWindow.setWindowFlags(
            self.contentsWindow.windowFlags() & ~Qt.WindowType.WindowMinMaxButtonsHint
        )
        self.contentsWindow.show()

    def update_status(self, msg, level="info", color=black):
        levels = {"info": logging.INFO, "error": logging.ERROR, "debug": logging.DEBUG}
        log_level = levels[level]
        if level in {"info", "error"}:
            if self.store.ms_word_gui:
                self.docxOutput.setTextColor(color)
                self.docxOutput.append(f"{dt.now().strftime(__dtfmt__)} - {msg}")
                self.docxOutput.setTextColor(black)
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
        docxErrorCount = 0
        update_status = self.update_status
        script_start = dt.now().strftime(__dtfmt__)
        update_status(f"Script executed: {script_start}")
        update_status("Summary of files parsed:")
        update_status(f'{"="*36}')
        remaining = int(self.numRemaining.toPlainText())
        for f in files:
            if not self.running:
                update_status("Processing stopped")
                self.stopButton.setEnabled(False)
                self.resetButton.setEnabled(True)
                self.resetButton.setStyleSheet(self.stylesheet)
                update_status("Attempting to write current results to Excel")
                try:
                    if self.store.excel:
                        write_to_excel(
                            self.store.excel_file,
                            self.store.triage_files,
                            store=self.store,
                        )
                    if self.store.sqlite:
                        write_to_sqlite(self.store)
                    if docxErrorCount > 0:
                        clr = red
                    else:
                        clr = black
                    update_status(
                        f"Finished writing to Excel. Errors detected: {docxErrorCount}",
                        color=clr,
                    )
                    if docxErrorCount > 0:
                        update_status(
                            "The following files had errors:", "error", color=clr
                        )
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
                except Exception as e:
                    update_status(f"Unable to write results to Excel: {e}")
                return
            try:
                with Docx(f, triage_files, hash_files, self.store) as doc:
                    process_docx(doc, triage_files, hash_files, self.store)
            except Exception as docxError:
                # If processing a DOCx file raises an error, let the user know, and write it
                # to the error log.
                docxErrorCount += 1  # increment error count by 1.
                self.numOfErrors.setText(str(docxErrorCount))
                update_status(
                    f"Error trying to process {f}. Skipping. Error: {str(docxError)}",
                    level="error",
                    color=red,
                )
                self.store.errors_worksheet["File Name"].append(f)
                self.store.errors_worksheet["Error"].append(str(docxError))
            if remaining != 0:
                remaining -= 1
            self.numRemaining.setText(str(remaining))
        if self.store.excel:
            write_to_excel(
                self.store.excel_file, self.store.triage_files, store=self.store
            )
        if self.store.sqlite:
            write_to_sqlite(self.store)
        update_status(f'{"="*24}')
        if docxErrorCount > 0:
            clr = red
        else:
            clr = black
        update_status(
            f"Processing finished for all files. Errors detected: {docxErrorCount}",
            color=clr,
        )
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
        self.resetButton.setEnabled(True)
        self.resetButton.setStyleSheet(self.stylesheet)
        self.stopButton.setEnabled(False)
        self.stopButton.setStyleSheet(self.disabled)
        self.openLogButton.setEnabled(True)
        self.openLogButton.setStyleSheet(self.stylesheet)
        reset_vars(self.store)


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

    def __init__(self, store: DataStore):
        """Call and setup the UI"""
        self.store = store
        super().__init__(store=store)
        style = self.style()
        dialog_icon = style.standardIcon(
            QStyle.StandardPixmap.SP_FileDialogDetailedView
        )
        self.setWindowIcon(dialog_icon)
        self.setupUi(self)


def process_docx(filename, triage, hashing, store: DataStore):
    """
    This function accepts a filename of type Docx and processes it.
    By placing this in a function, it allows the main part of the script to accept multiple file names and
    then loop through them, calling this function for each DOCx file.
    """
    if store.ms_word_gui:
        level = "info"
        update_status = store.ms_word_gui.update_status
    else:
        level = "debug"
        update_status = lambda msg, **kwargs: update_cli(msg, store=store, **kwargs)
    this_file = filename.msword_file
    this_rsid_root = filename.rsid_root()
    xml_files = filename.xml_files
    update_status(f"Processing {this_file}", level="info")
    file_details = filename.details()
    third_party_paths = [
        "word\\settings.xml",
        "docProps\\core.xml",
        "docProps\\app.xml",
    ]
    third_party = False
    for line in file_details.split("\n"):
        update_status(f"    {line.rstrip()}", level=level)
    for checkFile in (
        "word/settings.xml",
        "docProps/core.xml",
        "docProps/app.xml",
        "word\\settings.xml",
        "docProps\\core.xml",
        "docProps\\app.xml",
    ):  # checks if xml files being parsed
        # are present and notes same in the log file.
        xml_exists = checkFile in xml_files.keys()
        if xml_exists and checkFile in third_party_paths:
            third_party = True
        update_status(f"    {checkFile} exists: {xml_exists}", level=level)
        if third_party:
            update_status(
                f"    {this_file} may have been created using something other than MS Word",
                level=level,
            )

    # Writing document summary worksheet.
    headers = [
        "File Name",
        "MD5 Hash",
        "Unique rsidR",
        "RSID Root",
        "<w:p> tags",
        "<w:r> tags",
        "<w:t> tags",
        "<w:tr> tags",
        "<w14:docId>",
        "<w15:docId>",
        "<w16:docId>",
        "Hyperlinks",
        "Spell Check",
        "Grammar Check",
        "Has Comments",
        "Has Ink",
    ]
    if not hashing:
        headers.pop(1)
    store.doc_summary_worksheet = (
        {k: [] for k in headers}
        if not store.doc_summary_worksheet
        else store.doc_summary_worksheet
    )
    w14_id, w15_id, w16_id = filename.get_doc_ids()
    spelling, grammar = filename.get_proof_state()
    if hashing:
        values = [
            this_file,
            filename.hash(),
            len(filename.rsidr()),
            this_rsid_root,
            filename.paragraph_tags(),
            filename.runs_tags(),
            filename.text_tags(),
            filename.table_row_tags(),
            w14_id,
            w15_id,
            w16_id,
            filename.hyperlinks(),
            spelling,
            grammar,
            filename.has_comments,
            filename.has_ink,
        ]
    else:
        values = [
            this_file,
            len(filename.rsidr()),
            this_rsid_root,
            filename.paragraph_tags(),
            filename.runs_tags(),
            filename.text_tags(),
            filename.table_row_tags(),
            w14_id,
            w15_id,
            w16_id,
            filename.hyperlinks(),
            spelling,
            grammar,
            filename.has_comments,
            filename.has_ink,
        ]
    for k, v in zip(headers, values):
        store.doc_summary_worksheet[k].append(v)
    update_status("    Extracted Document Summary artifacts", level=level)

    # The keys will be used as the column heading in the spreadsheet
    # The order they are in is the order that the columns will be in the spreadsheet
    # Corresponding values passed, resulting in a dictionary being passed called allMetadata
    # containing column headings and associated extracted metadata value.

    headers = [
        "File Name",
        "Author",
        "Title",
        "Subject",
        "RSID Root",
        "Template",
        "Created Date",
        "Modified Date",
        "Last Printed Date",
        "Last Modified By",
        "Total Editing Time",
        "Revision",
        "Manager",
        "Company",
        "Pages",
        "Paragraphs",
        "Lines",
        "Words",
        "Characters",
        "Characters With Spaces",
        "Keywords",
        "Description",
        "Category",
        "Application",
        "App Version",
        "Doc Security",
        "Content Status",
        "Language",
        "Version",
        "Shared Doc",
        "Hyperlinks Changed",
    ]
    store.metadata_worksheet = (
        {k: [] for k in headers}
        if not store.metadata_worksheet
        else store.metadata_worksheet
    )
    values = [
        this_file,
        filename.get_metadata("creator"),
        filename.get_metadata("title"),
        filename.get_metadata("subject"),
        this_rsid_root,
        filename.get_metadata("Template"),
        filename.adjust_timestamp(filename.get_metadata("created")),
        filename.adjust_timestamp(filename.get_metadata("modified")),
        filename.adjust_timestamp(filename.get_metadata("lastPrinted")),
        filename.get_metadata("lastModifiedBy"),
        filename.get_metadata("TotalTime"),
        filename.get_metadata("revision"),
        filename.get_metadata("Manager"),
        filename.get_metadata("Company"),
        filename.get_metadata("Pages"),
        filename.get_metadata("Paragraphs"),
        filename.get_metadata("Lines"),
        filename.get_metadata("Words"),
        filename.get_metadata("Characters"),
        filename.get_metadata("CharactersWithSpaces"),
        filename.get_metadata("keywords"),
        filename.get_metadata("description"),
        filename.get_metadata("category"),
        filename.get_metadata("Application"),
        filename.get_metadata("AppVersion"),
        filename.get_metadata("DocSecurity"),
        filename.get_metadata("contentStatus"),
        filename.get_metadata("language"),
        filename.get_metadata("version"),
        filename.get_metadata("SharedDoc"),
        filename.get_metadata("HyperlinksChanged"),
    ]
    for k, v in zip(headers, values):
        store.metadata_worksheet[k].append(v)
    update_status("    Extracted metadata artifacts", level=level)

    if filename.any_comments():  # checks if there are comments
        headers = [
            "File Name",
            "Author",
            "Initials",
            "Timestamp (UTC)",
            "Comment ID",
            "Comment paraId",
            "paraId Text",
        ]
        store.comments_worksheet = (
            {k: [] for k in headers}
            if not store.comments_worksheet
            else store.comments_worksheet
        )
        for comment in filename.get_comments():
            values = [
                this_file,  # Filename
                comment[3],  # Author
                comment[4],  # Initials
                filename.adjust_timestamp(comment[2]),  # Timestamp
                comment[0],  # ID
                comment[1],  # paraId for later correlation
                comment[5],  # ParaId Text
            ]
            for k, v in zip(headers, values):
                store.comments_worksheet[k].append(v)
        update_status("    Extracted comments artifacts", level=level)

    if not triage:  # will generate these spreadsheets if not triage
        headers = [
            "File Name",
            "Archive File",
            "MD5 Hash",
            "Modified Time (local/UTC/Redmond, Washington)",
            # expressed local time if Mac/iOS Pages exported to MS Word
            # expressed in UTC if created by LibreOffice on Windows exporting to MS Word.
            # expressed Redmond, Washington time zone when edited with MS Word online.
            "Uncompressed Size (bytes)",
            "ZIP Compression Type",
            "ZIP Create System",
            "ZIP Created Version",
            "ZIP Extract Version",
            "ZIP Flag Bits (hex)",
            "ZIP Extra Flag (len)",
            "ZIP Extra Bytes (truncated)",
        ]
        if not hashing:
            headers.pop(2)

        store.archive_files_worksheet = (
            {k: [] for k in headers}
            if not store.archive_files_worksheet
            else store.archive_files_worksheet
        )
        for xml, xml_info in xml_files.items():
            values = [
                this_file,
                xml,
                xml_info["MD5"],
                filename.adjust_timestamp(xml_info["Modified Time"]),
                xml_info["File Size"],
                xml_info["Zip Compression"],
                xml_info["Zip Create System"],
                xml_info["Zip Create Version"],
                xml_info["Zip Extract Version"],
                xml_info["Zip Flag Bits"],
                xml_info["Zip Extra Fields Length"],
                xml_info["Zip Extra Fields Bytes"],
            ]
            if not hashing:
                values.pop(2)
            for k, v in zip(headers, values):
                store.archive_files_worksheet[k].append(v)

        update_status("    Extracted archive files artifacts", level=level)

        # Calculating count of rsidR, rsidRPr, rsidP, rsidRDefault, rsidTr, paraId, and textId in document.xml
        # and writing to "rsids" worksheet
        headers = [
            "File Name",
            "RSID Root",
            "RSID Type",
            "RSID Value",
            "Count in document.xml",
            "File Created Date",
            "File Modified Date",
        ]
        store.rsids_worksheet = (
            {k: [] for k in headers}
            if not store.rsids_worksheet
            else store.rsids_worksheet
        )
        file_idx = store.metadata_worksheet["File Name"].index(this_file)
        created_dt = store.metadata_worksheet["Created Date"][file_idx]
        modified_dt = store.metadata_worksheet["Modified Date"][file_idx]
        rsid_lookups = [
            ("rsidR", filename.rsidr_in_document_xml),
            ("rsidP", filename.rsidp_in_document_xml),
            ("rsidRPr", filename.rsidrpr_in_document_xml),
            ("rsidRDefault", filename.rsidrdefault_in_document_xml),
            ("rsidTr", filename.rsidtr_in_document_xml),
            ("paraID", filename.paragraph_id_tags),
            ("textID", filename.text_id_tags),
        ]
        ws = store.rsids_worksheet
        cols = [ws[h] for h in headers]
        for label, func in rsid_lookups:
            update_status(f"    Calculating {label} count", level=level)
            for k, v in func().items():
                cols[0].append(this_file)
                cols[1].append(this_rsid_root)
                cols[2].append(label)
                cols[3].append(k)
                cols[4].append(v)
                cols[5].append(created_dt)
                cols[6].append(modified_dt)
        all_people = filename.get_people()
        if all_people:
            update_status(
                "    Processing people information from document", level=level
            )
            headers = ["File Name", "Author", "providerId", "userId"]
            store.people_worksheet = (
                {k: [] for k in headers}
                if not store.people_worksheet
                else store.people_worksheet
            )
            for each_person in all_people:
                values = [this_file, each_person[0], each_person[1], each_person[2]]
                for k, v in zip(headers, values):
                    store.people_worksheet[k].append(v)

        extensible_comments = filename.get_extensible_comments()
        if extensible_comments:
            update_status("    Processing extensible comments data", level=level)
            headers = [
                "File Name",
                "durableId",
                "dateUtc",
                "reactionType",
                "reactionDateUtc",
                "uri",
                "userId",
                "userProvider",
                "userName",
            ]
            store.extensible_worksheet = (
                {k: [] for k in headers}
                if not store.extensible_worksheet
                else store.extensible_worksheet
            )
            for comment, data in extensible_comments.items():
                idx = 3
                while idx + 1 <= len(data):
                    values = [
                        this_file,
                        comment,
                        filename.adjust_timestamp(data[0]),
                        data[2],
                        filename.adjust_timestamp(data[idx][0]),
                        data[1],
                        data[idx][1],
                        data[idx][2],
                        data[idx][3],
                    ]
                    for k, v in zip(headers, values):
                        store.extensible_worksheet[k].append(v)
                    idx += 1

        extended_comments = filename.get_extended_comments()
        if extended_comments:
            update_status("    Processing extensible comments data", level=level)
            headers = ["File Name", "paraId", "paraIdParent", "Done"]
            store.extended_worksheet = (
                {k: [] for k in headers}
                if not store.extended_worksheet
                else store.extended_worksheet
            )
            for comment in extended_comments:
                values = [this_file, comment[0], comment[1], bool(int(comment[2]))]
                for k, v in zip(headers, values):
                    store.extended_worksheet[k].append(v)

        comments_ids = filename.get_comments_ids()
        if comments_ids:
            update_status("    Processing comments ids", level=level)
            headers = ["File Name", "paraId", "durableId"]
            store.comments_ids_worksheet = (
                {k: [] for k in headers}
                if not store.comments_ids_worksheet
                else store.comments_ids_worksheet
            )
            for comments_id in comments_ids:
                values = [this_file, comments_id[0], comments_id[1]]
                for k, v in zip(headers, values):
                    store.comments_ids_worksheet[k].append(v)

        custom_props = filename.get_custom_xml()
        if custom_props:
            update_status("    Processing custom properties", level=level)
            if not store.custom_xml_worksheet:
                headers = ["File Name"]
                store.custom_xml_worksheet = {h: [] for h in headers}
            else:
                headers = list(store.custom_xml_worksheet.keys())
            for k in custom_props.keys():
                if k not in store.custom_xml_worksheet:
                    headers.append(k)
                    store.custom_xml_worksheet[k] = ["None"] * len(
                        next(iter(store.custom_xml_worksheet.values()), [])
                    )
            for h in headers:
                if h == "File Name":
                    store.custom_xml_worksheet[h].append(this_file)
                else:
                    store.custom_xml_worksheet[h].append(custom_props.get(h, "None"))

        if filename.item_files:
            xml_content = filename.get_all_content(filename.item_files)
            if xml_content:
                item_xml_content = xml_content[this_file]
                update_status("    Processing item.xml files", level=level)
                if not store.item_worksheet:
                    headers = ["File Name", "Item XML File", "Content"]
                    store.item_worksheet = {h: [] for h in headers}
                else:
                    headers = list(store.item_worksheet.keys())
                for item_file in filename.item_files:
                    parsed_content = {}
                    store.item_worksheet["File Name"].append(this_file)
                    store.item_worksheet["Item XML File"].append(item_file)
                    entry = item_xml_content[item_file]
                    for k, v in entry.items():
                        if k in parsed_content:
                            parsed_content[k] = f"{parsed_content[k]},{v}"
                        else:
                            parsed_content[k] = v
                    if parsed_content:
                        store.item_worksheet["Content"].append(
                            json.dumps(parsed_content, indent=2)
                        )
                    else:
                        store.item_worksheet["Content"].append(None)

        if filename.ink_files:
            ink_content = filename.get_ink()
            if ink_content:
                update_status("    Processing ink.xml files", level=level)
                if not store.ink_worksheet:
                    headers = ["File Name", "Ink XML File", "Timestamp (UTC)"]
                    store.ink_worksheet = {h: [] for h in headers}
                else:
                    headers = list(store.ink_worksheet.keys())
                for ink_file in ink_content:
                    values = [
                        this_file,
                        ink_file[0],
                        filename.adjust_timestamp(ink_file[1]),
                    ]
                    for k, v in zip(headers, values):
                        store.ink_worksheet[k].append(v)

    update_status(f"Finished processing {this_file}", level="info")
    update_status(f'{"-"*36}', level="info")


def chunk_df(data, sheet_name, chunk_size=1000000):
    df = data if isinstance(data, pd.DataFrame) else pd.DataFrame(data)
    if len(df) > chunk_size:
        for i in range(0, len(df), chunk_size):
            chunk = df.iloc[i : i + chunk_size].copy()
            yield chunk, f"{sheet_name}_{ (i // chunk_size) + 1 }"
    else:
        yield df.copy(), sheet_name


def write_to_excel(excel_file, triage_files, store: DataStore):
    if store.ms_word_gui:
        update_status = store.ms_word_gui.update_status
    else:
        update_status = lambda msg, **kwargs: update_cli(msg, store=store, **kwargs)
    options = {
        "engine": "xlsxwriter",
        "mode": "w",
        "datetime_format": "yyyy-mm-dd hh:mm:ss",
    }
    LAYOUTS = {
        "summary": [
            (0, 0, 52),
            (1, 1, 36),
            (2, 8, 16),
            (9, 10, 42),
            (11, 11, 36),
            (12, None, 20),
        ],
        "metadata": [
            (0, 0, 52),
            (1, 3, 22),
            (4, 4, 14),
            (5, 5, 38),
            (6, 8, 20),
            (9, 9, 42),
            (10, 10, 20),
            (11, 11, 14),
            (12, 13, 42),
            (14, 19, 25),
            (20, 23, 42),
            (24, None, 20),
        ],
        "comments": [(0, 0, 52), (1, 1, 24), (2, 2, 10), (3, 5, 20), (-1, -1, 140)],
        "extensible": [(0, 0, 52), (1, 4, 20), (5, 5, 40), (6, None, 20)],
        "extended": [(0, 0, 52), (1, None, 20)],
        "comments_ids": [(0, 0, 52), (1, None, 14)],
        "people": [(0, 0, 52), (1, 2, 20), (3, 3, 52)],
        "rsids": [(0, 0, 52), (1, 3, 18), (4, 4, 26), (5, None, 20)],
        "custom": [(0, 0, 52), (1, None, 40)],
        "archive": [(0, 0, 52), (1, 2, 36), (3, 3, 50), (4, 10, 30), (11, 11, 44)],
        "item": [(0, 0, 52), (1, 1, 30), (2, 2, 255)],
        "ink": [(0, 0, 52), (1, 1, 30), (2, 2, 20)],
        "aggregated": [
            (0, 0, 52),
            (1, 1, 22),
            (2, 2, 14),
            (3, 3, 20),
            (4, 10, 25),
            (11, 14, 20),
            (15, None, 25),
        ],
        "timeline": [(0, 0, 52), (1, 3, 20), (4, None, 32)],
        "errors": [(0, None, 52)],
    }
    type_map = store.type_map
    with pd.ExcelWriter(path=excel_file, **options) as writer:
        aggregated = False

        def process_and_write(data, name, layout_type):
            if data is None or (
                isinstance(data, (pd.DataFrame, list, dict)) and len(data) == 0
            ):
                return
            first_chunk = True
            date_cols = []
            for df_chunk, actual_name in chunk_df(data, name):
                if df_chunk.empty:
                    continue
                if first_chunk:
                    date_cols = [
                        col
                        for col in df_chunk.columns
                        if any(w in col.lower() for w in ("date", "time", "timestamp"))
                    ]
                    first_chunk = False
                for col_name in df_chunk.columns:
                    if col_name in date_cols:
                        df_chunk[col_name] = pd.to_datetime(
                            df_chunk[col_name],
                            errors="coerce",
                            format="%Y-%m-%d %H:%M:%S",
                        )
                    elif col_name in type_map:
                        df_chunk[col_name] = df_chunk[col_name].astype(
                            type_map[col_name]
                        )
                    else:
                        df_chunk[col_name] = df_chunk[col_name].astype("string")
                df_chunk.to_excel(writer, sheet_name=actual_name, index=False)
                fn_col_max = max(
                    df_chunk["File Name"].astype(str).map(len).max(),
                    len(str("File Name")),
                )
                ws = writer.sheets[actual_name]
                max_row, max_col = df_chunk.shape
                layout = LAYOUTS.get(layout_type, [(0, None, 25)])
                for start, end, width in layout:
                    if start == 0:
                        width = fn_col_max
                    final_start = (max_col + start) if start < 0 else start
                    if end is None:
                        final_end = max_col - 1
                    elif end < 0:
                        final_end = max_col + end
                    else:
                        final_end = end
                    ws.set_column(final_start, final_end, width)
                ws.autofilter(0, 0, max_row, max_col - 1)
                ws.freeze_panes(1, 0)
                update_status(f'"{actual_name}" written.', level="info")
                del df_chunk
            del data

        triage_sheets = [
            (store.doc_summary_worksheet, "Document Summary", "summary"),
            (store.metadata_worksheet, "Metadata", "metadata"),
            (store.comments_worksheet, "Comments", "comments"),
        ]
        if not triage_files:
            full_sheets = [
                (store.extensible_worksheet, "Extensible Comments", "extensible"),
                (store.extended_worksheet, "Extended Comments", "extended"),
                (store.comments_ids_worksheet, "Comments IDs", "comments_ids"),
                (store.people_worksheet, "People", "people"),
                (store.rsids_worksheet, "RSIDs", "rsids"),
                (store.custom_xml_worksheet, "Custom Properties", "custom"),
                (store.archive_files_worksheet, "Archive Files", "archive"),
                (store.item_worksheet, "Item XML Files", "item"),
                (store.ink_worksheet, "Ink XML Files", "ink"),
                (store.errors_worksheet, "Errors", "errors"),
            ]
            if all(
                [
                    store.comments_worksheet,
                    store.comments_ids_worksheet,
                    store.extended_worksheet,
                    store.extensible_worksheet,
                ]
            ):
                df_c = pd.DataFrame(store.comments_worksheet)
                df_e = pd.DataFrame(store.extended_worksheet)
                df_ci = pd.DataFrame(store.comments_ids_worksheet)
                df_ex = pd.DataFrame(store.extensible_worksheet)
                merged = pd.merge(
                    df_c,
                    df_e,
                    left_on=["File Name", "Comment paraId"],
                    right_on=["File Name", "paraId"],
                    how="left",
                    suffixes=("", "_ext"),
                )
                merged = pd.merge(
                    merged,
                    df_ci,
                    on=["File Name", "paraId"],
                    how="left",
                    suffixes=("", "_cid"),
                )
                merged = pd.merge(
                    merged,
                    df_ex,
                    on=["File Name", "durableId"],
                    how="left",
                    suffixes=("", "_extensible"),
                )
                merged = merged.loc[
                    :, ~merged.columns.str.endswith(("_ext", "_cid", "_extensible"))
                ]
                store.aggregated_worksheet = merged
                del df_c, df_e, df_ci, df_ex, merged
                aggregated = True
        for sheet, sheet_name, layout in triage_sheets:
            process_and_write(sheet, sheet_name, layout)
        if not triage_files:
            for sheet, sheet_name, layout in full_sheets:
                process_and_write(sheet, sheet_name, layout)
        if aggregated:
            process_and_write(
                store.aggregated_worksheet, "Aggregated Comment Data", "aggregated"
            )
        if store.timeline:
            update_status(
                "Generating Timeline worksheet - this may take some time depending on the number of documents being parsed ...",
                level="info",
            )
            store.timeline_worksheet = generate_timeline(store)
            if isinstance(store.timeline_worksheet, pd.DataFrame) and not (store.timeline_worksheet).empty:
                process_and_write(store.timeline_worksheet, "Timeline", "timeline")
                generate_visual_timeline(writer, store.timeline_worksheet)
                update_status('"Visual Timeline" written.', level="info")
            else:
                update_status('"Timeline Worksheet" is empty. No data written.', level="info")
        write_tips(writer)
        update_status('"Tips" worksheet written.', level="info")
        update_status(f"All Excel data written to {store.excel_file}", level="info")


def write_to_sqlite(store):
    if store.ms_word_gui:
        update_status = store.ms_word_gui.update_status
        level = "info"
    else:
        update_status = lambda msg, **kwargs: update_cli(msg, store=store, **kwargs)
        level = "debug"
    sql_type_map = {
        re.sub(
            r"[^a-z0-9]",
            "_",
            k.lower()
            .replace("<", "")
            .replace(">", "")
            .replace("(", "")
            .replace(")", "")
            .replace(",", ""),
        ): store.sqlite_types.get(v, "TEXT")
        for k, v in store.type_map.items()
    }
    if os.path.exists(store.sqlite_file):
        try:
            os.remove(store.sqlite_file)
        except:
            update_status(f'Unable to remove "{store.sqlite_file}".', level=level)
            store.sqlite_file = os.path.normpath(
                f"{store.output_path}{os.sep}{store.basename}_2.db"
            )
    update_status(
        f'Writing results to SQLite database "{store.sqlite_file}".', level="info"
    )
    conn = sqlite3.connect(store.sqlite_file)
    triage_sheets = [
        (store.doc_summary_worksheet, "Document Summary", "summary"),
        (store.metadata_worksheet, "Metadata", "metadata"),
        (store.comments_worksheet, "Comments", "comments"),
    ]
    if not store.triage_files:
        full_sheets = [
            (store.extensible_worksheet, "Extensible Comments", "extensible"),
            (store.extended_worksheet, "Extended Comments", "extended"),
            (store.comments_ids_worksheet, "Comments IDs", "comments_ids"),
            (store.people_worksheet, "People", "people"),
            (store.rsids_worksheet, "RSIDs", "rsids"),
            (store.custom_xml_worksheet, "Custom Properties", "custom"),
            (store.archive_files_worksheet, "Archive Files", "archive"),
            (store.item_worksheet, "Item XML Files", "item"),
            (store.ink_worksheet, "Ink XML Files", "ink"),
            (store.errors_worksheet, "Errors", "errors"),
        ]
    for sheet, sheet_name, _ in triage_sheets:
        if sheet:
            cols = ["id INTEGER PRIMARY KEY AUTOINCREMENT"]
            new_sheet, new_name = restructure_sheet(sheet, sheet_name)
            df = pd.DataFrame(new_sheet)
            df_cols = df.columns.tolist()
            for col in df_cols:
                sql_type = sql_type_map.get(col, "TEXT")
                cols.append(f"{col} {sql_type}")
            all_cols = ",\n    ".join(cols)
            pk_stmt = f"CREATE TABLE IF NOT EXISTS {new_name} (\n    {all_cols}\n);"
            cursor = conn.cursor()
            cursor.execute(pk_stmt)
            df.to_sql(new_name, conn, if_exists="append", index=False)
            del new_sheet
            del sheet
    if not store.triage_files:
        for sheet, sheet_name, _ in full_sheets:
            if sheet:
                cols = ["id INTEGER PRIMARY KEY AUTOINCREMENT"]
                new_sheet, new_name = restructure_sheet(sheet, sheet_name)
                df = pd.DataFrame(new_sheet)
                df_cols = df.columns.tolist()
                for col in df_cols:
                    sql_type = sql_type_map.get(col, "TEXT")
                    cols.append(f"{col} {sql_type}")
                all_cols = ",\n    ".join(cols)
                pk_stmt = f"CREATE TABLE IF NOT EXISTS {new_name} (\n    {all_cols}\n);"
                cursor = conn.cursor()
                cursor.execute(pk_stmt)
                df.to_sql(new_name, conn, if_exists="append", index=False)
                del new_sheet
                del sheet
    if all(
        [
            store.comments_worksheet,
            store.comments_ids_worksheet,
            store.extended_worksheet,
            store.extensible_worksheet,
        ]
    ):
        agg_view_stmt = """
        CREATE VIEW "Aggregated Comments" AS
        SELECT DISTINCT C.file_name,
        C.author,
        C.initials,
        C.timestamp_utc,
        C.comment_id,
        C.comment_paraid,
        C.paraid_text,
        EC.paraidparent,
                CASE EC.done
                        WHEN 0 THEN "FALSE"
                        WHEN 1 THEN "TRUE"
                END AS done,
                CID.durableid,
                EC2.dateutc,
                EC2.reactiontype,
                EC2.reactiondateutc,
                EC2.uri,
                EC2.userid,
                EC2.userprovider,
                EC2.username
        FROM comments AS C
        LEFT JOIN (SELECT file_name, paraid, paraidparent, done FROM extended_comments) AS EC ON C.comment_paraid == EC.paraid AND C.file_name == EC.file_name
        LEFT JOIN (SELECT file_name, paraid, durableid FROM comments_ids) AS CID ON C.comment_paraid == CID.paraid AND C.file_name == CID.file_name
        LEFT JOIN (SELECT file_name, durableid, dateutc, reactiontype, reactiondateutc, uri, userid, userprovider, username FROM extensible_comments) AS EC2 ON CID.durableid == EC2.durableid AND CID.file_name == EC2.file_name
        """
        conn.execute(agg_view_stmt)
    if all(
        [
            store.metadata_worksheet,
            store.comments_worksheet,
            store.extensible_worksheet,
            store.rsids_worksheet,
            store.archive_files_worksheet,
            store.ink_worksheet,
        ]
    ):
        timeline_view_stmt = """
            CREATE VIEW "Timeline View" AS
            SELECT file_name AS "File Name", created_date AS "Timestamp", 'created' AS "Type", NULL AS "Value", 'Metadata' AS "Source"
            FROM metadata WHERE timestamp IS NOT NULL AND timestamp != ''
            UNION ALL
            SELECT file_name, modified_date, 'modified', NULL, 'Metadata'
            FROM metadata WHERE modified_date IS NOT NULL AND modified_date != ''
            UNION ALL
            SELECT file_name, last_printed_date, 'last printed', NULL, 'Metadata'
            FROM metadata WHERE last_printed_date IS NOT NULL AND last_printed_date != ''
            UNION ALL
            SELECT file_name, timestamp_utc, 'comment', paraid_text, 'Comments'
            FROM comments WHERE timestamp_utc IS NOT NULL AND timestamp_utc != ''
            UNION ALL
            SELECT file_name, dateutc, 'durableid', durableid, 'Extensible Comments'
            FROM extensible_comments WHERE dateutc IS NOT NULL AND dateutc != ''
            UNION ALL
            SELECT file_name, reactiondateutc, 'reaction', NULL, 'Extensible Comments'
            FROM extensible_comments WHERE reactiondateutc IS NOT NULL AND reactiondateutc != ''
            UNION ALL
            SELECT file_name, file_created_date, 'created - rsid', (rsid_type || ' - ' || rsid_value), 'RSIDs'
            FROM rsids WHERE file_created_date IS NOT NULL AND file_created_date != ''
            UNION ALL
            SELECT file_name, file_modified_date, 'modified - rsid', (rsid_type || ' - ' || rsid_value), 'RSIDs'
            FROM rsids WHERE file_modified_date IS NOT NULL AND file_modified_date != ''
            UNION ALL
            SELECT file_name, modified_time_local_utc_redmond_washington, 'modified - archive file', archive_file, 'Archive Files'
            FROM archive_files WHERE modified_time_local_utc_redmond_washington IS NOT NULL AND modified_time_local_utc_redmond_washington != ''
            UNION ALL
            SELECT file_name, timestamp_utc, 'ink file', ink_xml_file, 'Ink XML Files'
            FROM ink_xml_files WHERE timestamp_utc IS NOT NULL AND timestamp_utc != ''
            ORDER BY "Timestamp" ASC;
            """
        conn.execute(timeline_view_stmt)
    try:
        conn.close()
        update_status(
            f'All SQLite data written to "{store.sqlite_file}".', level="info"
        )
    except:
        update_status(
            f'SQLite database could not be written to "{store.sqlite_file}".',
            level="error",
        )


def restructure_sheet(sheet, sheet_name):
    if sheet:
        new_sheet = {
            re.sub(
                r"[^a-z0-9]",
                "_",
                k.lower()
                .replace("<", "")
                .replace(">", "")
                .replace("(", "")
                .replace(")", "")
                .replace(",", ""),
            ): v
            for k, v in sheet.items()
        }
        new_name = sheet_name.lower().replace(" ", "_")
        return new_sheet, new_name
    return None, None


def write_tips(writer):
    workbook = writer.book
    tips_ws = workbook.add_worksheet("Excel Tips")
    writer.sheets["Excel Tips"] = tips_ws
    tip_nums = {1: ["A1", [510, 180]], 2: ["I1", [890, 550]], 3: ["W1", [1000, 810]]}
    tip_num = 1
    for tip in (
        tip_sameRsidRoot,
        tip_numDocumentsEachRsidRoot,
        tip_docsCreatedBySameWindowsUser,
    ):
        text = f"{tip['Title']}\n\n{tip['Text']}"
        options = {
            "width": tip_nums[tip_num][1][0],
            "height": tip_nums[tip_num][1][1],
            "x_offset": 1,
            "y_offset": 1,
            "align": {"vertical": "top", "horizontal": "center"},
            "line": {"color": "black", "width": 2},
        }
        tips_ws.insert_textbox(tip_nums[tip_num][0], text, options)
        tip_num += 1


def generate_timeline(store):
    parts = []

    def create_part(sheet, timestamp_col, type_name, value_col=None, source_name=""):
        if not sheet or timestamp_col not in sheet:
            return
        temp_df = pd.DataFrame(
            {
                "File Name": sheet["File Name"],
                "Timestamp": sheet[timestamp_col],
                "Type": type_name,
                "Source": source_name,
            }
        )
        if isinstance(value_col, list):
            temp_df["Value"] = (
                pd.Series(sheet[value_col[0]]).astype(str)
                + " - "
                + pd.Series(sheet[value_col[1]]).astype(str)
            )
        elif isinstance(value_col, str) and value_col in sheet:
            temp_df["Value"] = sheet[value_col]
        else:
            temp_df["Value"] = None

        temp_df["Timestamp"] = pd.to_datetime(
            temp_df["Timestamp"], errors="coerce", format="%Y-%m-%d %H:%M:%S"
        )
        parts.append(temp_df.dropna(subset=["Timestamp"]))

    if store.metadata_worksheet:
        create_part(
            store.metadata_worksheet, "Created Date", "created", source_name="Metadata"
        )
        create_part(
            store.metadata_worksheet,
            "Modified Date",
            "modified",
            source_name="Metadata",
        )
        create_part(
            store.metadata_worksheet,
            "Last Printed Date",
            "last printed",
            source_name="Metadata",
        )
    if store.comments_worksheet:
        create_part(
            store.comments_worksheet,
            "Timestamp (UTC)",
            "comment",
            "paraId Text",
            "Comments",
        )
    if store.extensible_worksheet:
        create_part(
            store.extensible_worksheet,
            "dateUtc",
            "durableId",
            "durableId",
            "Extensible Comments",
        )
        if "reactionDateUtc" in store.extensible_worksheet:
            create_part(
                store.extensible_worksheet,
                "reactionDateUtc",
                "reaction",
                source_name="Extensible Comments",
            )
    if store.rsids_worksheet:
        create_part(
            store.rsids_worksheet,
            "File Created Date",
            "created - rsid",
            ["RSID Type", "RSID Value"],
            "RSIDs",
        )
        create_part(
            store.rsids_worksheet,
            "File Modified Date",
            "modified - rsid",
            ["RSID Type", "RSID Value"],
            "RSIDs",
        )
    if store.archive_files_worksheet:
        create_part(
            store.archive_files_worksheet,
            "Modified Time (local/UTC/Redmond, Washington)",
            "modified - archive file",
            "Archive File",
            "Archive Files",
        )
    if store.ink_worksheet:
        create_part(
            store.ink_worksheet,
            "Timestamp (UTC)",
            "ink file",
            "Ink XML File",
            "Ink XML Files",
        )

    if not parts:
        return pd.DataFrame(
            columns=["File Name", "Timestamp", "Type", "Value", "Source"]
        )
    full_timeline_df = pd.concat(parts, ignore_index=True)
    full_timeline_df.sort_values("Timestamp", inplace=True)
    return full_timeline_df


def generate_visual_timeline(writer, sheet):
    if isinstance(sheet, pd.DataFrame) and sheet.empty:
        return
    counts = sheet["Timestamp"].value_counts().sort_index()
    counts.index = counts.index.strftime("%Y-%m-%d %H:%M:%S")
    tl_chart = counts.reset_index()
    tl_chart.columns = ["Timestamp", "Count"]
    ts_width = max(tl_chart["Timestamp"].map(len).max(), len("Timestamp")) + 2
    tl_chart["Timestamp"] = pd.to_datetime(
        tl_chart["Timestamp"], format="%Y-%m-%d %H:%M:%S"
    )
    tl_chart.to_excel(writer, sheet_name="Visual Timeline", index=False)
    workbook = writer.book
    worksheet = writer.sheets["Visual Timeline"]

    worksheet.set_column(0, 0, ts_width)
    worksheet.set_column(1, 1, 10)
    num_rows = len(tl_chart)
    max_count = tl_chart["Count"].max()
    grid_count = max_count / 6
    min_date = tl_chart["Timestamp"].min()
    max_date = tl_chart["Timestamp"].max()
    day_span = (max_date - min_date).days
    day_count = day_span / 6

    def round_val(grid):
        if grid <= 0:
            return 1
        mag = 10 ** math.floor(math.log10(grid))
        norm = grid / mag
        if norm <= 1:
            step = 1
        elif norm <= 2:
            step = 2
        elif norm <= 5:
            step = 5
        else:
            step = 10
        return int(step * mag)

    major_unit = round_val(grid_count)
    major_unit_days = round_val(day_count)
    chart = workbook.add_chart({"type": "column"})
    chart.add_series(
        {
            "name": "Event Count",
            "categories": ["Visual Timeline", 1, 0, num_rows, 0],
            "values": ["Visual Timeline", 1, 1, num_rows, 1],
            "gap": 100,
            "fill": {"color": "#6366f1"},
            "border": {"color": "#4f46e5"},
            "data_labels": {
                "value": True,
                "position": "outside_end",
                "font": {"size": 7, "color": "#334155"},
            },
        }
    )
    chart.set_title(
        {
            "name": "Visual Timeline",
            "name_font": {"size": 14, "bold": True, "color": "#0f172a"},
        }
    )
    chart.set_x_axis(
        {
            "name": "Date / Time",
            "name_font": {"size": 11, "bold": True},
            "num_font": {"rotation": -45, "size": 7},
            "major_gridlines": {"visible": False},
            "date_axis": True,
            "min": min_date,
            "max": max_date,
            "major_unit": major_unit_days,
            "major_unit_type": "days",
            "minor_unit": max(1, major_unit_days // 4),
            "minor_unit_type": "days",
            "num_format": "yyyy-mm-dd",
        }
    )
    chart.set_y_axis(
        {
            "name": "Count",
            "name_font": {"size": 11, "bold": True},
            "major_gridlines": {"visible": True, "line": {"color": "#e2e8f0"}},
            "min": 0,
            "major_unit": major_unit,
            "max": max_count + major_unit,
        }
    )
    chart.set_legend({"none": True})
    chart.set_chartarea({"border": {"none": True}, "fill": {"color": "#f8fafc"}})
    chart.set_plotarea({"border": {"none": True}, "fill": {"color": "#ffffff"}})
    chart.set_size({"width": 1400, "height": 500})
    worksheet.insert_chart("D2", chart)


def get_files(folder_path, recursive=False):
    if recursive:
        yield from folder_path.rglob("*.docx")
        yield from folder_path.rglob("*.dotx")
        yield from folder_path.rglob("*.dotm")
        yield from folder_path.rglob("*.docm")
    else:
        yield from folder_path.glob("*.docx")
        yield from folder_path.glob("*.dotx")
        yield from folder_path.glob("*.dotm")
        yield from folder_path.glob("*.docm")


def update_cli(msg, level="info", color=__clr__, store: DataStore = None):
    if store is None:
        return
    levels = {"info": logging.INFO, "error": logging.ERROR, "debug": logging.DEBUG}
    log_level = levels[level]
    if isinstance(store.color_fmt, ColorFormatter):
        store.color_fmt.set_color(color)
        store.logger.log(log_level, msg)
        store.color_fmt.set_color("")
    else:
        store.logger.log(log_level, msg)

def start_keypress_listener(status_callback, quit_key="q", status_key="s"):
    stop_event = threading.Event()
    def _listen():
        while not stop_event.is_set():
            key = _read_key()
            if key == status_key:
                status_callback()
            elif key == quit_key:
                print("[QUIT KEY (q) PRESSED - STANDBY ...] Attempting to write already processed data")
                stop_event.set()
    thread = threading.Thread(target=_listen, daemon=True)
    thread.start()
    return stop_event

def process_cli(files, triage_files, hash_files, store: DataStore, ingest=False):
    def print_status():
        print(f"[STATUS] File: {f} | {store.done} / {store.total} | {round((int(store.done) / int(store.total) * 100), 2)} %")
    stop_event = start_keypress_listener(
        status_callback=print_status,
        status_key="s",
        quit_key="q",
    )
    docxErrorCount = 0
    store.start_time = dt.now().strftime(__dtfmt__)
    update_cli(f"{__appname__}", store=store)
    update_cli(f"Command line: {' '.join(sys.argv)}", store=store)
    update_cli(f"Output File Path: {store.output_path}", store=store)
    if store.excel:
        update_cli(f"Excel output file: {store.excel_file}", store=store)
    if store.sqlite:
        update_cli(f"SQLite DB file: {store.sqlite_file}", store=store)
    update_cli(f"Log file: {os.path.abspath(store.log_file)}", store=store)
    if ingest:
        file_list, missing = read_ingest(files)
        store.filenames = file_list
        if missing:
            update_cli(
                f"The following {len(missing)} file(s) do not exist:",
                color=__red__,
                store=store,
            )
            joiner = f"\n{dt.now().strftime(__dtfmt__)} -     "
            update_cli("    " + joiner.join(missing), color=__red__, store=store)
        if len(file_list) > 1:
            update_cli(
                f"The following {len(file_list)} files have been loaded:",
                store=store,
            )
        elif len(file_list) == 1:
            update_cli(
                f"The following {len(file_list)} file has been loaded:", store=store
            )
        else:
            update_cli(
                "No files were loaded. Please check the file paths and try again.",
                level="error",
                color=__red__,
                store=store,
            )
            return
        joiner = f"\n{dt.now().strftime(__dtfmt__)} -     "
        update_cli("    " + joiner.join(file_list), store=store)
        files = file_list
    else:
        update_cli(
            f"The following {len(files)} files are being processed:", store=store
        )
        joiner = f"\n{dt.now().strftime(__dtfmt__)} -     "
        update_cli(
            "    " + joiner.join(os.path.abspath(str(f)) for f in files), store=store
        )
    update_cli(f"Script executed: {store.start_time}", store=store)
    update_cli("Summary of files parsed:", store=store)
    update_cli(f'{"="*36}', store=store)
    store.remaining = len(files)
    store.total = len(files)
    store.done = 0
    for f in files:
        if stop_event.is_set():
            stop_cli(store)
            return
        try:
            f = os.path.abspath(str(f))
            with Docx(f, triage_files, hash_files, store=store) as doc:
                process_docx(doc, triage_files, hash_files, store)
        except Exception as docxError:
            # If processing a DOCx file raises an error, let the user know, and write it
            # to the error log.
            docxErrorCount += 1  # increment error count by 1.
            update_cli(
                f"Error trying to process {f}. Skipping. Error: {str(docxError)}",
                level="error",
                color=__red__,
                store=store,
            )
            store.errors_worksheet["File Name"].append(f)
            store.errors_worksheet["Error"].append(str(docxError))
        if store.remaining != 0:
            store.remaining -= 1
            store.done += 1
    if store.excel:
        write_to_excel(store.excel_file, store.triage_files, store)
    if store.sqlite:
        write_to_sqlite(store)
    update_cli(f'{"="*24}', store=store)
    if docxErrorCount > 0:
        clr = __red__
    else:
        clr = __clr__
    update_cli(
        f"Processing finished for all files. Errors detected: {docxErrorCount}",
        color=clr,
        store=store,
    )
    if docxErrorCount > 0:
        update_cli("The following files had errors:", "error", color=clr, store=store)
        for each_file in store.errors_worksheet["File Name"]:
            update_cli(f"  {each_file}", "error", color=clr, store=store)
    end_time = dt.now().strftime(__dtfmt__)
    update_cli(f"Script finished execution: {end_time}", color=__green__, store=store)
    run_time = str(
        timedelta(
            seconds=(
                dt.strptime(end_time, __dtfmt__)
                - dt.strptime(store.start_time, __dtfmt__)
            ).seconds
        )
    )
    update_cli(f"Total processing time: {run_time}", color=__green__, store=store)


class ColorFormatter(logging.Formatter):
    def __init__(self, fmt=None, datefmt=None, style="%"):
        super().__init__(fmt, datefmt, style)
        self.color = ""
        self.reset = __clr__

    def set_color(self, color):
        self.color = color

    def format(self, record):
        formatter = logging.Formatter(
            f"{self.color}%(asctime)s | %(levelname)-8s | %(message)s{self.reset}",
            datefmt=__dtfmt__,
        )
        return formatter.format(record)


def cli_log(output_path, verbose=0, store: DataStore = None):
    log = logging.getLogger("ms-word-parser")
    log.setLevel(logging.DEBUG)
    log_fmt = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt=__dtfmt__,
    )
    log_path = os.path.normpath(f"{output_path}{os.sep}{store.log_file}")
    file_handler = logging.FileHandler(log_path, "w", "utf-8")
    file_handler.setFormatter(log_fmt)
    log.addHandler(file_handler)
    verbosity = {0: None, 1: logging.INFO, 2: logging.DEBUG}
    stream_level = verbosity.get(verbose, logging.DEBUG)
    if stream_level is not None:
        store.color_fmt = ColorFormatter()
        stream_handler = logging.StreamHandler(stream=sys.stdout)
        stream_handler.setLevel(stream_level)
        stream_handler.setFormatter(store.color_fmt)
        log.addHandler(stream_handler)
    return log


def stop_cli(store: DataStore):
    update_cli("Processing stopped", store=store)
    if store.excel:
        update_cli("Attempting to write current results to Excel", store=store)
    if store.sqlite:
        update_cli("Attempting to write current results to SQLite", store=store)
    docxErrorCount = len(store.errors_worksheet["Error"])
    try:
        if store.excel:
            write_to_excel(store.excel_file, store.triage_files, store)
        if store.sqlite:
            write_to_sqlite(store)
        if docxErrorCount > 0:
            clr = __red__
        else:
            clr = __clr__
        update_cli(
            f"Finished writing to Excel. Errors detected: {docxErrorCount}",
            color=clr,
            store=store,
        )
        if docxErrorCount > 0:
            update_cli(
                "The following files had errors:", "error", color=clr, store=store
            )
            for each_file in store.errors_worksheet["File Name"]:
                update_cli(f"  {each_file}", "error", color=clr, store=store)
        end_time = dt.now().strftime(__dtfmt__)
        update_cli(
            f"Script finished execution: {end_time}", color=__green__, store=store
        )
        run_time = str(
            timedelta(
                seconds=(
                    dt.strptime(end_time, __dtfmt__)
                    - dt.strptime(store.start_time, __dtfmt__)
                ).seconds
            )
        )
        update_cli(f"Total processing time: {run_time}", color=__green__, store=store)
        return
    except Exception as e:
        update_cli(f"Unable to write results to Excel: {e}", store=store)


def reset_vars(store: DataStore):
    store.reset_vars()


def read_ingest(file):
    all_files = []
    exists = []
    no_files = []
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
    return all_files, no_files


def gui():
    store = DataStore()
    try:
        from ctypes import windll

        app_id = f"jjrboucher.ms-word-parser.gui.v{__version__.replace('.','-')}"
        windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except ImportError:
        pass
    ms_word_app = QApplication([__appname__, "windows:darkmode=2"])
    ms_word_app.setStyle("Universal")
    ms_word_app.setApplicationName(__appname__)
    ms_word_app.setApplicationDisplayName(__appname__)
    style = ms_word_app.style()
    icon = style.standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView)
    ms_word_app.setWindowIcon(icon)
    ms_word_gui = MsWordGui(store)
    store.ms_word_gui = ms_word_gui
    ms_word_gui.show()
    ms_word_app.exec()


def main():
    store = DataStore()
    arg_parse = argparse.ArgumentParser(description=f"MS Word Parser {__version__}")
    arg_parse.add_argument(
        "-e", "--excel", action="store_true", help="outputs data to an Excel document"
    )
    arg_parse.add_argument("-g", "--gui", action="store_true", help="launch the gui")
    arg_parse.add_argument(
        "-H",
        "--hash",
        help="hash the doc zip contents",
        action="store_true",
    )
    arg_parse.add_argument(
        "-o",
        "--output",
        help="output directory",
    )
    arg_parse.add_argument(
        "-r",
        "--recurse",
        action="store_true",
        help="recursively process files in directory",
    )
    arg_parse.add_argument(
        "-s",
        "--sqlite",
        action="store_true",
        help="save data to an sqlite database",
    )
    arg_parse.add_argument(
        "-T",
        "--timeline",
        action="store_true",
        help="produce a timeline view in SQLite or Timeline Sheets in Excel",
    )
    arg_parse.add_argument(
        "-v",
        "--verbose",
        action="count",
        default=0,
        help="Output to STDOUT as well as log, -v: INFO, -vv: DEBUG",
    )
    file_source = arg_parse.add_mutually_exclusive_group(required=False)
    file_source.add_argument(
        "--ingest",
        help="text file with a list of files to ingest",
    )
    file_source.add_argument("--dir", help="directory to process")
    file_source.add_argument(
        "--files", help="individual files to be processed", nargs="*"
    )
    process_mode = arg_parse.add_mutually_exclusive_group(required=False)
    process_mode.add_argument("-t", "--triage", action="store_true", help="triage mode")
    process_mode.add_argument("-f", "--full", action="store_true", help="full mode")

    if len(sys.argv[1:]) == 0:
        arg_parse.print_help()
        arg_parse.exit()

    args = arg_parse.parse_args()
    if args.gui:
        gui()

    if not args.gui:
        if not args.output:
            arg_parse.error("The -o/--output option is required")
        if not os.path.exists(args.output) or not os.path.isdir(args.output):
            arg_parse.error(
                f"The output path {args.output} does not exist. Check your path and try again"
            )
        output_path = os.path.normpath(os.path.abspath(args.output))
        store.output_path = output_path
        if not (args.dir or args.files or args.ingest):
            arg_parse.error("One of --files, --dir, or --ingest is required")
        if not (args.triage or args.full):
            arg_parse.error("One of --triage or --full is required")
        if not args.triage:
            store.triage_files = False
        if not (args.excel or args.sqlite):
            arg_parse.error("One of --excel or --sqlite is required")
        if args.hash:
            store.hash_files = True
        if args.excel or args.sqlite:
            store.logger = cli_log(output_path, verbose=args.verbose, store=store)
            if args.excel:
                store.excel = True
                store.excel_file = f"{output_path}{os.sep}{store.basename}.xlsx"
            if args.sqlite:
                store.sqlite = True
                store.sqlite_file = f"{output_path}{os.sep}{store.basename}.db"
        if args.files:
            file_list = args.files
            store.filenames = file_list
            try:
                process_cli(
                    file_list,
                    store.triage_files,
                    store.hash_files,
                    store,
                )
            except KeyboardInterrupt:
                stop_cli(store)
            except Exception as e:
                update_cli(
                    f"Error trying to process files - {e}",
                    level="error",
                    color=__red__,
                    store=store,
                )
        if args.dir:
            if not os.path.exists(args.dir) or not os.path.isdir(args.dir):
                arg_parse.error(
                    f"The path {args.dir} does not exist. Please check your path and try again."
                )
            folder_path = Path(args.dir)
            if args.recurse:
                files = get_files(folder_path, True)
                file_list = [str(file) for file in files]
            else:
                files = get_files(folder_path, False)
                file_list = [str(file) for file in files]
            store.filenames = file_list
            try:
                process_cli(
                    file_list,
                    store.triage_files,
                    store.hash_files,
                    store,
                )
            except KeyboardInterrupt:
                stop_cli(store)
            except Exception as e:
                update_cli(
                    f"Error trying to process directory - {e}",
                    level="error",
                    color=__red__,
                    store=store,
                )
        if args.ingest:
            if not os.path.exists(args.ingest) or not os.path.isfile(args.ingest):
                arg_parse.error(
                    f"The file {args.ingest} does not exist. Please check your path and try again."
                )
            try:
                process_cli(
                    args.ingest,
                    args.triage,
                    args.hash,
                    store,
                    ingest=True,
                )
            except KeyboardInterrupt:
                stop_cli(store)
            except Exception as e:
                update_cli(
                    f"Error trying to process files - {e}",
                    level="error",
                    color=__red__,
                    store=store,
                )


if __name__ == "__main__":
    main()
