#!/usr/bin/env python3

import hashlib
import os
import sys
import json
import math
import sqlite3
import zipfile
from zipfile import BadZipFile
import logging
import subprocess
import argparse
from datetime import datetime as dt, timedelta
from pathlib import Path
import struct
import xml.etree.ElementTree as ET
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
    from ms_word_parser.tips import (
        tip_sameRsidRoot,
        tip_numDocumentsEachRsidRoot,
        tip_docsCreatedBySameWindowsUser,
        tip_scriptOverview,
        tip_excelWorksheets,
        tip_processingOptions,
        tip_guiWorkFlow,
    )

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
__date__ = "23 Feb 2026"
__author__ = (
    "Jacques Boucher - jjrboucher@gmail.com\nCorey Forman - corey@digitalsleuth.ca"
)
__dtfmt__ = "%Y-%m-%d %H:%M:%S"


class DataStore:
    """Stores the state of all variables for use in multiple functions."""

    def __init__(self):
        """Main data stores"""
        self.doc_summary_worksheet = {}
        self.metadata_worksheet = {}
        self.archive_files_worksheet = {}
        self.rsids_worksheet = {}
        self.comments_worksheet = {}
        self.people_worksheet = {}
        self.extensible_worksheet = {}
        self.extended_worksheet = {}
        self.comments_ids_worksheet = {}
        self.custom_xml_worksheet = {}
        self.item_worksheet = {}
        self.ink_worksheet = {}
        self.timeline_worksheet = {}
        self.aggregated_worksheet = {}
        self.ink_content = []
        self.item_xml_content = None
        self.errors_worksheet = {"File Name": [], "Error": []}
        self.timestamp = dt.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = f"ms-word-parser-log-{self.timestamp}.log"
        self.ms_word_gui = None
        self.start_time = None
        self.color_fmt = None
        self.logger = None
        self.sqlite = False

    def reset_vars(self):
        """Reset variables"""
        self.doc_summary_worksheet = {}
        self.metadata_worksheet = {}
        self.archive_files_worksheet = {}
        self.rsids_worksheet = {}
        self.comments_worksheet = {}
        self.people_worksheet = {}
        self.extensible_worksheet = {}
        self.extended_worksheet = {}
        self.comments_ids_worksheet = {}
        self.custom_xml_worksheet = {}
        self.item_worksheet = {}
        self.ink_worksheet = {}
        self.timeline_worksheet = {}
        self.aggregated_worksheet = {}
        self.ink_content = []
        self.item_xml_content = None
        self.errors_worksheet = {"File Name": [], "Error": []}
        self.timestamp = dt.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = f"ms-word-parser-log-{self.timestamp}.log"
        self.sqlite = False


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
        self.d_width = 1152
        self.d_height = 330
        self.files = []
        self.excel_path = ""
        self.excel_full_path = ""
        self.log_path = ""
        self.log_handler = None
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
        self.actionSelect_Excel = QAction(MainWindow)
        self.actionSelect_Excel.setObjectName("actionSelect_Excel")
        self.actionSelect_Excel.triggered.connect(self.open_excel)
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
        self.centralWidget = QWidget(MainWindow)
        self.centralWidget.setObjectName("centralWidget")
        self.parsingOptions = QGroupBox(self.centralWidget)
        self.parsingOptions.setObjectName("parsingOptions")
        self.parsingOptions.setGeometry(QRect(10, 10, 160, 60))
        self.parsingOptions.setStyleSheet("background: #ffffff; color: black;")
        self.parsingOptions.setFont(self.text_font)
        self.processOptions = QGroupBox(self.centralWidget)
        self.processOptions.setObjectName("processOptions")
        self.processOptions.setGeometry(QRect(180, 10, 180, 60))
        self.processOptions.setStyleSheet("background: #ffffff; color: black;")
        self.processOptions.setFont(self.text_font)
        self.triageButton = QRadioButton(self.parsingOptions)
        self.triageButton.setObjectName("triageButton")
        self.triageButton.setGeometry(QRect(10, 30, 89, 20))
        self.triageButton.setStyleSheet(self.stylesheet)
        self.triageButton.setChecked(True)
        self.triageButton.setFont(self.text_font)
        self.fullButton = QRadioButton(self.parsingOptions)
        self.fullButton.setObjectName("fullButton")
        self.fullButton.setGeometry(QRect(88, 30, 60, 20))
        self.fullButton.setStyleSheet(self.stylesheet)
        self.fullButton.setFont(self.text_font)
        self.hashFiles = QCheckBox(self.processOptions)
        self.hashFiles.setObjectName("hashFiles")
        self.hashFiles.setGeometry(QRect(10, 30, 89, 20))
        self.hashFiles.setStyleSheet(self.stylesheet)
        self.hashFiles.setFont(self.text_font)
        self.sqliteButton = QCheckBox(self.processOptions)
        self.sqliteButton.setObjectName("sqliteButton")
        self.sqliteButton.setGeometry(QRect(90, 30, 89, 20))
        self.sqliteButton.setStyleSheet(self.stylesheet)
        self.sqliteButton.setFont(self.text_font)
        self.outputFiles = QGroupBox(self.centralWidget)
        self.outputFiles.setObjectName("outputFiles")
        self.outputFiles.setGeometry(QRect(10, 76, 350, 120))
        self.outputFiles.setStyleSheet("background-color: #ffffff; color: black;")
        self.outputFiles.setFont(self.text_font)
        self.excelFileLabel = QLabel(self.outputFiles)
        self.excelFileLabel.setObjectName("excelFileLabel")
        self.excelFileLabel.setGeometry(QRect(10, 30, 80, 16))
        self.excelFileLabel.setStyleSheet("background: #ffffff; color: black;")
        self.excelFileLabel.setFont(self.text_font)
        self.excelFileText = "File -> Select Excel or click 'Select Excel'"
        self.excelFile = QTextEdit(self.outputFiles)
        self.excelFile.setObjectName("excelFile")
        self.excelFile.setGeometry(QRect(92, 26, 250, 26))
        self.excelFile.setAlignment(
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft
        )
        self.excelFile.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.excelFile.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.excelFile.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.excelFile.setFont(self.text_font)
        self.excelFile.setReadOnly(True)
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
        self.generalLogFile.setGeometry(QRect(92, 58, 250, 26))
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
        self.outputPathLabel = QLabel(self.outputFiles)
        self.outputPathLabel.setObjectName("outputPathLabel")
        self.outputPathLabel.setGeometry(QRect(10, 92, 80, 16))
        self.outputPathLabel.setStyleSheet("background: #ffffff; color: black;")
        self.outputPathLabel.setFont(self.text_font)
        self.outputPath = QTextEdit(self.outputFiles)
        self.outputPath.setAlignment(
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft
        )
        self.outputPath.setObjectName("outputPath")
        self.outputPath.setGeometry(QRect(92, 88, 250, 26))
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

        self.operationOptions = QGroupBox(self.centralWidget)
        self.operationOptions.setObjectName("operationOptions")
        self.operationOptions.setGeometry(QRect(10, 200, 350, 90))
        self.operationOptions.setStyleSheet("background-color: #ffffff; color:black;")
        self.operationOptions.setFont(self.text_font)
        self.excelButton = QPushButton(self.operationOptions)
        self.excelButton.setObjectName("excelButton")
        self.excelButton.setGeometry(QRect(10, 28, 86, 24))
        self.excelButton.setStyleSheet(self.stylesheet)
        self.excelButton.clicked.connect(self.open_excel)
        self.excelButton.setFont(self.text_font)
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
        self.processButton = QPushButton(self.operationOptions)
        self.processButton.setObjectName("processButton")
        self.processButton.setGeometry(QRect(10, 58, 86, 24))
        self.processButton.setEnabled(False)
        self.processButton.setStyleSheet(self.disabled)
        self.processButton.clicked.connect(
            lambda: self.analyze_docs(
                self.files,
                self.triageButton.isChecked(),
                self.hashFiles.isChecked(),
                self.sqliteButton.isChecked(),
            )
        )
        self.processButton.setFont(self.text_font)
        self.stopButton = QPushButton(self.operationOptions)
        self.stopButton.setObjectName("stopButton")
        self.stopButton.setGeometry(QRect(112, 58, 86, 24))
        self.stopButton.setEnabled(False)
        self.stopButton.setStyleSheet(self.disabled)
        self.stopButton.clicked.connect(self._stop)
        self.stopButton.setFont(self.text_font)
        self.resetButton = QPushButton(self.operationOptions)
        self.resetButton.setObjectName("resetButton")
        self.resetButton.setGeometry(QRect(214, 58, 86, 24))
        self.resetButton.clicked.connect(self._reset)
        self.resetButton.setStyleSheet(self.stylesheet)
        self.resetButton.setFont(self.text_font)
        self.processStatus = QGroupBox(self.centralWidget)
        self.processStatus.setObjectName("processStatus")
        self.processStatus.setGeometry(QRect(370, 10, 768, 280))
        self.processStatus.setStyleSheet("background: #ffffff; color: black;")
        self.processStatus.setFont(self.text_font)
        self.docxOutput = QTextEdit(self.processStatus)
        self.docxOutput.setObjectName("docxOutput")
        self.docxOutput.setGeometry(QRect(16, 60, 737, 210))
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
        self.openLogButton.setGeometry(QRect(402, 29, 110, 24))
        self.openLogButton.setFont(self.text_font)
        self.openLogButton.setStyleSheet(self.disabled)
        self.openLogButton.setEnabled(False)
        self.openLogButton.clicked.connect(lambda: self.open_file(self.log_path))
        self.openExcelButton = QPushButton(self.processStatus)
        self.openExcelButton.setObjectName("openExcelButton")
        self.openExcelButton.setGeometry(QRect(522, 29, 110, 24))
        self.openExcelButton.setFont(self.text_font)
        self.openExcelButton.setStyleSheet(self.disabled)
        self.openExcelButton.setEnabled(False)
        self.openExcelButton.clicked.connect(
            lambda: self.open_file(self.excel_full_path)
        )
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

        self.menubar.addAction(self.menuFile.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())
        self.menuFile.addAction(self.actionSelect_Excel)
        self.menuFile.addSeparator()
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
        self.actionSelect_Excel.setText(
            QCoreApplication.translate("MainWindow", "Select &Excel ...", None)
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
        self.parsingOptions.setTitle(
            QCoreApplication.translate("MainWindow", "Parsing Options", None)
        )
        self.processOptions.setTitle(
            QCoreApplication.translate("MainWindow", "Processing Options", None)
        )
        self.triageButton.setText(
            QCoreApplication.translate("MainWindow", "Triage", None)
        )
        self.fullButton.setText(QCoreApplication.translate("MainWindow", "Full", None))
        self.hashFiles.setText(
            QCoreApplication.translate("MainWindow", "Hash Files", None)
        )
        self.sqliteButton.setText(
            QCoreApplication.translate("MainWindow", "SQLite DB", None)
        )
        self.outputFiles.setTitle(
            QCoreApplication.translate("MainWindow", "Output Files", None)
        )
        self.excelFile.setText(
            QCoreApplication.translate("MainWindow", self.excelFileText, None)
        )
        self.excelFileLabel.setText(
            QCoreApplication.translate("MainWindow", "Excel File:", None)
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
        self.resetButton.setText(
            QCoreApplication.translate("MainWindow", "Reset", None)
        )
        self.excelButton.setText(
            QCoreApplication.translate("MainWindow", "Select Excel", None)
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
        self.openExcelButton.setText(
            QCoreApplication.translate("MainWindow", "Open Excel File", None)
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
                if self.excelFile.toPlainText() != self.excelFileText:
                    self.processButton.setEnabled(True)
                    self.processButton.setStyleSheet(self.stylesheet)
                self.files = files
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
            if self.excelFile.toPlainText() != self.excelFileText:
                self.processButton.setEnabled(True)
                self.processButton.setStyleSheet(self.stylesheet)
            self.files = all_files

    def open_excel(self):
        excel_full_path, _ = QFileDialog.getSaveFileName(
            self, "Select an Excel document ...", "", "Excel Files (*.xlsx)"
        )
        if excel_full_path:
            self.excel_path = os.path.normpath(os.path.dirname(excel_full_path))
            self.log_path = os.path.normpath(
                f"{self.excel_path}{os.sep}{self.store.log_file}"
            )
            self.log_handler = logging.FileHandler(self.log_path, "w", "utf-8")
            self.log_handler.setFormatter(self.log_fmt)
            self.logger.addHandler(self.log_handler)
            update_status = self.update_status
            update_status(f"{__appname__}")
            if not excel_full_path.endswith(".xlsx"):
                excel_full_path += ".xlsx"
            excel_full_path = os.path.normpath(excel_full_path)
            self.excel_full_path = excel_full_path
            excel_file = os.path.basename(excel_full_path)
            update_status(f"Output File Path: {self.excel_path}")
            update_status(f"Excel output file: {excel_file}")
            update_status(f"Log file: {self.log_path}")
            self.excelFile.setText(excel_file)
            if self.numOfFiles.toPlainText() != "0":
                self.processButton.setEnabled(True)
                self.processButton.setStyleSheet(self.stylesheet)
            self.actionAdd_Files.setVisible(True)
            self.actionAdd_Directory.setVisible(True)
            self.generalLogFile.setText(self.store.log_file)
            self.outputPath.setText(self.excel_path)
            self.openButton.setEnabled(True)
            self.openButton.setStyleSheet(self.stylesheet)
            self.addFilesButton.setEnabled(True)
            self.addFilesButton.setStyleSheet(self.stylesheet)
            self.addDirectoryButton.setEnabled(True)
            self.addDirectoryButton.setStyleSheet(self.stylesheet)

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

    def _reset(self):
        self.store.reset_vars()
        self.excelFile.setText(self.excelFileText)
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
        self.openExcelButton.setEnabled(False)
        self.openExcelButton.setStyleSheet(self.disabled)
        self.openButton.setEnabled(False)
        self.openButton.setStyleSheet(self.disabled)
        self.actionAdd_Files.setVisible(False)
        self.actionAdd_Directory.setVisible(False)
        self.triageButton.setChecked(True)
        self.addFilesButton.setEnabled(False)
        self.addFilesButton.setStyleSheet(self.disabled)
        self.addDirectoryButton.setEnabled(False)
        self.addDirectoryButton.setStyleSheet(self.disabled)
        self.hashFiles.setChecked(False)
        self.sqliteButton.setChecked(False)
        self.stopButton.setEnabled(False)
        self.stopButton.setStyleSheet(self.disabled)

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

    def analyze_docs(self, files, triage_files, hash_files, sqlite_output):
        if not self.running:
            self.running = True
        start_time = dt.now().strftime(__dtfmt__)
        self.store.start_time = start_time
        self.store.sqlite = sqlite_output
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
                    write_to_excel(self.excel_full_path, triage_files, store=self.store)
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
                    self.openExcelButton.setStyleSheet(self.stylesheet)
                    self.openExcelButton.setEnabled(True)
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
                    f"Error trying to process {f}. Skipping. Error: {docxError}",
                    level="error",
                    color=red,
                )
                self.store.errors_worksheet["File Name"].append(f)
                self.store.errors_worksheet["Error"].append(docxError)
            if remaining != 0:
                remaining -= 1
            self.numRemaining.setText(str(remaining))
        write_to_excel(self.excel_full_path, triage_files, store=self.store)
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
        reset_vars(self.store)
        self.resetButton.setEnabled(True)
        self.resetButton.setStyleSheet(self.stylesheet)
        self.stopButton.setEnabled(False)
        self.stopButton.setStyleSheet(self.disabled)
        self.openLogButton.setEnabled(True)
        self.openLogButton.setStyleSheet(self.stylesheet)
        self.openExcelButton.setStyleSheet(self.stylesheet)
        self.openExcelButton.setEnabled(True)


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


class Docx:
    """
    Accepts a docx file. Has the following methods to extract data from core.xml, app.xml, document.xml

    app_version, application, category, characters, characters_with_spaces, company, content_status, created, creator,
    description, filename, keywords, last_modified_by, last_printed, lines, manager, modified, pages, paragraph_tags,
    paragraphs, revision, runs_tags, security, subject, template, text_tags, title, total_editing_time, words,
    xml_files, xml_hash, xml_size
    """

    def __init__(
        self, msword_file, triage=False, hashing=True, store: DataStore = None
    ):
        """
        .docx file to pass to the class
        Triage value can be True or False. If True, will parse less info to execute faster.
        When set to False, it does not try to parse RSID values from document.xml.
        If triage value not passed, it defaults to False and does full parsing.
        The script using this class still ultimately decides what methods it wants to use.
        But if in triage mode, some of the variables will not get assigned any value, thus
        will affect any methods that rely on those variables having a value assigned to them.
        """
        if store is None:
            store = DataStore()
        self.store = store
        if store.ms_word_gui:
            update_status = store.ms_word_gui.update_status
        else:
            update_status = lambda msg, **kwargs: update_cli(msg, store=store, **kwargs)
        self.update_status = update_status
        self.item_files = []
        self.ink_files = []
        self.namespaces = {
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "aink": "http://schemas.microsoft.com/office/drawing/2016/ink",
            "b": "http://schemas.openxmlformats.org/officeDocument/2006/bibliography",
            "ct": "http://schemas.microsoft.com/office/2006/metadata/contentType",
            "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            "cprop": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
            "cr": "http://schemas.microsoft.com/office/comments/2020/reactions",
            "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
            "dc": "http://purl.org/dc/elements/1.1/",
            "dcterms": "http://purl.org/dc/terms/",
            "dcmitype": "http://purl.org/dc/dcmitype/",
            "default": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
            "ds": "http://schemas.openxmlformats.org/officeDocument/2006/customXml",
            "inkml": "http://www.w3.org/2003/InkML",
            "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "ma": "http://schemas.microsoft.com/office/2006/metadata/properties/metaAttributes",
            "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
            "o": "urn:schemas-microsoft-com:office:office",
            "oel": "http://schemas.microsoft.com/office/2019/extlst",
            "p": "http://schemas.microsoft.com/office/2006/metadata/properties",
            "pc": "http://schemas.microsoft.com/office/infopath/2007/PartnerControls",
            "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "sc": "Microsoft.SharePoint.Taxonomy.ContentTypeSync",
            "sp": "http://schemas.microsoft.com/sharepoint/v3",
            "v": "urn:schemas-microsoft-com:vml",
            "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
            "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
            "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
            "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
            "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
            "w16du": "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
            "w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
            "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
            "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
            "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
            "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
            "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
            "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
            "xs": "http://www.w3.org/2001/XMLSchema",
            "xsd": "http://www.w3.org/2001/XMLSchema",
            "xsi": "http://www.w3.org/2001/XMLSchema-instance",
        }
        self.has_ink = False
        self.has_comments = False
        self.msword_file = msword_file
        self.hashing = hashing
        self.header_offsets, self.binary_content = self.__find_binary_string()
        self.extra_fields = self.__xml_extra_bytes()
        self.__load_all_xml()
        self.rsidRs = self.__extract_all_rsids_from_settings_xml()
        self.ns_lookup = {
            "title": [self.core_xml_content, "dc"],
            "subject": [self.core_xml_content, "dc"],
            "creator": [self.core_xml_content, "dc"],
            "keywords": [self.core_xml_content, "cp"],
            "description": [self.core_xml_content, "dc"],
            "revision": [self.core_xml_content, "cp"],
            "created": [self.core_xml_content, "dcterms"],
            "modified": [self.core_xml_content, "dcterms"],
            "lastModifiedBy": [self.core_xml_content, "cp"],
            "lastPrinted": [self.core_xml_content, "cp"],
            "category": [self.core_xml_content, "cp"],
            "contentStatus": [self.core_xml_content, "cp"],
            "language": [self.core_xml_content, "dc"],
            "version": [self.core_xml_content, "cp"],
            "Template": [self.app_xml_content, "default"],
            "TotalTime": [self.app_xml_content, "default"],
            "Pages": [self.app_xml_content, "default"],
            "Words": [self.app_xml_content, "default"],
            "Characters": [self.app_xml_content, "default"],
            "Application": [self.app_xml_content, "default"],
            "DocSecurity": [self.app_xml_content, "default"],
            "Lines": [self.app_xml_content, "default"],
            "Paragraphs": [self.app_xml_content, "default"],
            "CharactersWithSpaces": [self.app_xml_content, "default"],
            "AppVersion": [self.app_xml_content, "default"],
            "Manager": [self.app_xml_content, "default"],
            "Company": [self.app_xml_content, "default"],
            "SharedDoc": [self.app_xml_content, "default"],
            "HyperlinksChanged": [self.app_xml_content, "default"],
        }
        x = ET.fromstring(self.document_xml_content)
        self.p_tags = x.findall(".//w:p", self.namespaces)
        self.r_tags = x.findall(".//w:r", self.namespaces)
        self.t_tags = x.findall(".//w:t", self.namespaces)
        self.tr_tags = x.findall(".//w:tr", self.namespaces)
        self.shapedata = x.findall(".//v:shape", self.namespaces)
        self.drawing_tags = x.findall(".//w:drawing", self.namespaces)
        if self.drawing_tags or self.ink_files:
            self.has_ink = True
        if not triage:  # if not run in triage mode, do full parsing
            self.rsidR_in_document_xml = self.__rsids_in_document_xml("rsidR")
            self.rsidRPr = self.__rsids_in_document_xml("rsidRPr")
            self.rsidP = self.__rsids_in_document_xml("rsidP")
            self.rsidRDefault = self.__rsids_in_document_xml("rsidRDefault")
            self.rsidTr = self.__rsids_in_document_xml("rsidTr")
            self.para_id = self.__rsids_in_document_xml("paraId")
            self.text_id = self.__rsids_in_document_xml("textId")

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.core_xml_content = None
        self.app_xml_content = None
        self.document_xml_content = None
        self.comments_xml_content = None
        self.settings_xml_content = None
        self.people_xml_content = None
        self.extensible_xml_content = None
        self.extended_xml_content = None
        self.comments_ids_content = None
        self.custom_xml_content = None

    def __find_binary_string(self):

        pkzip_header = b"PK\x03\x04"
        with open(self.msword_file, "rb") as msword_binary:  # read the file as binary
            content = msword_binary.read()
        matches = []  # list of offsets where header is found
        index = 0

        while index < len(content):  # iterate over the list
            index = content.find(pkzip_header, index)  # search for
            if index == -1:  # no more items in the list.
                break
            matches.append(index)
            index += 1

        return (
            matches,
            content,
        )  # returns the list of offsets of each header, and the binary file.

    def __xml_extra_bytes(self):
        """
        ref: https://en.wikipedia.org/wiki/ZIP_(file_format)#Local_file_header

        return: list [xml file name, # of bytes in extra field, truncated bytes]
        """
        filename = ""
        extras = {}
        truncate_extra_field = 20  # extra field can be several hundred bytes, mostly 0x00. This grabs the first 20.

        for offset in self.header_offsets:
            (
                filename_len,
                extrafield_len,
            ) = struct.unpack("<2H", self.binary_content[offset + 26 : offset + 30])
            filename_start = offset + 30
            filename_end = offset + 30 + filename_len
            if filename_end - filename_start < 256:
                # some DOCx files somehow produce false positives of
                # excessively long filenames and results in an error. This avoids that error.
                filename = self.binary_content[filename_start:filename_end].decode(
                    "ascii"
                )
            extrafield_start = filename_end
            extrafield_end = extrafield_start + extrafield_len
            extrafield = self.binary_content[extrafield_start:extrafield_end]
            extrafield_hex_as_text = []

            for h in extrafield:
                extrafield_hex_as_text.append(f"{h:02x}")

            if not extrafield:
                extras[filename] = [extrafield_len, "nil"]
            elif (
                extrafield_len <= truncate_extra_field
            ):  # field size larger than truncate value
                extras[filename] = [
                    extrafield_len,
                    f"0x{''.join(extrafield_hex_as_text)}",
                ]
            else:
                extras[filename] = [
                    extrafield_len,
                    f"0x{''.join(extrafield_hex_as_text[0:truncate_extra_field])}",
                ]  # adds only
                # the select # of characters as specified in the variable truncate_extra_field. This is so that
                # we don't end up with hundreds of characters in a cell in Excel, as some extra fields can be
                # several hundred values long. But so far, most are 0x00, with only the first few being values other
                # than hex 0x00.

        return extras

    def __load_xml(self, xml_file):
        content = ""
        if (
            xml_file in self.xml_files()
        ):  # if the file exists, read it and return its content
            if "comments.xml" in xml_file:
                self.has_comments = True
            with zipfile.ZipFile(self.msword_file, "r") as zipref:
                with zipref.open(xml_file) as xmlFile:
                    content = xmlFile.read()
        else:
            self.update_status(
                f'"{xml_file}" does not exist in "{self.msword_file}". '
                f"Returning empty string.",
                level="debug",
            )
        return content

    def __load_all_xml(self):
        xml_map = {
            "core_xml_content": "docProps/core.xml",
            "app_xml_content": "docProps/app.xml",
            "document_xml_content": "word/document.xml",
            "comments_xml_content": "word/comments.xml",
            "settings_xml_content": "word/settings.xml",
            "people_xml_content": "word/people.xml",
            "extensible_xml_content": "word/commentsExtensible.xml",
            "extended_xml_content": "word/commentsExtended.xml",
            "comments_ids_content": "word/commentsIds.xml",
            "custom_xml_content": "docProps/custom.xml",
        }
        try:
            with zipfile.ZipFile(self.msword_file, "r") as zipref:
                zip_filenames = zipref.namelist()
                for attrib, file_path in xml_map.items():
                    alt_path = file_path.replace("/", "\\")
                    target = file_path if file_path in zip_filenames else alt_path
                    if target in zip_filenames:
                        if "comments.xml" in target:
                            self.has_comments = True
                        content = zipref.read(target)
                        setattr(self, attrib, content)
                    else:
                        setattr(self, attrib, "")
                        self.update_status(
                            f'"{target}" does not exist in "{self.msword_file}". '
                            f"Returning empty string.",
                            level="debug",
                        )
        except (zipfile.BadZipFile, FileNotFoundError) as e:
            raise Exception(f"Error accessing {self.msword_file}: {e}") from e

    def get_metadata(self, attrib):
        """
        :param: xmlcontent (self.core_xml_content or self.app_xml_content)
        :param: attrib (the attribute in the content to get)
        :return:
        """
        xmlcontent = self.ns_lookup[attrib][0]
        ns = self.namespaces[self.ns_lookup[attrib][1]]
        if xmlcontent:
            content = ET.fromstring(xmlcontent)
            ns_extract = content.find(f"{{{ns}}}{attrib}")
            meta_content = ns_extract.text if ns_extract is not None else None
        else:
            return None
        return meta_content

    def get_people(self):
        if self.people_xml_content != "":
            xml = ET.fromstring(self.people_xml_content)
            list_of_people = []
            all_people = xml.findall(".//w15:person", self.namespaces)
            for person in all_people:
                author = person.get(f"{{{self.namespaces['w15']}}}author")
                if len(person) > 0:
                    providerId = person[0].get(
                        f"{{{self.namespaces['w15']}}}providerId"
                    )
                    userId = person[0].get(f"{{{self.namespaces['w15']}}}userId")
                else:
                    providerId = userId = None
                list_of_people.append([author, providerId, userId])
            return list_of_people
        return None

    def any_comments(self):
        return self.has_comments

    def get_comments(self):
        """
        return the list all_comments that contains the following:
            comment ID #,
            Timestamp,
            Author,
            Initials,
            Text
        :return:
        """

        if not self.has_comments:
            return [None, None, None, None, None]
        xml = ET.fromstring(self.comments_xml_content)
        # Find all comments
        comments = xml.findall(".//w:comment", self.namespaces)
        all_comments = []
        for comment in comments:
            author = comment.get(f"{{{self.namespaces['w']}}}author")
            date_time = comment.get(f"{{{self.namespaces['w']}}}date")
            initials = comment.get(f"{{{self.namespaces['w']}}}initials")
            comment_id = comment.get(f"{{{self.namespaces['w']}}}id")
            comment_paras = comment.findall(".//w:p", self.namespaces)
            text = (
                "\n".join(
                    [
                        t.text
                        for t in comment.findall(".//w:t", self.namespaces)
                        if t.text
                    ]
                )
                .encode("utf-8", "surrogatepass")
                .decode()
            )
            if len(comment_paras) > 0:
                comment_paraId = comment_paras[-1].get(
                    f"{{{self.namespaces['w14']}}}paraId"
                )
            else:
                comment_paraId = None
            all_comments.append(
                [comment_id, comment_paraId, date_time, author, initials, text]
            )
        return all_comments

    def get_comments_ids(self):
        if self.comments_ids_content != "":
            all_comments_ids = []
            xml = ET.fromstring(self.comments_ids_content)
            comments_ids = xml.findall(".//w16cid:commentId", self.namespaces)
            for comment_id in comments_ids:
                paraId = comment_id.get(f"{{{self.namespaces['w16cid']}}}paraId", "")
                durableId = comment_id.get(
                    f"{{{self.namespaces['w16cid']}}}durableId", ""
                )
                all_comments_ids.append([paraId, durableId])
            return all_comments_ids
        return None

    def get_extended_comments(self):
        if self.extended_xml_content != "":
            all_extended_comments = []
            xml = ET.fromstring(self.extended_xml_content)
            extended_comments = xml.findall(".//w15:commentEx", self.namespaces)
            for values in extended_comments:
                paraId = values.get(f"{{{self.namespaces['w15']}}}paraId")
                done = values.get(f"{{{self.namespaces['w15']}}}done")
                paraIdParent = values.get(
                    f"{{{self.namespaces['w15']}}}paraIdParent", "IS_PARENT"
                )
                all_extended_comments.append([paraId, paraIdParent, done])
            return all_extended_comments
        return None

    def get_extensible_comments(self):
        if self.extensible_xml_content != "":
            all_extensible_comments = {}
            xml = ET.fromstring(self.extensible_xml_content)
            extensible_comments = xml.findall(
                ".//w16cex:commentExtensible", self.namespaces
            )
            reaction_types = {0: "Unknown", 1: "Like", 2: "Unknown"}
            for values in extensible_comments:
                uri = "None"
                reactionType = "None"
                userId = userProvider = userName = ""
                durableId = values.get(f"{{{self.namespaces['w16cex']}}}durableId")
                dateUtc = values.get(f"{{{self.namespaces['w16cex']}}}dateUtc")
                extLst = values.findall(".//w16cex:extLst", self.namespaces)
                all_extensible_comments[durableId] = []
                all_extensible_comments[durableId].append(dateUtc)
                if extLst:
                    ext = extLst[0].find("w16:ext", self.namespaces)
                    uri = ext.get(f"{{{self.namespaces['w16']}}}uri")
                    all_extensible_comments[durableId].append(uri)
                    for entry in ext.findall(".//cr:reaction", self.namespaces):
                        reactionType = entry.get("reactionType", "")
                        all_extensible_comments[durableId].append(
                            reaction_types[int(reactionType)]
                        )
                        for reactionInfo in entry.findall(
                            ".//cr:reactionInfo", self.namespaces
                        ):
                            reactionDateUtc = reactionInfo.get("dateUtc", "")
                            user = reactionInfo.find("cr:user", self.namespaces)
                            if user is not None:
                                userId = user.get("userId", "")
                                userProvider = user.get("userProvider", "")
                                userName = user.get("userName", "")
                            all_extensible_comments[durableId].append(
                                [reactionDateUtc, userId, userProvider, userName]
                            )
                else:
                    all_extensible_comments[durableId].append(uri)
                    all_extensible_comments[durableId].append(reactionType)
                    all_extensible_comments[durableId].append(["", "", "", ""])
            return all_extensible_comments
        return None

    def __extract_all_rsids_from_settings_xml(self):
        """
        function to extract all RSIDs at the beginning of the class.
        :return:
        """
        rsids = []
        x = ET.fromstring(self.settings_xml_content)
        rsid_tags = x.findall(".//w:rsid", self.namespaces)
        for tag in rsid_tags:
            rsid_tag = tag.get(f"{{{self.namespaces['w']}}}val", None)
            if rsid_tag:
                rsids.append(rsid_tag)
        return "" if not rsids else rsids

    def __rsids_in_document_xml(self, rsid):
        """
        :param rsid tag name (e.g. "rsidRPr", "rsidP", "rsidRDefault")
        The function accepts an rsid tag name as a parameter (e.g. rsidRPr, rsidP, rsidDefault).
        It searches document.xml for a pattern to find all instances of that rsid tag.
        It creates a dictionary that contains each unique rsid value as the key, and the count of how many times
        that rsid is in document.xml.
        E.g., {"00123456": 4, "00234567": 0, "00345678":11}

        :return: dictionary where the key is unique RSIDs, and the value is a count of the occurrences of that rsid
        in document.xml
        """
        rsids = {}
        all_rsids = []
        ns_list = {
            "rsidR": self.namespaces["w"],
            "rsidRDefault": self.namespaces["w"],
            "rsidRPr": self.namespaces["w"],
            "rsidP": self.namespaces["w"],
            "rsidTr": self.namespaces["w"],
            "paraId": self.namespaces["w14"],
            "textId": self.namespaces["w14"],
        }
        for entry in (self.p_tags, self.r_tags, self.t_tags, self.tr_tags):
            for item in entry:
                other_rsid = item.get(f"{{{ns_list[rsid]}}}{rsid}", None)
                if other_rsid:
                    all_rsids.append(other_rsid)
        unique_rsids = set(all_rsids)
        if rsid == "rsidR":
            for each in self.rsidRs:
                rsids[each] = all_rsids.count(each)
        else:
            for each_rsid in unique_rsids:
                rsids[each_rsid] = all_rsids.count(each_rsid)
        return rsids

    def hyperlinks(self):
        """
        :return: Hyperlink values in document.xml
        """
        doc_hyperlinks = []
        doc = ET.fromstring(self.document_xml_content)
        for hyperlink in doc.findall(f".//{{{self.namespaces['w']}}}hyperlink"):
            link_text = hyperlink.findall(f".//{{{self.namespaces['w']}}}t")
            hyperlinks = ",".join(link.text for link in link_text if link.text)
            hyperlinks = hyperlinks.replace("http", "hxxp")
            rel_id = hyperlink.get(f"{{{self.namespaces['r']}}}id", "")
            doc_hyperlinks.append([hyperlinks, rel_id])
        all_hyperlinks = "|".join(f"{url}: {rel}" for url, rel in doc_hyperlinks)
        return all_hyperlinks

    def filename(self):
        """
        :return: the filename of the DOCx file passed to the class
        """
        return self.msword_file

    def hash(self, content=None):
        """
        Function that will return the hash of the file itself
        """
        if self.hashing:  # if hashing option was selected
            filehash = hashlib.md5()
            if content is None:
                filehash.update(self.binary_content)
            else:
                filehash.update(content)
            return filehash.hexdigest().upper()
        return ""  # if no hashing was selected.

    def xml_files(self):
        """
        :return: A dictionary in the following format:
        {XML filename: [file hash,
                        modified date,
                        file size,
                        ZIP compression type,
                        ZIP Create System,
                        ZIP Created Version,
                        ZIP Extract Version,
                        ZIP Flag Bits (hex),
                        ZIP extra values (hex as text)
        }
        """
        compression_types = {0: "Store (None)", 8: "DEFLATE"}
        with zipfile.ZipFile(self.msword_file, "r") as zip_file:
            xml_files = {}
            for file_info in zip_file.infolist():
                if (
                    "customXml/item" in file_info.filename
                    and "Props" not in file_info.filename
                    and file_info.filename not in self.item_files
                ):
                    self.item_files.append(file_info.filename)
                if (
                    "ink/ink" in file_info.filename
                    and file_info.filename not in self.ink_files
                ):
                    self.ink_files.append(file_info.filename)
                with zipfile.ZipFile(self.msword_file, "r") as zip_ref:
                    try:
                        with zip_ref.open(file_info.filename) as xml_file:
                            if self.hashing:  # if hashing option selected
                                md5hash = self.hash(xml_file.read())
                            else:
                                md5hash = "Option Not Selected"  # else return blank for hash value.
                    except BadZipFile:
                        pass
                    except OSError as exc:
                        raise SystemError(
                            "Error processing the zip file header - likely offset is incorrect."
                        ) from exc
                m_time = file_info.date_time
                if m_time in ((1980, 1, 1, 0, 0, 0), (1980, 0, 0, 0, 0, 0)):
                    modified_time = None
                else:
                    modified_time = dt(*m_time).strftime(__dtfmt__)
                fname = file_info.filename
                if fname not in self.extra_fields:
                    fname = fname.replace("/", "\\")
                xml_files[file_info.filename] = [
                    md5hash,
                    modified_time,
                    file_info.file_size,
                    f'{str(file_info.compress_type)}: {compression_types.get(file_info.compress_type, "Unidentified")}',
                    file_info.create_system,
                    file_info.create_version,
                    file_info.extract_version,
                    f"{file_info.flag_bits:#0{6}x}",
                    self.extra_fields[fname][0],
                    self.extra_fields[fname][1],
                ]
            return xml_files

    def xml_hash(self, xmlfile: str):
        """
        :param: xmlfile
        :return: the hash of a specified XML file
        """
        return self.xml_files()[xmlfile][1]

    def xml_size(self, xmlfile: str):
        """
        :param: xmlfile
        :return: the size of a specified XML file
        """
        return self.xml_files()[xmlfile][0]

    def paragraph_tags(self):
        """
        :return: the total number of paragraph tags in document.xml
        """
        return len(self.p_tags)

    def runs_tags(self):
        """
        :return: the total number of runs tags in document.xml
        """
        return len(self.r_tags)

    def text_tags(self):
        """
        :return: the total number of text tags in document.xml
        """
        return len(self.t_tags)

    def table_row_tags(self):
        """
        :return: the total number of table row tags in document.xml
        """
        return len(self.tr_tags)

    def rsid_root(self):
        """
        :return: rsidRoot from settings.xml
        """
        x = ET.fromstring(self.settings_xml_content)
        rsid_root_entry = x.findall(".//w:rsidRoot", self.namespaces)
        root = None
        for entry in [rsid_root_entry]:
            for item in entry:
                root = item.get(
                    f"{{{self.namespaces['w']}}}val",
                    None,
                )
        return None if root is None else root

    def get_doc_ids(self):
        """
        :return: the w14, w15, and w16 docId's from settings.xml
        """
        x = ET.fromstring(self.settings_xml_content)
        w14_id = w15_id = w16_id = "None"
        w14_ns = x.find(f"{{{self.namespaces['w14']}}}docId")
        if w14_ns is not None:
            w14_id = w14_ns.get(f"{{{self.namespaces['w14']}}}val", "None")
        w15_ns = x.find(f"{{{self.namespaces['w15']}}}docId")
        if w15_ns is not None:
            w15_id = w15_ns.get(f"{{{self.namespaces['w15']}}}val", "None")
        w16_ns = x.find(f"{{{self.namespaces['w16']}}}docId")
        if w16_ns is not None:
            w16_id = w16_ns.get(f"{{{self.namespaces['w16']}}}val", "None")

        return [w14_id, w15_id, w16_id]

    def rsidr(self):
        """
        :return: a list containing all the rsidR in settings.xml
        Not all of these will necessarily still be in the document. If all text from a particular revision/save
        session is deleted, the associated rsidR will no longer be found in the document. Thus, the absence
        of an rsidR lets you know that all the data from that editing session has been deleted from the document.

        Because there are no duplicate rsidR values in settings.xml (as long as you don't also grab rsidRoot),
        there is no need for the method to deduplicate.
        """
        return self.rsidRs

    def rsidr_in_document_xml(self):
        """
        return dictionary with unique rsidR and count of how many times it is found in document.xml
        :return:
        """
        return self.rsidR_in_document_xml

    def rsidrpr_in_document_xml(self):
        """
        return dictionary with unique rsidRPr and count of how many times it is found in document.xml
        :return:
        """
        return self.rsidRPr

    def rsidp_in_document_xml(self):
        """
        return dictionary with unique rsidP and count of how many times it is found in document.xml
        :return:
        """
        return self.rsidP

    def rsidrdefault_in_document_xml(self):
        """
        return dictionary with unique rsidRDefault and count of how many times it is found in document.xml
        :return:
        """
        return self.rsidRDefault

    def rsidtr_in_document_xml(self):
        """
        return dictionary with unique rsidTr and count of how many times it is found in document.xml
        :return:
        """
        return self.rsidTr

    def paragraph_id_tags(self):
        return self.para_id

    def text_id_tags(self):
        return self.text_id

    def details(self):
        """
        :return: a text string that you can print out to get a summary of the document.
        This can be edited to suit your needs. You can naturally accomplish the same results by calling each of
        the methods in your print statement in the main script.
        """
        if self.get_metadata("lastPrinted") == "":
            printed = "Document was never printed"
        else:
            printed = f"Printed: {self.get_metadata('lastPrinted')}"
        return (
            f"Document: {self.filename()}\n"
            f"Created by: {self.get_metadata('creator')}\n"
            f"Created date: {self.get_metadata('created')}\n"
            f"Last edited by: {self.get_metadata('lastModifiedBy')}\n"
            f"Edited date: {self.get_metadata('modified')}\n"
            f"{printed}\n"
            f"Total pages: {self.get_metadata('Pages')}\n"
            f"Total editing time: {self.get_metadata('TotalTime')} minute(s)."
        )

    def get_proof_state(self):
        xml = ET.fromstring(self.settings_xml_content)
        proof_state = xml.find(f"{{{self.namespaces['w']}}}proofState")
        spelling = grammar = "None"
        if proof_state is not None:
            spelling = proof_state.get(f"{{{self.namespaces['w']}}}spelling", "None")
            grammar = proof_state.get(f"{{{self.namespaces['w']}}}grammar", "None")

        return [spelling, grammar]

    def get_custom_xml(self):
        if self.custom_xml_content:
            props = {}
            xml = ET.fromstring(self.custom_xml_content)
            for cprop in xml.findall(".//cprop:property", self.namespaces):
                attribs = cprop.attrib
                for attr_name, attr_val in attribs.items():
                    props[attr_name] = attr_val
                for sub_prop in cprop:
                    tag = (
                        sub_prop.tag.split("}", 1)[1]
                        if "}" in sub_prop.tag
                        else sub_prop.tag
                    )
                    value = sub_prop.text
                    props[tag] = value
            return props
        return None

    def get_all_content(self, files):
        if files:
            content = {self.msword_file: {}}
            for file in files:
                content[self.msword_file][file] = {}
                xml_content = self.__load_xml(file)
                if xml_content == "":
                    continue
                if b"<?mso-contentType?>" in xml_content:
                    xml_content = (
                        xml_content.replace(b"<?mso-contentType?>", b"")
                    ).decode("utf-8")
                xml = ET.fromstring(xml_content)
                for element in xml.iter():
                    tag = (
                        element.tag.split("}")[-1]
                        if "}" in element.tag
                        else element.tag
                    )
                    if tag not in content[self.msword_file][file]:
                        content[self.msword_file][file][tag] = []
                    attribs = {}
                    for name, value in element.attrib.items():
                        name = name.split("}", 1)[-1] if "}" in name else name
                        attribs[name] = value
                    text = (element.text or "").strip()
                    if text:
                        attribs["_text"] = text
                    tail = (element.tail or "").strip()
                    if tail:
                        attribs["_tail"] = tail
                    child_tags = list(element)
                    if child_tags:
                        attribs["_children"] = []
                        for child in child_tags:
                            attribs["_children"].append(
                                child.tag.split("}")[-1]
                                if "}" in child.tag
                                else child.tag
                            )
                    content[self.msword_file][file][tag].append(attribs)
            return content
        return None

    def get_ink(self):
        ts_data = []
        for ink_file in self.ink_files:
            load_ink = self.__load_xml(ink_file)
            xml = ET.fromstring(load_ink)
            for element in xml.iter():
                tag = element.tag.split("}")[-1] if "}" in element.tag else element.tag
                if tag == "timestamp":
                    (ts_ns, ts_id), (timestring, ts) = element.attrib.items()
            ts_data.append([ink_file, ts])
        return ts_data

    def adjust_timestamp(self, ts):
        if ts:
            adjusted_timestamp = ts.replace("T", " ").replace("Z", "")
            return adjusted_timestamp.split(".")[0]
        return ""


def process_docx(filename, triage, hashing, store: DataStore):
    """
    This function accepts a filename of type Docx and processes it.
    By placing this in a function, it allows the main part of the script to accept multiple file names and
    then loop through them, calling this function for each DOCx file.
    """
    if store.ms_word_gui:
        update_status = store.ms_word_gui.update_status
    else:
        update_status = lambda msg, **kwargs: update_cli(msg, store=store, **kwargs)
    this_file = filename.msword_file
    this_rsid_root = filename.rsid_root()
    xml_files = filename.xml_files()
    update_status(f"Processing {this_file}")
    file_details = filename.details()
    third_party_paths = [
        "word\\settings.xml",
        "docProps\\core.xml",
        "docProps\\app.xml",
    ]
    third_party = False
    for line in file_details.split("\n"):
        update_status(f"    {line.rstrip()}")
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
        update_status(f"    {checkFile} exists: {xml_exists}")
        if third_party:
            update_status(
                f"    {this_file} may have been created using something other than MS Word"
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
    update_status("    Extracted Document Summary artifacts")

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
    update_status("    Extracted metadata artifacts")

    if filename.any_comments():  # checks if there are comments
        headers = [
            "File Name",
            "Author",
            "Initials",
            "Timestamp (UTC)",
            "Comment ID #",
            "Comment paraId",
            "paraId Text",
        ]
        store.comments_worksheet = (
            {k: [] for k in headers}
            if not store.comments_worksheet
            else store.comments_worksheet
        )
        for comment in filename.get_comments():
            update_status(f"    Processing comment: {comment}", level="debug")
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
        update_status("    Extracted comments artifacts")

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
            "ZIP Extra Characters (truncated)",
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
                xml_info[0],
                filename.adjust_timestamp(xml_info[1]),
                xml_info[2],
                xml_info[3],
                xml_info[4],
                xml_info[5],
                xml_info[6],
                xml_info[7],
                xml_info[8],
                xml_info[9],
            ]
            if not hashing:
                values.pop(2)
            for k, v in zip(headers, values):
                store.archive_files_worksheet[k].append(v)

        update_status("    Extracted archive files artifacts")

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
            update_status(f"    Calculating {label} count")
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
            update_status("    Processing people information from document")
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
            update_status("    Processing extensible comments data")
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
            update_status("    Processing extensible comments data")
            headers = ["File Name", "paraId", "paraIdParent", "Done?"]
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
            update_status("    Processing comments ids")
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
            update_status("    Processing custom properties")
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
                update_status("    Processing item.xml files")
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
                update_status("    Processing ink.xml files")
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

    update_status(f"Finished processing {this_file}")
    update_status(f'{"-"*36}')


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
    type_map = {
        "<w:p> tags": "Int32",
        "<w:r> tags": "Int32",
        "<w:t> tags": "Int32",
        "<w:tr> tags": "Int32",
        "<w14:docId>": "string",
        "<w15:docId>": "string",
        "<w16:docId>": "string",
        "App Version": "string",
        "Application": "string",
        "Archive File": "string",
        "Author": "string",
        "Category": "string",
        "Characters With Spaces": "Int32",
        "Characters": "Int32",
        "Comment ID #": "Int32",
        "Comment paraId": "string",
        "Company": "string",
        "Content Status": "string",
        "Content": "string",
        "Count in document.xml": "Int32",
        "Created Date": "datetime64[ns]",
        "dateUtc": "datetime64[ns]",
        "Description": "string",
        "Doc Security": "Int32",
        "Done?": "boolean",
        "durableId": "string",
        "File Created Date": "datetime64[ns]",
        "File Modified Date": "datetime64[ns]",
        "File Name": "string",
        "Grammar Check": "string",
        "Has Comments": "boolean",
        "Has Ink": "boolean",
        "Hyperlinks Changed": "string",
        "Hyperlinks": "string",
        "Initials": "string",
        "Ink XML File": "string",
        "Item XML File": "string",
        "Keywords": "string",
        "Language": "string",
        "Last Modified By": "string",
        "Last Printed Date": "datetime64[ns]",
        "Lines": "Int32",
        "Manager": "string",
        "MD5 Hash": "string",
        "Modified Date": "datetime64[ns]",
        "Modified Time (local/UTC/Redmond, Washington)": "datetime64[ns]",
        "Pages": "Int32",
        "Paragraphs": "Int32",
        "paraId Text": "string",
        "paraId": "string",
        "paraIdParent": "string",
        "providerId": "string",
        "reactionDateUtc": "datetime64[ns]",
        "reactionType": "string",
        "Revision": "Int32",
        "RSID Root": "string",
        "RSID Type": "string",
        "RSID Value": "string",
        "Shared Doc": "string",
        "Source": "string",
        "Spell Check": "string",
        "Subject": "string",
        "Template": "string",
        "Timestamp (UTC)": "datetime64[ns]",
        "Timestamp": "datetime64[ns]",
        "Title": "string",
        "Total Editing Time": "string",
        "Type": "string",
        "Uncompressed Size (bytes)": "Int32",
        "Unique rsidR": "Int32",
        "uri": "string",
        "userId": "string",
        "userName": "string",
        "userProvider": "string",
        "Value": "string",
        "Version": "string",
        "Words": "Int32",
        "ZIP Compression Type": "string",
        "ZIP Create System": "Int32",
        "ZIP Created Version": "Int32",
        "ZIP Extra Characters (truncated)": "string",
        "ZIP Extra Flag (len)": "Int32",
        "ZIP Extract Version": "Int32",
        "ZIP Flag Bits (hex)": "string",
    }

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
                update_status(f'"{actual_name}" written.')
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
        update_status(
            "Generating Timeline worksheet - this may take some time depending on the number of documents being parsed ..."
        )
        store.timeline_worksheet = generate_timeline(store)
        process_and_write(store.timeline_worksheet, "Timeline", "timeline")
        generate_visual_timeline(writer, store.timeline_worksheet)
        update_status('"Visual Timeline" written.')
        write_tips(writer)
        update_status('"Tips" worksheet written.')
        if store.sqlite:
            excel_parent = os.path.dirname(excel_file)
            excel_name = os.path.splitext(os.path.basename(excel_file))[0]
            db_file = os.path.normpath(f"{excel_parent}{os.sep}{excel_name}.db")
            if os.path.exists(db_file):
                try:
                    os.remove(db_file)
                except:
                    update_status(f'Unable to remove "{db_file}".')
                    db_file = os.path.normpath(f"{excel_parent}{os.sep}{excel_name}_{store.timestamp}.db")
            update_status(f'Writing results to "{db_file}".')
            conn = sqlite3.connect(db_file)
            for sheet, sheet_name, layout in triage_sheets:
                if sheet:
                    pd.DataFrame(sheet).to_sql(sheet_name, conn, index=False)
            if not triage_files:
                for sheet, sheet_name, layout in full_sheets:
                    if sheet:
                        pd.DataFrame(sheet).to_sql(sheet_name, conn, index=False)
            if aggregated:
                pd.DataFrame(store.aggregated_worksheet).to_sql(
                    "Aggregated Comment Data", conn, index=False
                )
            if not store.timeline_worksheet.empty:
                pd.DataFrame(store.timeline_worksheet).to_sql("Timeline", conn, index=False)
            conn.close()
            update_status(f'All data written to "{db_file}".')


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
            #"major_unit": 120,
            "major_unit": major_unit_days,
            "major_unit_type": "days",
            #"minor_unit": 30,
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
        return
    store.logger.log(log_level, msg)


def process_cli(files, triage_files, hash_files, excel_file, store: DataStore):
    docxErrorCount = 0
    store.start_time = dt.now().strftime(__dtfmt__)
    update_cli(f"{__appname__}", store=store)
    update_cli(f"Command line: {' '.join(sys.argv)}", store=store)
    update_cli(f"Output File Path: {os.path.dirname(excel_file)}", store=store)
    update_cli(f"Excel output file: {os.path.basename(excel_file)}", store=store)
    update_cli(f"Log file: {os.path.abspath(store.log_file)}", store=store)
    update_cli(f"The following {len(files)} files are being processed:", store=store)
    joiner = f"\n{dt.now().strftime(__dtfmt__)} -     "
    update_cli(
        "    " + joiner.join(os.path.abspath(str(f)) for f in files), store=store
    )
    update_cli(f"Script executed: {store.start_time}", store=store)
    update_cli("Summary of files parsed:", store=store)
    update_cli(f'{"="*36}', store=store)

    remaining = len(files)
    for f in files:
        try:
            f = os.path.abspath(str(f))
            with Docx(f, triage_files, hash_files, store=store) as doc:
                process_docx(doc, triage_files, hash_files, store)
        except Exception as docxError:
            # If processing a DOCx file raises an error, let the user know, and write it
            # to the error log.
            docxErrorCount += 1  # increment error count by 1.
            update_cli(
                f"Error trying to process {f}. Skipping. Error: {docxError}",
                level="error",
                color=__red__,
                store=store,
            )
            store.errors_worksheet["File Name"].append(f)
            store.errors_worksheet["Error"].append(docxError)
        if remaining != 0:
            remaining -= 1
    write_to_excel(excel_file, triage_files, store)
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


def cli_log(excel_path, verbose=False, store: DataStore = None):
    log = logging.getLogger("ms-word-parser")
    log.setLevel(logging.INFO)
    log_fmt = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt=__dtfmt__,
    )
    log_path = os.path.normpath(f"{excel_path}{os.sep}{store.log_file}")
    file_handler = logging.FileHandler(log_path, "w", "utf-8")
    file_handler.setFormatter(log_fmt)
    log.addHandler(file_handler)
    if verbose:
        store.color_fmt = ColorFormatter()
        stream_handler = logging.StreamHandler(stream=sys.stdout)
        stream_handler.setLevel(logging.DEBUG)
        stream_handler.setFormatter(store.color_fmt)
        log.addHandler(stream_handler)
    return log


def stop_cli(triage_files, excel_file, store: DataStore = None):
    update_cli("Processing stopped", store=store)
    update_cli("Attempting to write current results to Excel", store=store)
    docxErrorCount = len(store.errors_worksheet["Error"])
    try:
        write_to_excel(excel_file, triage_files, store)
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
        "-e", "--excel", help="output path and filename for the Excel output"
    )
    arg_parse.add_argument("-g", "--gui", action="store_true", help="launch the gui")
    arg_parse.add_argument(
        "--hash", action="store_true", help="hash the doc zip contents"
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
        help="save data to an sqlite database - requires -e/--excel",
    )
    arg_parse.add_argument(
        "-V",
        "--verbose",
        action="store_true",
        help="Output to STDOUT as well as log",
        default=False,
    )
    file_source = arg_parse.add_mutually_exclusive_group(required=False)
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
        if not (args.dir or args.files):
            arg_parse.error(
                "One of --files or --dir is required, unless running in GUI mode"
            )
        if not (args.triage or args.full):
            arg_parse.error(
                "One of --triage or --full is required, unless running in GUI mode"
            )
        if not args.excel:
            arg_parse.error(
                "You must supply -e / --excel as a path and file name for the output Excel content"
            )
        if args.excel:
            if not os.path.exists(os.path.abspath(os.path.dirname(args.excel))):
                arg_parse.error(
                    f"The path {os.path.abspath(os.path.dirname(args.excel))} does not exist. Please check your path and try again."
                )
            store.logger = cli_log(
                os.path.abspath(os.path.dirname(args.excel)), args.verbose, store=store
            )
        if args.sqlite:
            store.sqlite = True
        if args.files:
            file_list = args.files
            try:
                process_cli(
                    file_list,
                    args.triage,
                    args.hash,
                    os.path.abspath(args.excel),
                    store,
                )
            except KeyboardInterrupt:
                stop_cli(args.triage, os.path.abspath(args.excel), store)
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
            try:
                process_cli(
                    file_list,
                    args.triage,
                    args.hash,
                    os.path.abspath(args.excel),
                    store,
                )
            except KeyboardInterrupt:
                stop_cli(args.triage, os.path.abspath(args.excel), store)
            except Exception as e:
                update_cli(
                    f"Error trying to process directory - {e}",
                    level="error",
                    color=__red__,
                    store=store,
                )


if __name__ == "__main__":
    main()
