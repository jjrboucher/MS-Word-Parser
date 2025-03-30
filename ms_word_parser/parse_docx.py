#!/usr/bin/env python3

# Written by Jacques Boucher
# jjrboucher@gmail.com
#
# ********** Description **********
"""The script does not attempt to validate the loaded file.
A docx file is nothing more than a ZIP file, hence why this script uses the zipfile library.

It will extract the results to an Excel file in a location, and with a name, both defined by you.
If the file does not exist, it creates it. If the file does exist, it will overwrite it.
You will have the option to load a single file, or load a directory. If you select directory,
you will be prompted to decide if you'd like the script to recursively load all files from that path.
This allows you to run this against many DOCx files at once for an investigation and compare results.

Usage:

- First, click File - Select Excel File. This will be the file containing the output. If it exists,
  it will be overwritten. If not, it will be created.
  
- Second, click File - Open Files ... or Open Directory ..., depending on what you'd like to do.
  Again, if you select a directory, you will be asked if you'd like to recursively load all files.
  
- Third, choose your processing options: Triage or Full.
  Triage will give cursory information about the document and the metadata contained within, as well as any
  comments.
  Full will do a full analysis of the document, including looking at RSID's and determining uniqueness,
  and examining w:p, w:r, and w:t tags.
  
  You also have the option to Hash the file and the contents of the file (ie: the files in the zip). 
  This will generate an MD5 Hash for each value.
  
- Fourth, click Process at the bottom left of the Window. The output will be placed both in the Processing
  Status window on the right, and in a log file named DOCx_Parser_Log_<date_time>.log in the same path
  as the Excel document. The date_time value, and subsequently the log name, are determined at launch,
  but the log file will only be created once an Excel document is chosen and files are selected, even
  if the Process button is not selected.

Processes that this script will do:

1 - It will extract a list of all the files within the zip file and save it to a worksheet called XML_files.
    In this worksheet, it will save the following information to a row:
    "File Name", "XML", "Size (bytes)", "MD5Hash"

2 - It will extract all the unique RSIDs from the file word/settings.xml and write it to a worksheet
    called doc_summary.
    In this worksheet, it will save the following information to a row:
    "File Name", "Unique rsidR", "RSID Root", "<w:p> tags", "<w:r> tags", "<w:t> tags"
    Where "Unique RSID" is a numerical count of the of RSIDs in the file settings.xml.

    What is an RSID (Revision Save ID)?
    See https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.rsid?view=openxml-2.8.1

3 - It will extract all the unique RSIDs from the file word/settings.xml and write it to a worksheet
    called RSIDs, along with a count of how many times that RSID is in document.xml
    It will also search document.xml for all unique rsidRPr, rsidP, and rsidRDefault values and count 
    of how many are in document.xml.
    It also extracts the unique paraId and textId tags from the <w:p> tag and saves the values and count
    of how many are in document.xml.
    In this worksheet, it will save the following information to rows (one for each unique RSID):
    "File Name", "RSID Type", "RSID Value", "Count in document.xml"

4 - It will extract all known relevant metadata from the files docProps/app.xml and docProps/core.xml
    and write it to a worksheet called metadata.
    In this worksheet, it will save the following information to a row:
    "File Name", "Author", "Created Date", "Last Modified By", "Modified Date", "Last Printed Date",
    "Manager", "Company", "Revision", "Total Editing Time", "Pages", "Paragraphs", "Lines", "Words",
    "Characters", "Characters With Spaces", "Title", "Subject", "Keywords", "Description", "Application",
    "App Version", "Template", "Doc Security", "Category", "contentStatus" """
# ********** Possible future enhancements **********


import hashlib
import os
import zipfile
import logging
import subprocess
from datetime import datetime as dt, timedelta
from pathlib import Path
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
    QFrame,
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

warnings.filterwarnings("ignore", category=DeprecationWarning)
filesUnableToProcess = []  # list of files that produced an error
doc_summary_worksheet = {}  # contains summary data parsed from each file processed
metadata_worksheet = {}  # contains the metadata parsed from each file processed
archive_files_worksheet = {}  # contains the archive files data from each file processed
rsids_worksheet = {}  # contains the RSID artifacts extracted from each file processed
comments_worksheet = {}  # contains the comments within each file processed
timestamp = dt.now().strftime("%Y%m%d_%H%M%S")
log_file = f"DOCx_Parser_Log_{timestamp}.log"
ms_word_form = None
green = QColor(86, 208, 50)
red = QColor(204, 0, 0)
black = QColor(0, 0, 0)
__version__ = "2.0.0"
__appname__ = f"MS Word Parser v{__version__}"
__source__ = "https://github.com/jjrboucher/MS-Word-Parser"
__date__ = "28 March 2025"
__author__ = (
    "Jacques Boucher - jjrboucher@gmail.com\nCorey Forman - corey@digitalsleuth.ca"
)
__dtfmt__ = "%Y-%m-%d %H:%M:%S"


class AboutWindow(QWidget):
    """Sets the structure for the About window"""

    def __init__(self):
        super().__init__()
        layout = QGridLayout()
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

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Contents")
        self.setFixedSize(600, 800)
        self.text_edit = QPlainTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setPlainText(__doc__)
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


class UiDialog:

    def __init__(self):
        super().__init__()
        self.d_width = 1152
        self.d_height = 330
        self.files = []
        self.excel_path = ""
        self.excel_full_path = ""
        self.log_path = ""
        self.log_handler = None
        self.logger = logging.getLogger("ms-word-parser")
        self.logger.setLevel(logging.DEBUG)
        self.log_fmt = logging.Formatter(
            "%(asctime)s | %(levelname)-8s | %(message)s",
            datefmt=__dtfmt__,
        )
        self.text_font = QFont()
        self.text_font.setPointSize(9)
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
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.parsingOptions = QGroupBox(self.centralwidget)
        self.parsingOptions.setObjectName("parsingOptions")
        self.parsingOptions.setGeometry(QRect(10, 10, 350, 60))
        self.parsingOptions.setStyleSheet("background: #ffffff; color: black;")
        self.parsingOptions.setFont(self.text_font)
        self.triageButton = QRadioButton(self.parsingOptions)
        self.triageButton.setObjectName("triageButton")
        self.triageButton.setGeometry(QRect(10, 30, 89, 20))
        self.triageButton.setStyleSheet(self.stylesheet)
        self.triageButton.setChecked(True)
        self.triageButton.setFont(self.text_font)
        self.fullButton = QRadioButton(self.parsingOptions)
        self.fullButton.setObjectName("fullButton")
        self.fullButton.setGeometry(QRect(90, 30, 89, 20))
        self.fullButton.setStyleSheet(self.stylesheet)
        self.fullButton.setFont(self.text_font)
        self.separator = QFrame(self.parsingOptions)
        self.separator.setFrameShape(QFrame.Shape.Box)
        self.separator.setFrameShadow(QFrame.Shadow.Plain)
        self.separator.setGeometry(QRect(220, 20, 6, 60))
        self.separator.setStyleSheet(self.separator_sheet)
        self.hashFiles = QCheckBox(self.parsingOptions)
        self.hashFiles.setObjectName("hashFiles")
        self.hashFiles.setGeometry(QRect(250, 30, 75, 20))
        self.hashFiles.setStyleSheet(self.stylesheet)
        self.hashFiles.setFont(self.text_font)
        self.outputFiles = QGroupBox(self.centralwidget)
        self.outputFiles.setObjectName("outputFiles")
        self.outputFiles.setGeometry(QRect(10, 76, 350, 120))
        self.outputFiles.setStyleSheet("background-color: #ffffff; color: black;")
        self.outputFiles.setFont(self.text_font)
        self.excelFileLabel = QLabel(self.outputFiles)
        self.excelFileLabel.setObjectName("excelFileLabel")
        self.excelFileLabel.setGeometry(QRect(10, 30, 80, 16))
        self.excelFileLabel.setStyleSheet("background: #fcfcfc; color: black;")
        self.excelFileLabel.setFont(self.text_font)
        self.excelFileText = "File -> Select Excel File or click 'Select Excel'"
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
        self.generalLog.setStyleSheet("background: #fcfcfc; color: black;")
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
        self.outputPathLabel.setStyleSheet("background: #fcfcfc; color: black;")
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

        self.operationOptions = QGroupBox(self.centralwidget)
        self.operationOptions.setObjectName("operationOptions")
        self.operationOptions.setGeometry(QRect(10, 200, 350, 90))
        self.operationOptions.setStyleSheet("background-color: #ffffff; color:black;")
        self.operationOptions.setFont(self.text_font)
        self.excelButton = QPushButton(self.operationOptions)
        self.excelButton.setObjectName("excelButton")
        self.excelButton.setGeometry(QRect(10, 28, 86, 24))
        self.excelButton.setStyleSheet(self.stylesheet)
        self.excelButton.clicked.connect(self.open_excel)
        self.addFilesButton = QPushButton(self.operationOptions)
        self.addFilesButton.setObjectName("addFilesButton")
        self.addFilesButton.setGeometry(QRect(112, 28, 86, 24))
        self.addFilesButton.setEnabled(False)
        self.addFilesButton.setStyleSheet(self.disabled)
        self.addFilesButton.clicked.connect(self.add_files)
        self.addDirectoryButton = QPushButton(self.operationOptions)
        self.addDirectoryButton.setObjectName("addDirectoryButton")
        self.addDirectoryButton.setGeometry(QRect(214, 28, 86, 24))
        self.addDirectoryButton.setEnabled(False)
        self.addDirectoryButton.setStyleSheet(self.disabled)
        self.addDirectoryButton.clicked.connect(self.add_directory)
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
        self.processStatus = QGroupBox(self.centralwidget)
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
        self.numOfFilesLabel.setStyleSheet("background: #fcfcfc; color: black;")
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
        self.numOfErrorsLabel.setStyleSheet("background: #fcfcfc; color: black;")
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
        self.numRemainingLabel.setStyleSheet("background: #fcfcfc; color: black;")
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
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setObjectName("menubar")
        self.menubar.setGeometry(QRect(0, 0, 1192, 22))
        self.menuFile = QMenu(self.menubar)
        self.menuFile.setObjectName("menuFile")
        self.menuHelp = QMenu(self.menubar)
        self.menuHelp.setObjectName("menuHelp")
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

    # setupUi

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(
            QCoreApplication.translate("MainWindow", __appname__, None)
        )
        self.actionSelect_Excel.setText(
            QCoreApplication.translate("MainWindow", "Select &Excel File ...", None)
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
        self.triageButton.setText(
            QCoreApplication.translate("MainWindow", "Triage", None)
        )
        self.fullButton.setText(QCoreApplication.translate("MainWindow", "Full", None))
        self.hashFiles.setText(
            QCoreApplication.translate("MainWindow", "Hash Files", None)
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
            QCoreApplication.translate("MainWindow", log_file, None)
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
            response = QMessageBox.question(
                None,
                "Load recursively",
                "Do you want to recursively load all files in this directory?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if response == QMessageBox.StandardButton.Yes:
                recursive_list = (
                    list(folder_path.rglob("*.docx"))
                    + list(folder_path.rglob("*.dotx"))
                    + list(folder_path.rglob("*.dotm"))
                )
                files = [str(file) for file in recursive_list]
            else:
                non_recursive_list = (
                    list(folder_path.glob("*.docx"))
                    + list(folder_path.glob("*.dotx"))
                    + list(folder_path.glob("*.dotm"))
                )
                files = [str(file) for file in non_recursive_list]
            self.numOfFiles.setText(str(len(files)))
            self.numRemaining.setText(str(len(files)))
            if files:
                update_status(f"The following {len(files)} files have been loaded:")
                joiner = f"\n{dt.now().strftime(__dtfmt__)} -     "
                logging_patch = []
                for file in files:
                    file = file.encode("utf-8", "ignore").decode()
                    logging_patch.append(file)
                update_status("    " + joiner.join(logging_patch))
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
            "docx, dotx, dotm Files (*.docx *.dotx *.dotm)",
        )
        if files:
            for file in files:
                all_files.append(os.path.normpath(file))
            self.numOfFiles.setText(str(len(all_files)))
            self.numRemaining.setText(str(len(all_files)))
            update_status(f"The following {len(all_files)} files have been loaded:")
            joiner = f"\n{dt.now().strftime(__dtfmt__)} -     "
            logging_patch = []
            for file in all_files:
                file = file.encode("utf-8", "ignore").decode()
                logging_patch.append(file)
            update_status("    " + joiner.join(logging_patch))
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
            self.log_path = os.path.normpath(f"{self.excel_path}{os.sep}{log_file}")
            self.log_handler = logging.FileHandler(self.log_path)
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
            self.generalLogFile.setText(log_file)
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
        this_os = os.sys.platform
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
                subprocess.Popen([launch, file])
        except Exception as e:
            self.update_status(f"Unable to open {file}: {e}", level="error")

    def _reset(self):
        global timestamp, log_file
        timestamp = dt.now().strftime("%Y%m%d_%H%M%S")
        log_file = f"DOCx_Parser_Log_{timestamp}.log"
        self.excelFile.setText(self.excelFileText)
        self.generalLogFile.setText(log_file)
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
        self.hashFiles.setChecked(False)
        self.stopButton.setEnabled(False)
        self.stopButton.setStyleSheet(self.disabled)

    def _stop(self):
        self.running = False
        self.stopButton.setStyleSheet(self.disabled)
        self.stopButton.setEnabled(False)

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
        self.aboutWindow.show()

    def _contents(self):
        self.contentsWindow = ContentsWindow()
        self.contentsWindow.setWindowFlags(
            self.contentsWindow.windowFlags() & ~Qt.WindowType.WindowMinMaxButtonsHint
        )
        self.contentsWindow.show()

    def update_status(self, msg, level="info", color=black):
        if level == "info":
            self.docxOutput.setTextColor(color)
            self.docxOutput.append(f"{dt.now().strftime(__dtfmt__)} - {msg}")
            self.docxOutput.setTextColor(black)
            if not msg.isascii():
                msg = msg.encode("latin-1", "ignore").decode()
            self.logger.info(msg)
        elif level == "error":
            self.docxOutput.setTextColor(color)
            self.docxOutput.append(f"{dt.now().strftime(__dtfmt__)} - {msg}")
            self.docxOutput.setTextColor(black)
            if not msg.isascii():
                msg = msg.encode("latin-1", "ignore").decode()
            self.logger.error(msg)
        elif level == "debug":
            if not msg.isascii():
                msg = msg.encode("latin-1", "ignore").decode()
            self.logger.debug(msg)
        QApplication.processEvents()

    def analyze_docs(self, files, triage_files, hash_files):
        if not self.running:
            self.running = True
        start_time = dt.now().strftime(__dtfmt__)
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
        errors_worksheet = {"File Name": [], "Error": []}
        for f in files:  # loop over the files selected, processing each.
            if not self.running:
                update_status("Processing stopped")
                self.stopButton.setEnabled(False)
                self.resetButton.setEnabled(True)
                self.resetButton.setStyleSheet(self.stylesheet)
                return
            try:
                process_docx(Docx(f, triage_files, hash_files))
            except Exception as docxError:
                # If processing a DOCx file raises an error, let the user know, and write it
                # to the error log.
                docxErrorCount += 1  # increment error count by 1.
                self.numOfErrors.setText(str(docxErrorCount))
                filesUnableToProcess.append(f)
                update_status(
                    f"Error trying to process {f}. Skipping. Error: {docxError}",
                    level="error",
                    color=red,
                )
                errors_worksheet["File Name"].append(
                    f
                )  ## DEBUG - Prepare for Errors Worksheet
                errors_worksheet["Error"].append(docxError)
            if remaining != 0:
                remaining -= 1
            self.numRemaining.setText(str(remaining))
        with pd.ExcelWriter(
            path=self.excel_full_path, engine="xlsxwriter", mode="w"
        ) as writer:
            df_summary = chunk_list(doc_summary_worksheet, "Doc_Summary")
            for chunk_dict, sheet_name in df_summary:
                df_summary_chunk = pd.DataFrame(data=chunk_dict)
                if not df_summary_chunk.empty:
                    df_summary_chunk.to_excel(
                        excel_writer=writer, sheet_name=sheet_name, index=False
                    )
                    worksheet = writer.sheets[sheet_name]
                    (max_row, max_col) = df_summary_chunk.shape
                    worksheet.set_column(0, 1, 34)
                    worksheet.set_column(2, max_col - 4, 16)
                    worksheet.set_column(max_col - 3, max_col - 1, 40)
                    worksheet.autofilter(0, 0, max_row, max_col - 1)
                    update_status(f'"{sheet_name}" worksheet written to Excel.')
            df_metadata = chunk_list(metadata_worksheet, "Metadata")
            for chunk_dict, sheet_name in df_metadata:
                df_metadata_chunk = pd.DataFrame(data=chunk_dict)
                if not df_metadata_chunk.empty:
                    df_metadata_chunk.to_excel(
                        excel_writer=writer, sheet_name=sheet_name, index=False
                    )
                    worksheet = writer.sheets[sheet_name]
                    (max_row, max_col) = df_metadata_chunk.shape
                    worksheet.set_column(0, max_col - 1, 20)
                    worksheet.autofilter(0, 0, max_row, max_col - 1)
                    update_status(f'"{sheet_name}" worksheet written to Excel.')
            df_comments = chunk_list(comments_worksheet, "Comments")
            for chunk_dict, sheet_name in df_comments:
                df_comments_chunk = pd.DataFrame(data=chunk_dict)
                if not df_comments_chunk.empty:
                    df_comments_chunk.to_excel(
                        excel_writer=writer, sheet_name=sheet_name, index=False
                    )
                    worksheet = writer.sheets[sheet_name]
                    (max_row, max_col) = df_comments_chunk.shape
                    worksheet.set_column(0, max_col - 2, 20)
                    worksheet.set_column(max_col - 1, max_col - 1, 140)
                    worksheet.autofilter(0, 0, max_row, max_col - 1)
                    update_status(f'"{sheet_name}" worksheet written to Excel.')
            if not triage_files:
                df_rsids = chunk_list(rsids_worksheet, "RSIDs")
                for chunk_dict, sheet_name in df_rsids:
                    df_rsids_chunk = pd.DataFrame(data=chunk_dict)
                    if not df_rsids_chunk.empty:
                        df_rsids_chunk.to_excel(
                            excel_writer=writer, sheet_name=sheet_name, index=False
                        )
                        worksheet = writer.sheets[sheet_name]
                        (max_row, max_col) = df_rsids_chunk.shape
                        worksheet.set_column(0, max_col - 1, 20)
                        worksheet.autofilter(0, 0, max_row, max_col - 1)
                        update_status(f'"{sheet_name}" worksheet written to Excel.')
                df_archive = chunk_list(archive_files_worksheet, "Archive Files")
                for chunk_dict, sheet_name in df_archive:
                    df_archive_chunk = pd.DataFrame(data=chunk_dict)
                    if not df_archive_chunk.empty:
                        df_archive_chunk.to_excel(
                            excel_writer=writer, sheet_name=sheet_name, index=False
                        )
                        worksheet = writer.sheets[sheet_name]
                        (max_row, max_col) = df_archive_chunk.shape
                        worksheet.set_column(0, max_col - 1, 35)
                        worksheet.autofilter(0, 0, max_row, max_col - 1)
                        update_status(f'"{sheet_name}" worksheet written to Excel.')
            df_errors = chunk_list(errors_worksheet, "Errors")
            for chunk_dict, sheet_name in df_errors:
                df_errors_chunk = pd.DataFrame(data=chunk_dict)
                if not df_errors_chunk.empty:
                    df_errors_chunk.to_excel(
                        excel_writer=writer, sheet_name=sheet_name, index=False
                    )
                    worksheet = writer.sheets[sheet_name]
                    (max_row, max_col) = df_errors_chunk.shape
                    worksheet.set_column(0, max_col - 1, 34)
                    update_status(f'"{sheet_name}" worksheet written to Excel.')
        script_end = dt.now().strftime(__dtfmt__)
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
            for each_file in filesUnableToProcess:
                update_status(f"  {each_file}", "error", color=clr)
        update_status(f"Script finished execution: {script_end}", color=green)
        end_time = dt.now().strftime(__dtfmt__)
        run_time = str(
            timedelta(
                seconds=(
                    dt.strptime(end_time, __dtfmt__)
                    - dt.strptime(start_time, __dtfmt__)
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
        self.openExcelButton.setStyleSheet(self.stylesheet)
        self.openExcelButton.setEnabled(True)


def chunk_list(sheet_dict, sheet_name):
    chunks = []
    if "File Name" in sheet_dict and len(sheet_dict["File Name"]) > 1000000:
        file_names = sheet_dict["File Name"]
        list_len = len(file_names)
        chunk_size = 1000000

        for start in range(0, list_len, chunk_size):
            end = min(start + chunk_size, list_len)
            chunk_dict = {
                key: value[start:end] if isinstance(value, list) else value
                for key, value in sheet_dict.items()
            }
            chunks.append((chunk_dict, f"{sheet_name}_{len(chunks) + 1}"))
    else:
        chunks.append((sheet_dict, sheet_name))
    return chunks


class MsWordGui(QMainWindow, UiDialog):
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
            background: #fcfcfc; color:black;
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
            background-color: #ffffff; border: 1px solid black; color: black;
        }
        QPushButton:hover {
            background-color: #d9ebfb; border: 1px solid black;
        }
        QRadioButton {
            background: #fcfcfc; color:black;
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

    def __init__(self):
        """Call and setup the UI"""
        super().__init__()
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

    def __init__(self, msword_file, triage=False, hashing=True):
        """
        .docx file to pass to the class
        Triage value can be True or False. If True, will parse less info to execute faster.
        When set to False, it does not try to parse RSID values from document.xml.
        If triage value not passed, it defaults to False and does full parsing.
        The script using this class still ultimately decides what methods it wants to use.
        But if in triage mode, some of the variables will not get assigned any value, thus
        will affect any methods that rely on those variables having a value assigned to them.
        """
        self.namespaces = {
            "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            "dc": "http://purl.org/dc/elements/1.1/",
            "dcterms": "http://purl.org/dc/terms/",
            "dcmitype": "http://purl.org/dc/dcmitype/",
            "default": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
            "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
            "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
            "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "xsi": "http://www.w3.org/2001/XMLSchema-instance",
        }
        self.msword_file = msword_file
        self.hashing = hashing
        self.header_offsets, self.binary_content = self.__find_binary_string()
        self.extra_fields = self.__xml_extra_bytes()
        self.core_xml_file = "docProps/core.xml"
        self.core_xml_content = self.__load_xml(self.core_xml_file)
        if self.core_xml_content == "":
            self.core_xml_file = "docProps\\core.xml"
            self.core_xml_content = self.__load_xml(self.core_xml_file)
        self.app_xml_file = "docProps/app.xml"
        self.app_xml_content = self.__load_xml(self.app_xml_file)
        if self.app_xml_content == "":
            self.app_xml_file = "docProps\\app.xml"
            self.app_xml_content = self.__load_xml(self.app_xml_file)
        self.document_xml_file = "word/document.xml"
        self.document_xml_content = self.__load_xml(self.document_xml_file)
        if self.document_xml_content == "":
            self.document_xml_file = "word\\document.xml"
            self.document_xml_content = self.__load_xml(self.document_xml_file)
        self.has_comments = ""
        self.comments_file = "word/comments.xml"
        self.comments_xml_content = self.__load_xml(self.comments_file)
        if self.comments_xml_content == "":
            self.comments_file = "word\\comments.xml"
            self.comments_xml_content = self.__load_xml(self.comments_file)
        self.settings_xml_file = "word/settings.xml"
        self.settings_xml_content = self.__load_xml(self.settings_xml_file)
        if self.settings_xml_content == "":
            self.settings_xml_file = "word\\settings.xml"
            self.settings_xml_content = self.__load_xml(self.settings_xml_file)
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

        if not triage:  # if not run in triage mode, do full parsing

            self.rsidR_in_document_xml = self.__rsids_in_document_xml("rsidR")
            self.rsidRPr = self.__rsids_in_document_xml("rsidRPr")
            self.rsidP = self.__rsids_in_document_xml("rsidP")
            self.rsidRDefault = self.__rsids_in_document_xml("rsidRDefault")
            self.para_id = self.__rsids_in_document_xml("paraId")
            self.text_id = self.__rsids_in_document_xml("textId")

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
        zip_header = {
            "signature": [0, 4],  # byte 0 for 4 bytes
            "extract version": [4, 2],  # byte 4 for 2 bytes
            "bitflag": [6, 2],  # byte 6 for 2 bytes
            "compression": [8, 2],  # byte 8 for 2 bytes
            "modification time": [10, 2],  # byte 10 for 2 bytes
            "modification date": [12, 2],  # byte 12 for 2 bytes
            "CRC-32": [14, 4],  # byte 14 for 4 bytes
            "compressed size": [18, 4],  # byte 18 for 4 bytes
            "uncompressed size": [22, 4],  # byte 22 for 4 bytes
            "filename length": [26, 2],  # byte 26 for 2 bytes
            "extra field length": [28, 2],  # byte 28 for 2 bytes
        }
        # filename is at offset 30 for n where n is "filename length". Extra field is at offset 30
        # + filename length for z bytes where z is "extra field length

        extras = {}  # empty dictionary where values will be stored.

        truncate_extra_field = 20  # extra field can be several hundred bytes, mostly 0x00. Grab display first 10

        for offset in self.header_offsets:

            filename_len = int.from_bytes(
                self.binary_content[
                    zip_header["filename length"][0]
                    + offset : zip_header["filename length"][1]
                    + offset
                    + zip_header["filename length"][0]
                ],
                "little",
            )

            filename_start = offset + 30
            filename_end = offset + 30 + filename_len

            if filename_end - filename_start < 256:
                # some DOCx files somehow produce false positives of
                # excessively long filenames and results in an error. This avoids that error.
                filename = self.binary_content[filename_start:filename_end].decode(
                    "ascii"
                )
            extrafield_len = int.from_bytes(
                self.binary_content[
                    zip_header["extra field length"][0]
                    + offset : zip_header["extra field length"][1]
                    + offset
                    + zip_header["extra field length"][0]
                ],
                "little",
            )  # getting binary value, little endien

            extrafield_start = filename_end
            extrafield_end = extrafield_start + extrafield_len

            extrafield = self.binary_content[extrafield_start:extrafield_end]

            extrafield_hex_as_text = []
            # List that will contain the extra characters represented as text.

            for h in extrafield:
                extrafield_hex_as_text.append(str(hex(h)))

            if extrafield_len == 0:  # many are 0 bytes, so skipping those.
                extras[filename] = [extrafield_len, "nil"]
            elif (
                extrafield_len <= truncate_extra_field
            ):  # field size larger than truncate value
                extras[filename] = [extrafield_len, extrafield_hex_as_text]
            else:
                extras[filename] = [
                    extrafield_len,
                    extrafield_hex_as_text[0:truncate_extra_field],
                ]  # adds only
                # the select # of characters as specified in the variable truncate_extra_field. This is so that
                # we don't end up with hundreds of characters in a cell in Excel, as some extra fields can be
                # several hundred values long. But so far, most are 0x00, with only the first few being values other
                # than hex 0x00.

        return extras

    def __load_xml(self, xml_file):
        if (
            xml_file in self.xml_files()
        ):  # if the file exists, read it and return its content
            if "comments.xml" in xml_file:
                self.has_comments = True
            with zipfile.ZipFile(self.msword_file, "r") as zipref:
                with zipref.open(xml_file) as xmlFile:
                    return xmlFile.read()
        else:
            if "comments.xml" in xml_file:
                self.has_comments = False
            ms_word_form.update_status(
                f'"{xml_file}" does not exist in "{self.msword_file}". '
                f"Returning empty string.",
                level="debug",
            )
            return ""

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
            meta_content = ns_extract.text if ns_extract is not None else ""
        else:
            return ""
        return meta_content

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

        if not self.has_comments:  # There are no comments
            return ["", "", "", "", ""]
        xml = ET.fromstring(self.comments_xml_content)
        # Find all comments
        comments = xml.findall(".//w:comment", self.namespaces)
        all_comments = []  # list to contain all comments
        for comment in comments:
            author = comment.get(f"{{{self.namespaces['w']}}}author")
            date_time = comment.get(f"{{{self.namespaces['w']}}}date")
            initials = comment.get(f"{{{self.namespaces['w']}}}initials")
            comment_id = comment.get(f"{{{self.namespaces['w']}}}id")
            text = (
                "".join(
                    [
                        t.text
                        for t in comment.findall(".//w:t", self.namespaces)
                        if t.text
                    ]
                )
                .encode("latin-1", "ignore")
                .decode()
            )
            all_comments.append([comment_id, date_time, author, initials, text])
        return all_comments

    def any_comments(self):
        return self.has_comments

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
            "paraId": self.namespaces["w14"],
            "textId": self.namespaces["w14"],
        }
        for entry in [self.p_tags, self.r_tags, self.t_tags]:
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

    def hash(self):
        """
        Function that will return the hash of the file itself
        """
        if self.hashing:  # if hashing option was selected
            filehash = hashlib.md5()
            filehash.update(self.binary_content)
            return filehash.hexdigest()
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
        with zipfile.ZipFile(self.msword_file, "r") as zip_file:
            # returns XML files in the DOCx
            xml_files = {}
            for file_info in zip_file.infolist():
                with zipfile.ZipFile(self.msword_file, "r") as zip_ref:
                    with zip_ref.open(file_info.filename) as xml_file:
                        if self.hashing:  # if hashing option selected
                            md5hash = hashlib.md5(xml_file.read()).hexdigest()
                        else:
                            md5hash = "Option Not Selected"  # else return blank for hash value.
                m_time = file_info.date_time
                if m_time in ((1980, 1, 1, 0, 0, 0), (1980, 0, 0, 0, 0, 0)):
                    modified_time = "nil"
                else:
                    modified_time = dt(*m_time).strftime(__dtfmt__)
                fname = file_info.filename
                if fname not in self.extra_fields:
                    fname = fname.replace("/", "\\")
                xml_files[file_info.filename] = [
                    md5hash,
                    modified_time,
                    file_info.file_size,
                    file_info.compress_type,
                    file_info.create_system,
                    file_info.create_version,
                    file_info.extract_version,
                    f"{file_info.flag_bits:#0{6}x}",
                    self.extra_fields[fname][0],
                    self.extra_fields[fname][1],
                ]
            return (
                xml_files  # returns dictionary {xml_filename: [file size, file hash]}
            )

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
        return "" if root is None else root

    def doc_ids(self):
        """
        :return: the w14, w15, and w16 docId's from settings.xml
        """
        x = ET.fromstring(self.settings_xml_content)
        w14_ns = x.find(f"{{{self.namespaces['w14']}}}docId")
        w14_id = (
            w14_ns.get(f"{{{self.namespaces['w14']}}}val", "")
            if w14_ns is not None
            else ""
        )
        w15_ns = x.find(f"{{{self.namespaces['w15']}}}docId")
        w15_id = (
            w15_ns.get(f"{{{self.namespaces['w15']}}}val", "")
            if w15_ns is not None
            else ""
        )
        w16_ns = x.find(f"{{{self.namespaces['w16']}}}docId")
        w16_id = (
            w16_ns.get(f"{{{self.namespaces['w16']}}}val", "")
            if w16_ns is not None
            else ""
        )
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


def process_docx(filename):
    """
    This function accepts a filename of type Docx and processes it.
    By placing this in a function, it allows the main part of the script to accept multiple file names and
    then loop through them, calling this function for each DOCx file.
    """
    update_status = ms_word_form.update_status
    excel_file_path = ms_word_form.excel_full_path
    triage = ms_word_form.triageButton.isChecked()
    hashing = ms_word_form.hashFiles.isChecked()
    global doc_summary_worksheet, metadata_worksheet, archive_files_worksheet, rsids_worksheet, comments_worksheet
    update_status(f"Processing {filename.msword_file}")
    file_details = filename.details()
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
        xml_exists = checkFile in filename.xml_files().keys()
        update_status(f"    {checkFile} exists: {xml_exists}")

    # Writing document summary worksheet.

    headers = [
        "File Name",
        "MD5 Hash",
        "Unique rsidR",
        "RSID Root",
        "<w:p> tags",
        "<w:r> tags",
        "<w:t> tags",
        "<w14:docId>",
        "<w15:docId>",
        "<w16:docId>",
        "Hyperlinks",
    ]
    if not hashing:
        headers.pop(1)
    if not bool(
        doc_summary_worksheet
    ):  # if it's an empty dictionary, add headers to it.
        doc_summary_worksheet = dict((k, []) for k in headers)
    w14_id, w15_id, w16_id = filename.doc_ids()
    doc_summary_worksheet["File Name"].append(filename.filename())
    if hashing:
        doc_summary_worksheet["MD5 Hash"].append(filename.hash())
    doc_summary_worksheet["Unique rsidR"].append(len(filename.rsidr()))
    doc_summary_worksheet["RSID Root"].append(filename.rsid_root())
    doc_summary_worksheet["<w:p> tags"].append(filename.paragraph_tags())
    doc_summary_worksheet["<w:r> tags"].append(filename.runs_tags())
    doc_summary_worksheet["<w:t> tags"].append(filename.text_tags())
    doc_summary_worksheet["<w14:docId>"].append(w14_id)
    doc_summary_worksheet["<w15:docId>"].append(w15_id)
    doc_summary_worksheet["<w16:docId>"].append(w16_id)
    doc_summary_worksheet["Hyperlinks"].append(filename.hyperlinks())

    update_status("    Extracted Doc_Summary artifacts")

    # The keys will be used as the column heading in the spreadsheet
    # The order they are in is the order that the columns will be in the spreadsheet
    # Corresponding values passed, resulting in a dictionary being passed called allMetadata
    # containing column headings and associated extracted metadata value.

    headers = [
        "File Name",
        "Author",
        "Created Date",
        "Last Modified By",
        "Modified Date",
        "Last Printed Date",
        "Manager",
        "Company",
        "Revision",
        "Total Editing Time",
        "Pages",
        "Paragraphs",
        "Lines",
        "Words",
        "Characters",
        "Characters With Spaces",
        "Title",
        "Subject",
        "Keywords",
        "Description",
        "Application",
        "App Version",
        "Template",
        "Doc Security",
        "Category",
        "Content Status",
        "RSID Root",
        "Language",
        "Version",
        "Shared Doc",
        "Hyperlinks Changed",
    ]

    if not bool(metadata_worksheet):  # if it's an empty dictionary, add headers to it.
        metadata_worksheet = dict((k, []) for k in headers)

    metadata_worksheet[headers[0]].append(filename.filename())
    metadata_worksheet[headers[1]].append(filename.get_metadata("creator"))
    metadata_worksheet[headers[2]].append(filename.get_metadata("created"))
    metadata_worksheet[headers[3]].append(filename.get_metadata("lastModifiedBy"))
    metadata_worksheet[headers[4]].append(filename.get_metadata("modified"))
    metadata_worksheet[headers[5]].append(filename.get_metadata("lastPrinted"))
    metadata_worksheet[headers[6]].append(filename.get_metadata("Manager"))
    metadata_worksheet[headers[7]].append(filename.get_metadata("Company"))
    metadata_worksheet[headers[8]].append(filename.get_metadata("revision"))
    metadata_worksheet[headers[9]].append(filename.get_metadata("TotalTime"))
    metadata_worksheet[headers[10]].append(filename.get_metadata("Pages"))
    metadata_worksheet[headers[11]].append(filename.get_metadata("Paragraphs"))
    metadata_worksheet[headers[12]].append(filename.get_metadata("Lines"))
    metadata_worksheet[headers[13]].append(filename.get_metadata("Words"))
    metadata_worksheet[headers[14]].append(filename.get_metadata("Characters"))
    metadata_worksheet[headers[15]].append(
        filename.get_metadata("CharactersWithSpaces")
    )
    metadata_worksheet[headers[16]].append(filename.get_metadata("title"))
    metadata_worksheet[headers[17]].append(filename.get_metadata("subject"))
    metadata_worksheet[headers[18]].append(filename.get_metadata("keywords"))
    metadata_worksheet[headers[19]].append(filename.get_metadata("description"))
    metadata_worksheet[headers[20]].append(filename.get_metadata("Application"))
    metadata_worksheet[headers[21]].append(filename.get_metadata("AppVersion"))
    metadata_worksheet[headers[22]].append(filename.get_metadata("Template"))
    metadata_worksheet[headers[23]].append(filename.get_metadata("DocSecurity"))
    metadata_worksheet[headers[24]].append(filename.get_metadata("category"))
    metadata_worksheet[headers[25]].append(filename.get_metadata("contentStatus"))
    metadata_worksheet[headers[26]].append(filename.rsid_root())
    metadata_worksheet[headers[27]].append(filename.get_metadata("language"))
    metadata_worksheet[headers[28]].append(filename.get_metadata("version"))
    metadata_worksheet[headers[29]].append(filename.get_metadata("SharedDoc"))
    metadata_worksheet[headers[30]].append(filename.get_metadata("HyperlinksChanged"))

    update_status("    Extracted metadata artifacts")

    if filename.any_comments():  # checks if there are comments
        headers = [
            "File Name",
            "Comment ID #",
            "Timestamp (UTC)",
            "Author",
            "Initials",
            "Comment",
        ]
        if not bool(
            comments_worksheet
        ):  # if it's an empty dictionary, add headers to it.
            comments_worksheet = dict((k, []) for k in headers)

        for comment in filename.get_comments():
            update_status(f"    Processing comment: {comment}", level="debug")
            comments_worksheet[headers[0]].append(filename.filename())  # Filename
            comments_worksheet[headers[1]].append(comment[0])  # ID
            comments_worksheet[headers[2]].append(comment[1])  # Timestamp
            comments_worksheet[headers[3]].append(comment[2])  # Author
            comments_worksheet[headers[4]].append(comment[3])  # Initials
            comments_worksheet[headers[5]].append(comment[4])  # Text

        update_status("    Extracted comments artifacts")

    if not triage:  # will generate these spreadsheet if not triage
        update_status(
            f'    Updating "Archive Files" worksheet in {excel_file_path}',
            level="debug",
        )
        # Writing XML files to "Archive Files" worksheet
        headers = [
            "File Name",
            "Archive File",
            "MD5 Hash",
            "Modified Time (local/UTC/Redmond, Washington)",
            # expressed local time if Mac/iOS Pages exported to MS Word
            # expressed in UTC if created by LibreOffice on Windows exporting to MS Word.
            # expressed Redmond, Washington time zone when edited with MS Word online.
            "Size (bytes)",
            "ZIP Compression Type",
            "ZIP Create System",
            "ZIP Created Version",
            "ZIP Extract Version",
            "ZIP Flag Bits (hex)",
            "ZIP Extra Flag (len)",
            "ZIP Extra Characters (truncated)",
        ]

        if not bool(
            archive_files_worksheet
        ):  # if it's an empty dictionary, add headers to it.
            archive_files_worksheet = dict((k, []) for k in headers)

        for xml, xml_info in filename.xml_files().items():
            extra_characters = (
                xml_info[9] if xml_info[8] == 0 else ",".join(xml_info[9])
            )  # If no extra characters,
            # leave assigned value as "nil". Otherwise, join.

            archive_files_worksheet[headers[0]].append(filename.filename())
            archive_files_worksheet[headers[1]].append(xml)
            archive_files_worksheet[headers[2]].append(xml_info[0])
            archive_files_worksheet[headers[3]].append(xml_info[1])
            archive_files_worksheet[headers[4]].append(xml_info[2])
            archive_files_worksheet[headers[5]].append(xml_info[3])
            archive_files_worksheet[headers[6]].append(xml_info[4])
            archive_files_worksheet[headers[7]].append(xml_info[5])
            archive_files_worksheet[headers[8]].append(xml_info[6])
            archive_files_worksheet[headers[9]].append(xml_info[7])
            archive_files_worksheet[headers[10]].append(xml_info[8])
            archive_files_worksheet[headers[11]].append(extra_characters)

        update_status("    Extracted archive files artifacts")

        # Calculating count of rsidR, rsidRPr, rsidP, rsidRDefault, paraId, and textId in document.xml
        # and writing to "rsids" worksheet
        headers = [
            "File Name",
            "RSID Type",
            "RSID Value",
            "Count in document.xml",
            "RSID Root",
        ]

        if not bool(rsids_worksheet):  # if it's an empty dictionary, add headers to it.
            rsids_worksheet = dict((k, []) for k in headers)

        update_status("    Calculating rsidR count")
        for k, v in filename.rsidr_in_document_xml().items():
            rsids_worksheet[headers[0]].append(filename.filename())
            rsids_worksheet[headers[1]].append("rsidR")
            rsids_worksheet[headers[2]].append(k)
            rsids_worksheet[headers[3]].append(v)
            rsids_worksheet[headers[4]].append(filename.rsid_root())

        update_status("    Calculating rsidP count")
        for k, v in filename.rsidp_in_document_xml().items():
            rsids_worksheet[headers[0]].append(filename.filename())
            rsids_worksheet[headers[1]].append("rsidP")
            rsids_worksheet[headers[2]].append(k)
            rsids_worksheet[headers[3]].append(v)
            rsids_worksheet[headers[4]].append(filename.rsid_root())

        update_status("    Calculating rsidPr count")
        for k, v in filename.rsidrpr_in_document_xml().items():
            rsids_worksheet[headers[0]].append(filename.filename())
            rsids_worksheet[headers[1]].append("rsidRPr")
            rsids_worksheet[headers[2]].append(k)
            rsids_worksheet[headers[3]].append(v)
            rsids_worksheet[headers[4]].append(filename.rsid_root())

        update_status("    Calculating rsidRDefault count")
        for k, v in filename.rsidrdefault_in_document_xml().items():
            rsids_worksheet[headers[0]].append(filename.filename())
            rsids_worksheet[headers[1]].append("rsidRDefault")
            rsids_worksheet[headers[2]].append(k)
            rsids_worksheet[headers[3]].append(v)
            rsids_worksheet[headers[4]].append(filename.rsid_root())

        update_status("    Calculating paraID count")
        for k, v in filename.paragraph_id_tags().items():
            rsids_worksheet[headers[0]].append(filename.filename())
            rsids_worksheet[headers[1]].append("paraID")
            rsids_worksheet[headers[2]].append(k)
            rsids_worksheet[headers[3]].append(v)
            rsids_worksheet[headers[4]].append(filename.rsid_root())

        update_status("    Calculating textID count")
        for k, v in filename.text_id_tags().items():
            rsids_worksheet[headers[0]].append(filename.filename())
            rsids_worksheet[headers[1]].append("textID")
            rsids_worksheet[headers[2]].append(k)
            rsids_worksheet[headers[3]].append(v)
            rsids_worksheet[headers[4]].append(filename.rsid_root())
    update_status(f"Finished processing {filename.filename()}")
    update_status(f'{"-"*36}')


def main():
    global ms_word_form
    ms_word_app = QApplication([__appname__, "windows:darkmode=2"])
    ms_word_app.setStyle("Fusion")
    ms_word_form = MsWordGui()
    ms_word_form.show()
    ms_word_app.exec()


if __name__ == "__main__":
    main()
