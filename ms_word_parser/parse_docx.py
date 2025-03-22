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
import re
import os
import time
import zipfile
import logging
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
    QGuiApplication,
    QIcon,
    # QImage,
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
timestamp = time.strftime("%Y%m%d_%H%M%S")
log_file = f"DOCx_Parser_Log_{timestamp}.log"
ms_word_form = None
green = QColor(86, 208, 50)
red = QColor(204, 0, 0)
black = QColor(0, 0, 0)
__version__ = "2.0.0"
__appname__ = f"MS Word Parser v{__version__}"
__source__ = "https://github.com/jjrboucher/MS-Word-Parser"
__date__ = "22 March 2025"
__author__ = (
    "Jacques Boucher - jjrboucher@gmail.com\nCorey Forman - corey@digitalsleuth.ca"
)


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
        screen = QApplication.primaryScreen()
        screen_geometry = screen.geometry()
        x = (screen_geometry.width() - self.width()) // 2
        y = (screen_geometry.height() - self.height()) // 2
        self.move(x, y)
        self.setLayout(layout)


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
            datefmt="%Y-%m-%d %H:%M:%S",
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
        self.screen_layout = QGuiApplication.primaryScreen().availableGeometry()
        self.scr_width, self.scr_height = (
            self.screen_layout.width(),
            self.screen_layout.height(),
        )
        self.center_x = (self.scr_width // 2) - (self.d_width // 2)
        self.center_y = (self.scr_height // 2) - (self.d_height // 2)
        self.actionSelect_Excel = QAction(MainWindow)
        self.actionSelect_Excel.setObjectName("actionSelect_Excel")
        self.actionSelect_Excel.triggered.connect(self.open_excel)
        self.actionOpen_File = QAction(MainWindow)
        self.actionOpen_File.setObjectName("actionOpen_File")
        self.actionOpen_File.triggered.connect(self.open_files)
        self.actionOpen_File.setVisible(False)
        self.actionOpen_Directory = QAction(MainWindow)
        self.actionOpen_Directory.setObjectName("actionOpen_Directory")
        self.actionOpen_Directory.triggered.connect(self.open_directory)
        self.actionOpen_Directory.setVisible(False)
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
        self.separator.setStyleSheet(
            """
            QFrame {
                border-top: white;
                border-bottom: white;
                border-left: 1px solid #e4e4e4;
                border-right: 1px solid #e4e4e4;
            }
        """
        )
        self.hashFiles = QCheckBox(self.parsingOptions)
        self.hashFiles.setObjectName("hashFiles")
        self.hashFiles.setGeometry(QRect(250, 30, 75, 20))
        self.hashFiles.setStyleSheet(self.stylesheet)
        self.hashFiles.setFont(self.text_font)
        self.outputFiles = QGroupBox(self.centralwidget)
        self.outputFiles.setObjectName("outputFiles")
        self.outputFiles.setGeometry(QRect(10, 80, 350, 120))
        self.outputFiles.setStyleSheet("background-color: #ffffff; color: black;")
        self.outputFiles.setFont(self.text_font)
        self.excelFileLabel = QLabel(self.outputFiles)
        self.excelFileLabel.setObjectName("excelFileLabel")
        self.excelFileLabel.setGeometry(QRect(10, 30, 80, 16))
        self.excelFileLabel.setStyleSheet("background: #fcfcfc; color: black;")
        self.excelFileLabel.setFont(self.text_font)
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
        self.processStatus = QGroupBox(self.centralwidget)
        self.processStatus.setObjectName("processStatus")
        self.processStatus.setGeometry(QRect(370, 10, 768, 270))
        self.processStatus.setStyleSheet("background: #ffffff; color: black;")
        self.processStatus.setFont(self.text_font)
        self.docxOutput = QTextEdit(self.processStatus)
        self.docxOutput.setObjectName("docxOutput")
        self.docxOutput.setGeometry(QRect(16, 60, 737, 200))
        self.docxOutput.setStyleSheet(self.scrollbar_sheet)
        self.docxOutput.setReadOnly(True)
        self.docxOutput.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )
        self.docxOutput.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.docxOutput.setFont(self.text_font)
        self.operationOptions = QGroupBox(self.centralwidget)
        self.operationOptions.setObjectName("operationOptions")
        self.operationOptions.setGeometry(QRect(10, 210, 350, 70))
        self.operationOptions.setStyleSheet("background-color: #ffffff; color:black;")
        self.operationOptions.setFont(self.text_font)
        self.processButton = QPushButton(self.operationOptions)
        self.processButton.setObjectName("processButton")
        self.processButton.setGeometry(QRect(10, 32, 80, 24))
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
        self.stopButton.setGeometry(QRect(100, 32, 80, 24))
        self.stopButton.setEnabled(False)
        self.stopButton.setStyleSheet(self.disabled)
        self.stopButton.clicked.connect(self._stop)
        self.stopButton.setFont(self.text_font)
        self.resetButton = QPushButton(self.operationOptions)
        self.resetButton.setObjectName("resetButton")
        self.resetButton.setGeometry(QRect(190, 32, 80, 24))
        self.resetButton.clicked.connect(self._reset)
        self.resetButton.setStyleSheet(self.stylesheet)
        self.resetButton.setFont(self.text_font)
        self.numOfFilesLabel = QLabel(self.processStatus)
        self.numOfFilesLabel.setObjectName("numOfFilesLabel")
        self.numOfFilesLabel.setGeometry(QRect(18, 28, 120, 26))
        self.numOfFilesLabel.setStyleSheet("background: #fcfcfc; color: black;")
        self.numOfFilesLabel.setFont(self.text_font)
        self.numOfFiles = QTextEdit(self.processStatus)
        self.numOfFiles.setObjectName("numOfFiles")
        self.numOfFiles.setGeometry(QRect(130, 28, 40, 26))
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
        self.numOfErrorsLabel.setGeometry(QRect(180, 28, 80, 26))
        self.numOfErrorsLabel.setStyleSheet("background: #fcfcfc; color: black;")
        self.numOfErrorsLabel.setFont(self.text_font)
        self.numOfErrors = QTextEdit(self.processStatus)
        self.numOfErrors.setObjectName("numOfErrors")
        self.numOfErrors.setGeometry(QRect(252, 28, 40, 26))
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
        self.numRemainingLabel.setGeometry(QRect(302, 28, 120, 26))
        self.numRemainingLabel.setStyleSheet("background: #fcfcfc; color: black;")
        self.numRemainingLabel.setFont(self.text_font)
        self.numRemaining = QTextEdit(self.processStatus)
        self.numRemaining.setObjectName("numRemaining")
        self.numRemaining.setGeometry(QRect(384, 28, 40, 26))
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
        self.menuFile.addAction(self.actionOpen_File)
        self.menuFile.addAction(self.actionOpen_Directory)
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
        self.actionOpen_File.setText(
            QCoreApplication.translate("MainWindow", "Open &Files ...", None)
        )
        self.actionOpen_Directory.setText(
            QCoreApplication.translate("MainWindow", "Open &Directory ...", None)
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
            QCoreApplication.translate("MainWindow", "File -> Select Excel File", None)
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
        self.openButton.setText(
            QCoreApplication.translate("MainWindow", "Open Output Path", None)
        )
        self.numOfFilesLabel.setText(
            QCoreApplication.translate("MainWindow", "# of Files Selected", None)
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

    def open_directory(self):
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
                recursive_list = list(folder_path.rglob("*.docx"))
                files = [str(file) for file in recursive_list]
            else:
                non_recursive_list = list(folder_path.glob("*.docx"))
                files = [str(file) for file in non_recursive_list]
            self.numOfFiles.setText(str(len(files)))
            self.numRemaining.setText(str(len(files)))
            if files:
                update_status(f"The following {len(files)} files have been loaded:")
                for file in files:
                    update_status(f"    {file}")
                if self.excelFile.toPlainText() != "File -> Select Excel File":
                    self.processButton.setEnabled(True)
                    self.processButton.setStyleSheet(self.stylesheet)
                self.files = files
            else:
                update_status("No files found. Please check your path and try again.")

    def open_files(self):
        update_status = self.update_status
        all_files = []
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select files ...", "", "DOCX Files (*.docx)"
        )
        if files:
            for file in files:
                all_files.append(os.path.normpath(file))
            self.numOfFiles.setText(str(len(all_files)))
            self.numRemaining.setText(str(len(all_files)))
            update_status(f"The following {len(all_files)} files have been loaded:")
            for file in all_files:
                update_status(f"    {file}")
            if self.excelFile.toPlainText() != "File -> Select Excel File":
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
            self.actionOpen_File.setVisible(True)
            self.actionOpen_Directory.setVisible(True)
            self.generalLogFile.setText(log_file)
            self.outputPath.setText(self.excel_path)
            self.openButton.setEnabled(True)
            self.openButton.setStyleSheet(self.stylesheet)

    def open_path(self):
        out_path = self.outputPath.toPlainText().strip()
        if out_path:
            QDesktopServices.openUrl(QUrl.fromLocalFile(out_path))

    def _reset(self):
        global timestamp, log_file
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        log_file = f"DOCx_Parser_Log_{timestamp}.log"
        self.excelFile.setText("File -> Select Excel File")
        self.generalLogFile.setText(log_file)
        self.outputPath.clear()
        self.numOfFiles.setText("0")
        self.numOfErrors.setText("0")
        self.numRemaining.setText("0")
        self.docxOutput.setTextColor(black)
        self.docxOutput.clear()
        self.processButton.setEnabled(False)
        self.processButton.setStyleSheet(self.disabled)
        self.openButton.setEnabled(False)
        self.openButton.setStyleSheet(self.disabled)
        self.actionOpen_File.setVisible(False)
        self.actionOpen_Directory.setVisible(False)
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
            self.logger.info(msg)
        elif level == "error":
            color = red
            self.logger.error(msg)
        self.docxOutput.setTextColor(color)
        self.docxOutput.append(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {msg}")
        self.docxOutput.setTextColor(black)
        QApplication.processEvents()

    def analyze_docs(self, files, triage_files, hash_files):
        if not self.running:
            self.running = True
        self.stopButton.setEnabled(True)
        self.stopButton.setStyleSheet(self.stylesheet)
        self.resetButton.setEnabled(False)
        self.resetButton.setStyleSheet(self.disabled)
        self.processButton.setEnabled(False)
        self.processButton.setStyleSheet(self.disabled)
        docxErrorCount = 0
        update_status = self.update_status
        script_start = time.strftime("%Y-%m-%d %H:%M:%S")
        update_status(f"Script executed: {script_start}")
        update_status("Summary of files parsed:")
        update_status(f'{"="*36}')
        remaining = int(self.numRemaining.toPlainText())
        for f in files:  # loop over the files selected, processing each.
            if not self.running:
                update_status("Processing stopped")
                self.stopButton.setEnabled(False)
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
                )
            if remaining != 0:
                remaining -= 1
            self.numRemaining.setText(str(remaining))
        df_summary = pd.DataFrame(data=doc_summary_worksheet)
        df_metadata = pd.DataFrame(data=metadata_worksheet)
        df_comments = pd.DataFrame(data=comments_worksheet)
        with pd.ExcelWriter(
            path=self.excel_full_path, engine="xlsxwriter", mode="w"
        ) as writer:
            df_summary.to_excel(
                excel_writer=writer, sheet_name="Doc_Summary", index=False
            )
            update_status('"Doc_Summary" worksheet written to Excel.')
            df_metadata.to_excel(
                excel_writer=writer, sheet_name="Metadata", index=False
            )
            update_status('"Metadata" worksheet written to Excel.')
            if not df_comments.empty:
                df_comments.to_excel(
                    excel_writer=writer, sheet_name="Comments", index=False
                )
                update_status('"Comments" worksheet written to Excel.')
            if not triage_files:
                df_archive = pd.DataFrame(data=archive_files_worksheet)
                df_rsids = pd.DataFrame(data=rsids_worksheet)
                df_archive.to_excel(
                    excel_writer=writer, sheet_name="Archive Files", index=False
                )
                update_status('"Archive Files" worksheet written to Excel.')
                df_rsids.to_excel(excel_writer=writer, sheet_name="RSIDs", index=False)
                update_status('"RSIDs" worksheet written to Excel.')
        script_end = time.strftime("%Y-%m-%d %H:%M:%S")
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
            update_status("The following files had errors:", "error")
            for each_file in filesUnableToProcess:
                update_status(f"  {each_file}", "error")
        update_status(f"Script finished execution: {script_end}", color=green)
        self.resetButton.setEnabled(True)
        self.resetButton.setStyleSheet(self.stylesheet)


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

    def __init__(self):
        """Call and setup the UI"""
        super().__init__()
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

        self.msword_file = msword_file
        self.hashing = hashing
        self.header_offsets, self.binary_content = self.__find_binary_string()
        self.extra_fields = self.__xml_extra_bytes()
        self.core_xml_file = "docProps/core.xml"
        self.core_xml_content = self.__load_core_xml()
        if self.core_xml_content == "":
            self.core_xml_file = "docProps\\core.xml"
            self.core_xml_content = self.__load_core_xml()
        self.app_xml_file = "docProps/app.xml"
        self.app_xml_content = self.__load_app_xml()
        if self.app_xml_content == "":
            self.app_xml_file = "docProps\\app.xml"
            self.app_xml_content = self.__load_app_xml()
        self.document_xml_file = "word/document.xml"
        self.document_xml_content = self.__load_document_xml()
        if self.document_xml_content == "":
            self.document_xml_file = "word\\document.xml"
            self.document_xml_content = self.__load_document_xml()
        self.has_comments = ""  # Flag to denote if there are comments in the document.
        self.comments = "word/comments.xml"
        self.comments_xml_content = self.__load_comments_xml()
        if self.comments_xml_content == "":
            self.comments = "word\\comments.xml"
            self.comments_xml_content = self.__load_comments_xml()
        self.settings_xml_file = "word/settings.xml"
        self.settings_xml_content = self.__load_settings_xml()
        if self.settings_xml_content == "":
            self.settings_xml_file = "word\\settings.xml"
            self.settings_xml_content = self.__load_settings_xml()
        self.rsidRs = self.__extract_all_rsidr_from_summary_xml()

        self.p_tags = re.findall(r"<w:p>|<w:p [^>]*/?>", self.document_xml_content)
        self.r_tags = re.findall(r"<w:r>|<w:r [^>]*/?>", self.document_xml_content)
        self.t_tags = re.findall(r"<w:t>|<w:t.? [^>]*/?>", self.document_xml_content)

        if not triage:  # if not run in triage mode, do full parsing

            self.rsidR_in_document_xml = self.__rsidr_in_document_xml()
            self.rsidRPr = self.__other_rsids_in_document_xml("rsidRPr")
            self.rsidP = self.__other_rsids_in_document_xml("rsidP")
            self.rsidRDefault = self.__other_rsids_in_document_xml("rsidRDefault")

            self.para_id = self.__para_id_tags__()
            self.text_id = self.__text_id_tags__()

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

    def __load_core_xml(self):
        # load core.xml
        if (
            self.core_xml_file in self.xml_files()
        ):  # if the file exists, read it and return its content
            with zipfile.ZipFile(self.msword_file, "r") as zipref:
                with zipref.open(self.core_xml_file) as xmlFile:
                    return xmlFile.read().decode("utf-8")
        else:  # if it doesn't exist, return an empty string.
            ms_word_form.update_status(
                f'"{self.core_xml_file}" does not exist in "{self.filename()}". '
                f"Returning empty string."
            )
            return ""

    def __load_app_xml(self):
        # load app.xml
        if (
            self.app_xml_file in self.xml_files()
        ):  # if the file exists, read it and return its content
            with zipfile.ZipFile(self.msword_file, "r") as zipref:
                with zipref.open(self.app_xml_file) as xmlFile:
                    return xmlFile.read().decode("utf-8")
        else:  # if it doesn't exist, return an empty string.
            ms_word_form.update_status(
                f'"{self.app_xml_file}" does not exist in "{self.filename()}". '
                f"Returning empty string."
            )
            return ""

    def __load_document_xml(self):
        # load document.xml
        if (
            self.document_xml_file in self.xml_files()
        ):  # if the file exists, read it and return its content
            with zipfile.ZipFile(self.msword_file, "r") as zipref:
                with zipref.open(self.document_xml_file) as xmlFile:
                    return xmlFile.read().decode("utf-8")
        else:  # if it doesn't exist, return an empty string.
            ms_word_form.update_status(
                f'"{self.document_xml_file}" does not exist in "{self.filename()}". '
                f"Returning empty string."
            )
            return ""

    def __load_settings_xml(self):
        if (
            self.settings_xml_file in self.xml_files()
        ):  # if the file exists, read it and return its content
            with zipfile.ZipFile(self.msword_file, "r") as zipref:
                with zipref.open(self.settings_xml_file) as xmlFile:
                    return xmlFile.read().decode("utf-8")
        else:
            ms_word_form.update_status(
                f'"{self.settings_xml_file}" does not exist in "{self.filename()}". '
                f"Returning empty string."
            )
            return ""

    def __extract_all_rsidr_from_summary_xml(self):
        """
        function to extract all RSIDs at the beginning of the class. If you were to put this in the method,
        it would have to do this every time you called the method.
        :return:
        """
        rsids_list = []
        # Find all RSIDs, not rsidRoot. rsidRoot is repeated in rsids.
        matches = re.findall(
            r'<w:rsid w:val="[0-9A-F]{8}" ?/>', self.settings_xml_content
        )

        for match in matches:  # loops through all matches
            # greps for rsid using a group to extract the actual RSID from the string.
            rsid_match = re.search(r'<w:rsid w:val="([0-9A-F]{8})"', match)
            if rsid_match:
                rsids_list.append(rsid_match.group(1))  # Appends it to the list
        return "" if len(rsids_list) == 0 else rsids_list

    def __rsidr_in_document_xml(self):
        """
        This function calculates the count of each rsidR in document.xml
        It searches the previously extracted tags rather than the full document.
        :return:
        """
        rsidr_count = {}
        for rsid in self.rsidRs:
            pattern = re.compile(rf'w:rsidR="{rsid}"')

            count_rsids = 0

            count_rsids += len(re.findall(pattern, ",".join(self.p_tags)))
            count_rsids += len(re.findall(pattern, ",".join(self.r_tags)))
            count_rsids += len(re.findall(pattern, ",".join(self.t_tags)))

            rsidr_count[rsid] = count_rsids

        return rsidr_count

    def __other_rsids_in_document_xml(self, rsid):
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
        pattern = re.compile("w:" + rsid + '="[0-9A-F]{8}"')
        # Find all rsid types passed to the function (rsidRPr, rsidP, rsidRDefault in document.xml file

        matches = re.findall(pattern, ",".join(self.p_tags))  # searches p_tags
        matches += re.findall(pattern, ",".join(self.r_tags))  # searches r_tags
        matches += re.findall(pattern, ",".join(self.t_tags))  # searches t_tags

        for match in matches:  # loops through all matches
            # greps for rsid using a group to extract the actual RSID from the string.
            group_pattern = rf'w:{rsid}="([0-9A-F]{8})"'
            rsid_match = re.search(group_pattern, match)
            if rsid_match:
                if rsid_match.group(1) in rsids:
                    rsids[rsid_match.group(1)] += 1  # increment count by 1
                else:
                    rsids[rsid_match.group(1)] = 1  # Appends it to the list

        return rsids

    def __para_id_tags__(self):
        """
        :return: list of unique paraId tags and count in document.xml
        """
        pid_tags = {}  # empty dictionary to start

        for pid_tag in self.p_tags:
            pidtag = re.search(r'paraId="([0-9A-F]{8})"', pid_tag)
            if pidtag is None:  # no paraId= tag in this <w:p> paragraph tag.
                pass
            elif pidtag.group(1) in pid_tags:
                pid_tags[pidtag.group(1)] += 1  # increment count by 1
            else:
                pid_tags[pidtag.group(1)] = 1  # append to the list

        return pid_tags

    def __text_id_tags__(self):
        """
        :return: list of unique paraId tags and count in document.xml
        """
        text_tags = {}  # empty dictionary to start

        for text_tag in self.p_tags:
            texttag = re.search(r'textId="([0-9A-F]{8})"', text_tag)
            if texttag is None:  # no paraId= tag in this <w:p> paragraph tag.
                pass
            elif texttag.group(1) in text_tags:
                text_tags[texttag.group(1)] += 1  # increment count by 1
            else:
                text_tags[texttag.group(1)] = 1  # append to the list

        return text_tags

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
        month = {
            1: "Jan",
            2: "Feb",
            3: "Mar",
            4: "Apr",
            5: "May",
            6: "Jun",
            7: "Jul",
            8: "Aug",
            9: "Sep",
            10: "Oct",
            11: "Nov",
            12: "Dec",
        }
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
                    modified_time = f"{m_time[0]}-{month[m_time[1]]}-{m_time[2]:02d} {m_time[3]:02d}:{m_time[4]:02d}:{m_time[5]:02d}"  ##TODO, fix timestamp formatting
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

    def xml_hash(self, xmlfile):
        """
        :param xmlfile
        :return: the hash of a specified XML file
        """
        return self.xml_files()[xmlfile][1]

    def xml_size(self, xmlfile):
        """
        :param xmlfile
        :return: the size of a specified XML file
        """
        return self.xml_files()[xmlfile][0]

    def title(self):
        """
        :return: the title metadata in core.xml
        """
        doc_title = re.search(
            r"<.{0,2}:?title>(.*?)</.{0,2}:?title>", self.core_xml_content
        )
        return "" if doc_title is None else doc_title.group(1)

    def subject(self):
        """
        :return: the subject metadata from core.xml
        """
        doc_subject = re.search(
            r"<.{0,2}:?subject>(.*?)</.{0,2}:?subject>", self.core_xml_content
        )
        return "" if doc_subject is None else doc_subject.group(1)

    def creator(self):
        """
        :return: the creator metadata from core.xml
        """
        doc_creator = re.search(
            r"<.{0,2}:?creator>(.*?)</.{0,2}:?creator>", self.core_xml_content
        )
        return "" if doc_creator is None else doc_creator.group(1)

    def keywords(self):
        """
        :return: the keywords metadata from core.xml
        """
        doc_keywords = re.search(
            r"<.{0,2}:?keywords>(.*?)</.{0,2}:?keywords>", self.core_xml_content
        )
        return "" if doc_keywords is None else doc_keywords.group(1)

    def description(self):
        """
        :return: the description metadata from core.xml
        """
        doc_description = re.search(
            r"<.{0,2}:?description>(.*?)</.{0,2}:?description>", self.core_xml_content
        )
        return "" if doc_description is None else doc_description.group(1)

    def revision(self):
        """
        :return: the revision # metadata from core.xml
        """
        doc_revision = re.search(
            r"<.{0,2}:?revision>(.*?)</.{0,2}:?revision>", self.core_xml_content
        )
        return "" if doc_revision is None else doc_revision.group(1)

    def created(self):
        """
        :return: the created date metadata from core.xml
        """
        doc_created = re.search(
            r"<dcterms:created[^>].*?>(.*?)</dcterms:created>", self.core_xml_content
        )
        return "" if doc_created is None else doc_created.group(1)

    def modified(self):
        """
        :return: the modified date metadata from core.xml
        """
        doc_modified = re.search(
            r"<dcterms:modified[^>].*?>(.*?)</dcterms:modified>", self.core_xml_content
        )
        return "" if doc_modified is None else doc_modified.group(1)

    def last_modified_by(self):
        """
        :return: the last modified by metadata from core.xml
        """
        doc_lastmodifiedby = re.search(
            r"<.{0,2}:?lastModifiedBy>(.*?)</.{0,2}:?lastModifiedBy>",
            self.core_xml_content,
        )
        return "" if doc_lastmodifiedby is None else doc_lastmodifiedby.group(1)

    def last_printed(self):
        """
        :return: the last printed date metadata from core.xml
        """
        doc_lastprinted = re.search(
            r"<.{0,2}:?lastPrinted>(.*?)</.{0,2}:?lastPrinted>", self.core_xml_content
        )
        return "" if doc_lastprinted is None else doc_lastprinted.group(1)

    def category(self):
        """
        :return: the category metadata from core.xml
        """
        doc_category = re.search(
            r"<.{0,2}:?category>(.*?)</.{0,2}:?category>", self.core_xml_content
        )
        return "" if doc_category is None else doc_category.group(1)

    def content_status(self):
        """
        :return: the content status metadata from core.xml
        """
        doc_contentstatus = re.search(
            r"<.{0,2}:?contentStatus>(.*?)</.{0,2}:?contentStatus>",
            self.core_xml_content,
        )
        return "" if doc_contentstatus is None else doc_contentstatus.group(1)

    def template(self):
        """
        :return: the template metadata from app.xml
        """
        doc_template = re.search(
            r"<.{0,2}:?Template>(.*?)</.{0,2}:?Template>", self.app_xml_content
        )
        return "" if doc_template is None else doc_template.group(1)

    def total_editing_time(self):
        """
        :return: the total editing time in minutes metadata from app.xml
        """
        doc_edit_time = re.search(
            r"<.{0,2}:?TotalTime>(.*?)</.{0,2}:?TotalTime>", self.app_xml_content
        )
        return "" if doc_edit_time is None else doc_edit_time.group(1)

    def pages(self):
        """
        :return: the # of pages in the document metadata from app.xml
        Note: the author has observed that in some cases, this is not properly updated within the XML file itself.
        It is not an error in the script. It's an error in the metadata. Opening the document and allowing it to
        fully load and then saving it updates this. But of course, it changes other metadata as well if you do that.
        """
        doc_pages = re.search(
            r"<.{0,2}:?Pages>(.*?)</.{0,2}:?Pages>", self.app_xml_content
        )
        return "" if doc_pages is None else doc_pages.group(1)

    def words(self):
        """
        :return: the number of words in the document metadata from app.xml
        """
        doc_words = re.search(
            r"<.{0,2}:?Words>(.*?)</.{0,2}:?Words>", self.app_xml_content
        )
        return "" if doc_words is None else doc_words.group(1)

    def characters(self):
        """
        :return: the number of characters in the document metadata from app.xml
        """
        doc_characters = re.search(
            r"<.{0,2}:?Characters>(.*?)</.{0,2}:?Characters>", self.app_xml_content
        )
        return "" if doc_characters is None else doc_characters.group(1)

    def application(self):
        """
        :return: the application name that created the document metadata from app.xml
        """
        doc_application = re.search(
            r"<.{0,2}:?Application>(.*?)</.{0,2}:?Application>", self.app_xml_content
        )
        return "" if doc_application is None else doc_application.group(1)

    def security(self):
        """
        :return: the security metadata from app.xml
        """
        doc_security = re.search(
            r"<.{0,2}:?DocSecurity>(.*?)</.{0,2}:?DocSecurity>", self.app_xml_content
        )
        return "" if doc_security is None else doc_security.group(1)

    def lines(self):
        """
        :return: the number of lines in the document metadata from app.xml
        """
        doc_lines = re.search(
            r"<.{0,2}:?Lines>(.*?)</.{0,2}:?Lines>", self.app_xml_content
        )
        return "" if doc_lines is None else doc_lines.group(1)

    def paragraphs(self):
        """
        :return: the number of paragraphs in the document metadata from app.xml
        Note: similar to # of pages, the author has noted in testing that sometimes, this may not be accurate in
        the metadata for some reason. It's not an error in this program. It's an error with the metadata itself
        in the document.
        """
        doc_paragraphs = re.search(
            r"<.{0,2}:?Paragraphs>(.*?)</.{0,2}:?Paragraphs>", self.app_xml_content
        )
        return "" if doc_paragraphs is None else doc_paragraphs.group(1)

    def characters_with_spaces(self):
        """
        :return: the total characters including spaces in the document metadatafrom app.xml
        """
        doc_characters_with_spaces = re.search(
            r"<.{0,2}:?CharactersWithSpaces>(.*?)</.{0,2}:?CharactersWithSpaces>",
            self.app_xml_content,
        )
        return (
            ""
            if doc_characters_with_spaces is None
            else doc_characters_with_spaces.group(1)
        )

    def app_version(self):
        """
        :return: the version of the app that created the document metadatafrom app.xml
        """
        doc_app_version = re.search(
            r"<.{0,2}:?AppVersion>(.*?)</.{0,2}:?AppVersion>", self.app_xml_content
        )
        return "" if doc_app_version is None else doc_app_version.group(1)

    def manager(self):
        """
        :return: the manager metadata from app.xml
        """
        doc_manager = re.search(
            r"<.{0,2}:?Manager>(.*?)</.{0,2}:?Manager>", self.app_xml_content
        )
        return "" if doc_manager is None else doc_manager.group(1)

    def company(self):
        """
        :return: the company metadata from app.xml
        """
        doc_company = re.search(
            r"<.{0,2}:?Company>(.*?)</.{0,2}:?Company>", self.app_xml_content
        )
        return "" if doc_company is None else doc_company.group(1)

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
        root = re.search(r'<w:rsidRoot w:val="([^"]*)"', self.settings_xml_content)
        return "" if root is None else root.group(1)

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

    def __load_comments_xml(self):
        # load comments.xml
        if (
            self.comments in self.xml_files()
        ):  # if the file exists, read it and return its content
            self.has_comments = True
            with zipfile.ZipFile(self.msword_file, "r") as zipref:
                with zipref.open(self.comments) as xmlFile:
                    return xmlFile.read().decode("utf-8")
        else:  # if it doesn't exist, return an empty string.
            self.has_comments = False
            return ""

    def get_comments(self):
        """ "
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

        namespaces = {
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",  # Main namespace
            "w14": "http://schemas.microsoft.com/office/word/2010/wordml",  # Other used namespace
            "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        }

        # Find all comments
        comments = xml.findall(".//w:comment", namespaces)

        all_comments = []  # list to contain all comments

        for comment in comments:
            author = comment.get("{" + namespaces["w"] + "}" + "author")
            date_time = comment.get("{" + namespaces["w"] + "}" + "date")
            initials = comment.get("{" + namespaces["w"] + "}" + "initials")
            comment_id = comment.get("{" + namespaces["w"] + "}" + "id")
            text = "".join([t.text for t in comment.findall(".//w:t", namespaces)])

            all_comments.append([comment_id, date_time, author, initials, text])

        return all_comments

    def any_comments(self):
        return self.has_comments

    def details(self):
        """
        :return: a text string that you can print out to get a summary of the document.
        This can be edited to suit your needs. You can naturally accomplish the same results by calling each of
        the methods in your print statement in the main script.
        """
        if self.last_printed() == "":
            printed = "Document was never printed"
        else:
            printed = f"Printed: {self.last_printed()}"
        return (
            f"Document: {self.filename()}\n"
            f"Created by: {self.creator()}\n"
            f"Created date: {self.created()}\n"
            f"Last edited by: {self.last_modified_by()}\n"
            f"Edited date: {self.modified()}\n"
            f"{printed}\n"
            f"Total pages: {self.pages()}\n"
            f"Total editing time: {self.total_editing_time()} minute(s)."
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
    ]

    if not bool(
        doc_summary_worksheet
    ):  # if it's an empty dictionary, add headers to it.
        doc_summary_worksheet = dict((k, []) for k in headers)

    doc_summary_worksheet[headers[0]].append(filename.filename())
    if hashing:
        doc_summary_worksheet[headers[1]].append(filename.hash())
    else:
        doc_summary_worksheet[headers[1]].append("Option Not Selected")
    doc_summary_worksheet[headers[2]].append(len(filename.rsidr()))
    doc_summary_worksheet[headers[3]].append(filename.rsid_root())
    doc_summary_worksheet[headers[4]].append(filename.paragraph_tags())
    doc_summary_worksheet[headers[5]].append(filename.runs_tags())
    doc_summary_worksheet[headers[6]].append(filename.text_tags())

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
    ]

    if not bool(metadata_worksheet):  # if it's an empty dictionary, add headers to it.
        metadata_worksheet = dict((k, []) for k in headers)

    metadata_worksheet[headers[0]].append(filename.filename())
    metadata_worksheet[headers[1]].append(filename.creator())
    metadata_worksheet[headers[2]].append(filename.created())
    metadata_worksheet[headers[3]].append(filename.last_modified_by())
    metadata_worksheet[headers[4]].append(filename.modified())
    metadata_worksheet[headers[5]].append(filename.last_printed())
    metadata_worksheet[headers[6]].append(filename.manager())
    metadata_worksheet[headers[7]].append(filename.company())
    metadata_worksheet[headers[8]].append(filename.revision())
    metadata_worksheet[headers[9]].append(filename.total_editing_time())
    metadata_worksheet[headers[10]].append(filename.pages())
    metadata_worksheet[headers[11]].append(filename.paragraphs())
    metadata_worksheet[headers[12]].append(filename.lines())
    metadata_worksheet[headers[13]].append(filename.words())
    metadata_worksheet[headers[14]].append(filename.characters())
    metadata_worksheet[headers[15]].append(filename.characters_with_spaces())
    metadata_worksheet[headers[16]].append(filename.title())
    metadata_worksheet[headers[17]].append(filename.subject())
    metadata_worksheet[headers[18]].append(filename.keywords())
    metadata_worksheet[headers[19]].append(filename.description())
    metadata_worksheet[headers[20]].append(filename.application())
    metadata_worksheet[headers[21]].append(filename.app_version())
    metadata_worksheet[headers[22]].append(filename.template())
    metadata_worksheet[headers[23]].append(filename.security())
    metadata_worksheet[headers[24]].append(filename.category())
    metadata_worksheet[headers[25]].append(filename.content_status())

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
            update_status(f"    Processing comment: {comment}")
            comments_worksheet[headers[0]].append(filename.filename())  # Filename
            comments_worksheet[headers[1]].append(comment[0])  # ID
            comments_worksheet[headers[2]].append(comment[1])  # Timestamp
            comments_worksheet[headers[3]].append(comment[2])  # Author
            comments_worksheet[headers[4]].append(comment[3])  # Initials
            comments_worksheet[headers[5]].append(comment[4])  # Text

        update_status("    Extracted comments artifacts")

    if not triage:  # will generate these spreadsheet if not triage
        update_status(f'    Updating "Archive Files" worksheet in "{excel_file_path}"')
        # Writing XML files to "Archive Files" worksheet
        headers = [
            "File Name",
            "Archive File",
            "MD5Hash",
            "Modified Time (local/UTC/Redmond, Washington)",
            # expressed local time if Mac/iOS Pages exported to MS Word
            # expressed in UTC if created by LibreOffice on Windows exportinug to MS Word.
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
        headers = ["File Name", "RSID Type", "RSID Value", "Count in document.xml"]

        if not bool(rsids_worksheet):  # if it's an empty dictionary, add headers to it.
            rsids_worksheet = dict((k, []) for k in headers)

        update_status("    Calculating rsidR count")
        for k, v in filename.rsidr_in_document_xml().items():
            rsids_worksheet[headers[0]].append(filename.filename())
            rsids_worksheet[headers[1]].append("rsidR")
            rsids_worksheet[headers[2]].append(k)
            rsids_worksheet[headers[3]].append(v)

        update_status("    Calculating rsidP count")
        for k, v in filename.rsidp_in_document_xml().items():
            rsids_worksheet[headers[0]].append(filename.filename())
            rsids_worksheet[headers[1]].append("rsidP")
            rsids_worksheet[headers[2]].append(k)
            rsids_worksheet[headers[3]].append(v)

        update_status("    Calculating rsidPr count")
        for k, v in filename.rsidrpr_in_document_xml().items():
            rsids_worksheet[headers[0]].append(filename.filename())
            rsids_worksheet[headers[1]].append("rsidRPr")
            rsids_worksheet[headers[2]].append(k)
            rsids_worksheet[headers[3]].append(v)

        update_status("    Calculating rsidRDefault count")
        for k, v in filename.rsidrdefault_in_document_xml().items():
            rsids_worksheet[headers[0]].append(filename.filename())
            rsids_worksheet[headers[1]].append("rsidRDefault")
            rsids_worksheet[headers[2]].append(k)
            rsids_worksheet[headers[3]].append(v)

        update_status("    Calculating paraID count")
        for k, v in filename.paragraph_id_tags().items():
            rsids_worksheet[headers[0]].append(filename.filename())
            rsids_worksheet[headers[1]].append("paraID")
            rsids_worksheet[headers[2]].append(k)
            rsids_worksheet[headers[3]].append(v)

        update_status("    Calculating textID count")
        for k, v in filename.text_id_tags().items():
            rsids_worksheet[headers[0]].append(filename.filename())
            rsids_worksheet[headers[1]].append("textID")
            rsids_worksheet[headers[2]].append(k)
            rsids_worksheet[headers[3]].append(v)
    update_status(f"Finished processing {filename.msword_file}")
    update_status(f'{"-"*36}')


def main():
    global ms_word_form
    ms_word_app = QApplication([__appname__, "windows:darkmode=2"])
    # ms_word_app.setWindowIcon(QIcon("logo.ico"))
    ms_word_app.setStyle("Fusion")
    ms_word_form = MsWordGui()
    ms_word_form.show()
    ms_word_app.exec()


if __name__ == "__main__":
    main()
    # ms_word_app = QApplication([__appname__, "windows:darkmode=2"])
    ## ms_word_app.setWindowIcon(QIcon("logo.ico"))
    # ms_word_app.setStyle("Fusion")
    # ms_word_form = MsWordGui()
    # ms_word_form.show()
    # ms_word_app.exec()
