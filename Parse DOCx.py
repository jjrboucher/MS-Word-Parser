####################################
# Written by Jacques Boucher
# jjrboucher@gmail.com
#
# Version Date: 23 September 2023
#
# Written in Python 3.11
#
# ********** Description **********
#
# Script will open a windows dialog to allow you to select a DOCx file.
# The script does not attempt to validate the file.
# A docx file is nothing more than a ZIP file, hence why this script uses the zipfile library.
#
# It will extract the results to a file called docx-artifacts.xlsx as defined by the variable excel_file_path at the
# start of the main part of the script.
# If the file does not exist, it creates it. If the file does exist, it appends to it.
# The file will be located in the folder where the script is executed from.
# If executing from the GUI by double-clicking on the .py file, it should be stored in that same folder.
# If executing it from the command line, it will create it in whichever folder you are in when executing it.
#
# This allows you to run this repeatedly against many DOCx file for an investigation and compare them.
# You can then copy/move/rename the default file, docx-artifacts.xlsx to another file name for your case.
#
# Processes that this script will do:
#
# 1 - It will extract a list of all the files within the zip file and save it to a worksheet called XML_files.
#     In this worksheet, it will save the following information to a row:
#     "File Name", "XML", "Size (bytes)", "MD5Hash"
#
# 2 - It will extract all the unique RSIDs from the file word/settings.xml and write it to a worksheet
#     called doc_summary.
#     In this worksheet, it will save the following information to a row:
#     "File Name", "Unique RSIDs", "RSID Root", "<w:p> tags", "<w:r> tags", "<w:t> tags"
#     Where "Unique RSID" is a numerical count of the # of RSIDs in the file.
#
#     What is an RSID (Revision Save ID)?
#     See https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.rsid?view=openxml-2.8.1
#
# 3 - It will extract all the unique RSIDs from the file word/settings.xml and write it to a worksheet called rsids.
#     In this worksheet, it will save the following information to rows (one for each unique RSID):
#     "File Name", "RSID"
#
# 3 - It will extract all known relevant metadata from the files docProps/app.xml and docProps/core.xml
#     and write it to a worksheet called metadata.
#     In this worksheet, it will save the following information to a row:
#     "File Name", "Author", "Created Date","Last Modified By","Modified Date","Last Printed Date","Manager","Company",
#     "Revision","Total Editing Time","Pages","Paragraphs","Lines","Words","Characters","Characters With Spaces",
#     "Title","Subject","Keywords","Description","Application","App Version","Template","Doc Security","Category",
#     "contentStatus"
#
#
# ********** Dependencies **********
#
# If running the script on a Linux system, you may need to install python-tk. You can do this with the following
# command on a Debian (e.g. Ubuntu) system from the terminal window:
# sudo apt-get install python3-tk
#
# Whether running on Linux, Mac, or Windows, you may need to install some of the libraries if they are not included in
# your installation of Python 3.
# In particular, you may need to install openpyxl and hashlib. You can do so as follows from a terminal window:
#
# pip3 install openpyxl
# pip3 install hashlib
#
# If any other libraries are missing when trying to execute the script, install those in the same manner.
#
###################################

import os
import tkinter as tk
from tkinter import filedialog
from functions.metadata import core_xml, app_xml  # functions to extract metadata from core.xml and app.xml
from functions.excel import write_worksheet  # function to write results to an Excel file
from functions.rsids import extract_rsids_from_settings_xml  # function to extract rsids and rsidRoot from settings.xml
from functions.xml import list_of_xml_files  # function to return list of xml files in a DOCx file.
from functions.xml import extract_content_of_xml  # function to read an XML file and return as utf-8 text.
from functions.extracttags import extract_tags_from_document_xml  # extracts count of p, r, and t tags


if __name__ == "__main__":

    # Output file - same path as where the script is run. It will create it if it does not exist,
    # or append to it if it does.
    excel_file_path = "docx-artifacts.xlsx"  # default file name - will be created in the script folder.

    root = tk.Tk()
    root.withdraw()  # Hide the main window

    msword_file_path = filedialog.askopenfilename(title="Select DOCx file to process", initialdir=".",
                                                  filetypes=[("DOCx Files", "*.docx")])
    if not msword_file_path:
        print("No DOCx file selected. Exiting.")
    else:

        filename = os.path.basename(msword_file_path)

        # Executes the function to get a list of all XML files in DOCx file
        XMLFiles = list_of_xml_files(msword_file_path)

        # parse word/settings.xml artifacts
        xml_file_path_within_zip = "word/settings.xml"  # Path of the XML file within the ZIP
        # Executes the function to get rsids and rsidRoot from settings.xml
        rsids, rsidRoot = extract_rsids_from_settings_xml(extract_content_of_xml(msword_file_path,
                                                                                 xml_file_path_within_zip))

        # parse docProps/app.xml artifacts
        xml_file_path_within_zip = "docProps/app.xml"  # Path of the XML file within the ZIP
        # Executes the function to get metadata from app.xml
        app_xml_metadata = app_xml(extract_content_of_xml(msword_file_path, xml_file_path_within_zip))

        # parse docProps/core.xml artifacts
        xml_file_path_within_zip = "docProps/core.xml"  # Path of the XML file within the ZIP
        # Executes the function to get the metadata from core.xml
        core_xml_metadata = core_xml(extract_content_of_xml(msword_file_path, xml_file_path_within_zip))

        # parse word/document.xml artifacts
        xml_file_path_within_zip = "word/document.xml"  # Path of the XML file within the ZIP
        # Executes the function to get the metadata from document.xml
        documentXMLTagSummary = extract_tags_from_document_xml(extract_content_of_xml
                                                               (msword_file_path, xml_file_path_within_zip))

        # The keys will be used as the column heading in the spreadsheet
        # The order they are in is the order that the columns will be in the spreadsheet
        # Corresponding values passed, resulting in a dictionary being passed called allMetadata
        # containing column headings and associated extracted metadata value.
        allMetadata = {"File Name": filename,
                       "Author": core_xml_metadata["creator"],
                       "Created Date": core_xml_metadata["created"],
                       "Last Modified By": core_xml_metadata["lastModifiedBy"],
                       "Modified Date": core_xml_metadata["modified"],
                       "Last Printed Date": core_xml_metadata["lastPrinted"],
                       "Manager": app_xml_metadata["manager"],
                       "Company": app_xml_metadata["company"],
                       "Revision": core_xml_metadata["revision"],
                       "Total Editing Time": app_xml_metadata["totalTime"],
                       "Pages": app_xml_metadata["pages"],
                       "Paragraphs": app_xml_metadata["paragraphs"],
                       "Lines": app_xml_metadata["lines"],
                       "Words": app_xml_metadata["words"],
                       "Characters": app_xml_metadata["characters"],
                       "Characters With Spaces": app_xml_metadata["charactersWithSpaces"],
                       "Title": core_xml_metadata["title"],
                       "Subject": core_xml_metadata["subject"],
                       "Keywords": core_xml_metadata["keywords"],
                       "Description": core_xml_metadata["description"],
                       "Application": app_xml_metadata["application"],
                       "App Version": app_xml_metadata["appVersion"],
                       "Template": app_xml_metadata["template"],
                       "Doc Security": app_xml_metadata["docSecurity"],
                       "Category": core_xml_metadata["category"],
                       "Content Status": core_xml_metadata["contentStatus"]
                       }

        # Writing document summary worksheet.
        headers = ["File Name", "Unique rsidR", "RSID Root", "<w:p> tags", "<w:r> tags", "<w:t> tags"]
        rows = [[filename, len(rsids), rsidRoot, documentXMLTagSummary["paragraphs"],
                 documentXMLTagSummary["runs"], documentXMLTagSummary["text"]]]
        write_worksheet(excel_file_path, "Doc_Summary", headers, rows)  # "Doc_Summary" worksheet

        # Writing rsids from settings.xml to "rsids" worksheet
        headers = ["File Name", "rsid Type", "RSID", "Count in document.xml"]
        rows = []  # declare empty list
        for rsid in rsids:
            rows.append([filename, "rsidR", rsid, "pending function"])
        write_worksheet(excel_file_path, "RSIDs", headers, rows)  # "RSIDs" worksheet

        # Writing XML files to "XML Files" worksheet
        headers = ["File Name", "XML", "Size (bytes)", "MD5Hash"]
        rows = []  # declare empty list
        for xml in XMLFiles:
            xml.insert(0, filename)
            rows.append(xml)
        write_worksheet(excel_file_path, "XML Files", headers, rows)  # "XML Files" worksheet

        # Writing metadata "metadata" worksheet
        headers = (list(allMetadata.keys()))
        rows = [list(allMetadata.values())]
        write_worksheet(excel_file_path, "metadata", headers, rows)  # "metadata" worksheet
