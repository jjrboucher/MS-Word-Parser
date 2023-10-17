####################################
# Written by Jacques Boucher
# jjrboucher@gmail.com
#
# Version Date: 15 October 2023
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
#     "File Name", "Unique rsidR", "RSID Root", "<w:p> tags", "<w:r> tags", "<w:t> tags"
#     Where "Unique RSID" is a numerical count of the # of RSIDs in the file settings.xml.
#
#     What is an RSID (Revision Save ID)?
#     See https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.rsid?view=openxml-2.8.1
#
# 3 - It will extract all the unique RSIDs from the file word/settings.xml and write it to a worksheet called RSIDs,
#     along with a count of how many times that RSID is in document.xml
#     It will also search document.xml for all unique rsidRPr, rsidP, and rsidRDefault values and count of how many
#     are in document.xml.
#     In this worksheet, it will save the following information to rows (one for each unique RSID):
#     "File Name", "RSID Type", "RSID Value", "Count in document.xml"
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
import re
import tkinter as tk
from tkinter import filedialog
from functions.excel import write_worksheet  # function to write results to an Excel file
from classes.ms_word import Docx

red = f'\033[91m'
white = f'\033[00m'
green = f'\033[92m'


def process_docx(filename):
    """
    This function accepts a filename of type Docx and processes it.
    By placing this in a function, it allows the main part of the script to accept multiple file names and
    then loop through them, calling this function for each DOCx file.
    """

    global excel_file_path
    print(f'Updating {green}"Doc_Summary"{white} worksheet in {excel_file_path}')
    # Writing document summary worksheet.
    headers = ["File Name", "Unique rsidR", "RSID Root", "<w:p> tags", "<w:r> tags", "<w:t> tags"]
    rows = [[filename.filename(), len(filename.rsidr()), filename.rsid_root(), filename.paragraph_tags(),
             filename.runs_tags(), filename.text_tags()]]
    write_worksheet(excel_file_path, "Doc_Summary", headers, rows)  # "Doc_Summary" worksheet

    # The keys will be used as the column heading in the spreadsheet
    # The order they are in is the order that the columns will be in the spreadsheet
    # Corresponding values passed, resulting in a dictionary being passed called allMetadata
    # containing column headings and associated extracted metadata value.
    allmetadata = {"File Name": filename.filename(),
                   "Author": filename.creator(),
                   "Created Date": filename.created(),
                   "Last Modified By": filename.last_modified_by(),
                   "Modified Date": filename.modified(),
                   "Last Printed Date": filename.last_printed(),
                   "Manager": filename.manager(),
                   "Company": filename.company(),
                   "Revision": filename.revision(),
                   "Total Editing Time": filename.total_editing_time(),
                   "Pages": filename.pages(),
                   "Paragraphs": filename.paragraphs(),
                   "Lines": filename.lines(),
                   "Words": filename.words(),
                   "Characters": filename.characters(),
                   "Characters With Spaces": filename.characters_with_spaces(),
                   "Title": filename.title(),
                   "Subject": filename.subject(),
                   "Keywords": filename.keywords(),
                   "Description": filename.description(),
                   "Application": filename.application(),
                   "App Version": filename.app_version(),
                   "Template": filename.template(),
                   "Doc Security": filename.security(),
                   "Category": filename.category(),
                   "Content Status": filename.content_status()
                   }

    print(f'Updating {green}"Metadata"{white} worksheet in "{excel_file_path}"')
    # Writing metadata "metadata" worksheet
    headers = (list(allmetadata.keys()))
    rows = [list(allmetadata.values())]
    write_worksheet(excel_file_path, "Metadata", headers, rows)  # "metadata" worksheet

    print(f'Updating {green}"Archive Files"{white} worksheet in "{excel_file_path}"')
    # Writing XML files to "XML Files" worksheet
    headers = ["File Name", "Archive File", "Size (bytes)", "MD5Hash"]
    rows = []  # declare empty list

    for xml, size_hash in filename.xml_files().items():
        rows.append([filename.filename(), xml, size_hash[0], size_hash[1]])  # add the row to the list "rows"
    write_worksheet(excel_file_path, "Archive Files", headers, rows)  # "XML Files" worksheet

    # Calculating count of rsidR, rsidRPr, rsidP, and rsidRDefault in document.xml and writing to "rsids" worksheet
    headers = ["File Name", "RSID Type", "RSID Value", "Count in document.xml"]
    rows = []  # declare empty list

    print(f'Adding {green}rsidR{white} count to "RSIDs" worksheet in "{excel_file_path}"')
    for k, v in filename.rsidr_in_document_xml().items():
        rows.append([filename.filename(), "rsidR", k, v])
    write_worksheet(excel_file_path, "RSIDs", headers, rows)  # "RSIDs" worksheet

    rows = []  # declare empty list

    print(f'Adding {green}rsidP{white} count to "RSIDs" worksheet in {excel_file_path}')
    for k, v in filename.rsidp_in_document_xml().items():
        rows.append([filename.filename(), "rsidP", k, v])
    write_worksheet(excel_file_path, "RSIDs", headers, rows)  # "RSIDs" worksheet

    rows = []  # declare empty list

    print(f'Adding {green}rsidPr{white} count to "RSIDs" worksheet in {excel_file_path}')
    for k, v in filename.rsidrpr_in_document_xml().items():
        rows.append([filename.filename(), "rsidRPr", k, v])
    write_worksheet(excel_file_path, "RSIDs", headers, rows)  # "RSIDs" worksheet

    rows = []  # declare empty list

    print(f'Adding {green}rsidRDefault{white} count to "RSIDs" worksheet in {excel_file_path}')
    for k, v in filename.rsidrdefault_in_document_xml().items():
        rows.append([filename.filename(), "rsidRDefault", k, v])
    write_worksheet(excel_file_path, "RSIDs", headers, rows)  # "RSIDs" worksheet

    return


if __name__ == "__main__":

    # Output file - same path as where the script is run. It will create it if it does not exist,
    # or append to it if it does.
    # excel_file_path = "docx-artifacts(class).xlsx"  # default file name - will be created in the script folder.

    root = tk.Tk()
    root.withdraw()  # Hide the main window

    msword_file_path = filedialog.askopenfilenames(title="Select DOCx file(s) to process", initialdir=".",
                                                   filetypes=[("DOCx Files", "*.docx")])  # ask for file(s) to process

    if not msword_file_path:  # if no docx file name was selected to process
        print(f'{red}No DOCx file selected.{white} Exiting.')
    else:
        excel_file_path = filedialog.asksaveasfilename(title="Select new or existing XLSX file for output.",
                                                       initialdir=".", filetypes=[("Excel Files", "*.xlsx")],
                                                       defaultextension="*.xlsx",
                                                       confirmoverwrite=False)  # ask for output file

        if not excel_file_path:  # if no output file selected
            print(f'{red}No output file selected.{white} Exiting.')
            exit()

        if not re.search(r'\.xlsx$', excel_file_path):  # if .xlsx was not included in file name, add it.
            excel_file_path += ".xlsx"

        for f in msword_file_path:  # loop over the files selected, processing each.
            print(f'Processing {green}"{f}"{white}')
            process_docx(Docx(f))
            print(f'Finished processing {green}"{f}"{white}. Results are found in '
                  f'{green}"{excel_file_path}"{white}')
