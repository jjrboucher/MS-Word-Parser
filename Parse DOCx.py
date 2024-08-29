####################################
# Written by Jacques Boucher
# jjrboucher@gmail.com
# Version Date: 29 August 2024
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
#     It also extracts the unique paraId and textId tags from the <w:p> tag and saves the values and count of how
#     many are in document.xml.
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
# ********** Possible future enhancements **********
#
# Option to not hash files (at least in triage mode). Processing a large # of files is time-consuming when needing to
# hash each file. That may not be needed in some cases. Removing the hashing from the summary worksheet would
# significantly increase the speed of execution.
#
###################################

from classes.ms_word import Docx
from functions.ms_word_menu import docx_menu
from colorama import just_fix_windows_console
import pandas as pd
import re
from sys import exit
import time

red = f'\033[91m'
white = f'\033[00m'
green = f'\033[92m'
triage = False
just_fix_windows_console()
docxErrorCount = 0  # tracks how many files it could not process.
filesUnableToProcess = []  # list of files that produced an error
doc_summary_worksheet = {}  # contains summary data parsed from each file processed
metadata_worksheet = {}  # contains the metadata parsed from each file processed
archive_files_worksheet = {}  # contains the archive files data from each file processed
rsids_worksheet = {}  # contains the RSID artifacts extracted from each file processed
process_or_cancel = ""  # variable to capture whether the user clicked to process, or cancel
logFile = ""
errorLog = ""
excel_file_path = ""


def process_docx(filename):
    """
    This function accepts a filename of type Docx and processes it.
    By placing this in a function, it allows the main part of the script to accept multiple file names and
    then loop through them, calling this function for each DOCx file.
    """

    global excel_file_path, triage, doc_summary_worksheet, metadata_worksheet, archive_files_worksheet, rsids_worksheet

    write_log(f'{filename.__str__()}\n')

    for checkFile in ("word/settings.xml", "docProps/core.xml", "docProps/app.xml"):  # checks if xml files being parsed
        # are present and notes same in the log file.
        xml_exists = checkFile in filename.xml_files().keys()
        write_log(f'**{checkFile} exists? {xml_exists}\n')

    # Writing document summary worksheet.

    headers = ["File Name", "MD5 Hash", "Unique rsidR", "RSID Root", "<w:p> tags", "<w:r> tags", "<w:t> tags"]

    if not bool(doc_summary_worksheet):  # if it's an empty dictionary, add headers to it.
        doc_summary_worksheet = dict((k, []) for k in headers)

    doc_summary_worksheet[headers[0]].append(filename.filename())
    doc_summary_worksheet[headers[1]].append(filename.hash())
    doc_summary_worksheet[headers[2]].append(len(filename.rsidr()))
    doc_summary_worksheet[headers[3]].append(filename.rsid_root())
    doc_summary_worksheet[headers[4]].append(filename.paragraph_tags())
    doc_summary_worksheet[headers[5]].append(filename.runs_tags())
    doc_summary_worksheet[headers[6]].append(filename.text_tags())

    print(f'Extracted {green}Doc_Summary{white} artifacts')

    # The keys will be used as the column heading in the spreadsheet
    # The order they are in is the order that the columns will be in the spreadsheet
    # Corresponding values passed, resulting in a dictionary being passed called allMetadata
    # containing column headings and associated extracted metadata value.

    headers = ["File Name", "Author", "Created Date", "Last Modified By", "Modified Date", "Last Printed Date",
               "Manager", "Company", "Revision", "Total Editing Time", "Pages", "Paragraphs", "Lines", "Words",
               "Characters", "Characters With Spaces", "Title", "Subject", "Keywords", "Description",
               "Application", "App Version", "Template", "Doc Security", "Category", "Content Status"]

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

    print(f'Extracted {green}metadata{white} artifacts')

    if not triage:  # will generate these spreadsheet if not triage
        print(f'Updating {green}"Archive Files"{white} worksheet in "{excel_file_path}"')
        # Writing XML files to "Archive Files" worksheet
        headers = ["File Name",
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
                   "ZIP Extra Characters (truncated)"
                   ]

        if not bool(archive_files_worksheet):  # if it's an empty dictionary, add headers to it.
            archive_files_worksheet = dict((k, []) for k in headers)

        for xml, xml_info in filename.xml_files().items():
            extra_characters = xml_info[9] if xml_info[8] == 0 else ",".join(xml_info[9])  # If no extra characters,
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

        print(f'Extracted {green}archive files{white} artifacts')

        # Calculating count of rsidR, rsidRPr, rsidP, rsidRDefault, paraId, and textId in document.xml
        # and writing to "rsids" worksheet
        headers = ["File Name", "RSID Type", "RSID Value", "Count in document.xml"]

        if not bool(rsids_worksheet):  # if it's an empty dictionary, add headers to it.
            rsids_worksheet = dict((k, []) for k in headers)

        print(f'Calculating {green}rsidR{white} count')
        for k, v in filename.rsidr_in_document_xml().items():
            rsids_worksheet[headers[0]].append(filename.filename())
            rsids_worksheet[headers[1]].append('rsidR')
            rsids_worksheet[headers[2]].append(k)
            rsids_worksheet[headers[3]].append(v)

        print(f'Calculating {green}rsidP{white} count')
        for k, v in filename.rsidp_in_document_xml().items():
            rsids_worksheet[headers[0]].append(filename.filename())
            rsids_worksheet[headers[1]].append('rsidP')
            rsids_worksheet[headers[2]].append(k)
            rsids_worksheet[headers[3]].append(v)

        print(f'Calculating {green}rsidPr{white} count')
        for k, v in filename.rsidrpr_in_document_xml().items():
            rsids_worksheet[headers[0]].append(filename.filename())
            rsids_worksheet[headers[1]].append('rsidRPr')
            rsids_worksheet[headers[2]].append(k)
            rsids_worksheet[headers[3]].append(v)

        print(f'Calculating {green}rsidRDefault{white} count')
        for k, v in filename.rsidrdefault_in_document_xml().items():
            rsids_worksheet[headers[0]].append(filename.filename())
            rsids_worksheet[headers[1]].append('rsidRDefault')
            rsids_worksheet[headers[2]].append(k)
            rsids_worksheet[headers[3]].append(v)

        print(f'Calculating {green}paraID{white} count')
        for k, v in filename.paragraph_id_tags().items():
            rsids_worksheet[headers[0]].append(filename.filename())
            rsids_worksheet[headers[1]].append('paraID')
            rsids_worksheet[headers[2]].append(k)
            rsids_worksheet[headers[3]].append(v)

        print(f'Calculating {green}textID{white} count')
        for k, v in filename.text_id_tags().items():
            rsids_worksheet[headers[0]].append(filename.filename())
            rsids_worksheet[headers[1]].append('textID')
            rsids_worksheet[headers[2]].append(k)
            rsids_worksheet[headers[3]].append(v)

    write_log(f'\n------------------------------------\n')
    return


def write_log(text):
    """
    Write to log file
    """
    global logFile
    #  Open file to write
    lf = open(logFile, "a", encoding='utf8')
    #  Write text to it
    lf.write(text)
    #  Close file.
    lf.close()


def write_error_log(text):
    """
    Write to the error log file
    """
    global errorLog
    #  Open file to write
    lf = open(errorLog, "a", encoding='utf8')
    #  Write text to it
    lf.write(text)
    #  Close file.
    lf.close()


if __name__ == "__main__":

    process_or_cancel, logFile, errorLog, processingOption, hashFiles, excel_file_path, msword_file_path = docx_menu()

    if process_or_cancel == "CANCEL":
        print(f'You clicked on {red}CANCEL{white}.')
        input(f'Press {green}ENTER{white} to exit script.')
        exit()
    elif process_or_cancel == "":
        print(f'You clicked on the {red}X{white} and {red}closed{white} the window.')
        input(f'Press {green}ENTER{white} to exit script.')
        exit()

    if processingOption == "triage":
        triage = True

    docxPath = msword_file_path[0][0:msword_file_path[0].rindex("/") + 1]  # extract path of DOCx file(s) to process

    logFilesPath = (excel_file_path[0:excel_file_path.rindex("/") + 1])
    logFile = (logFilesPath + logFile)

    errorLog = (logFilesPath + errorLog)

    write_log("Script executed: " + time.strftime("%Y-%m-%d_%H:%M:%S") + '\n')

    write_log(f'Excel output file: {excel_file_path}\n')
    write_log(f'\nSummary of files parsed:\n========================\n')

    if not re.search(r'\.xlsx$', excel_file_path):  # if .xlsx was not included in file name, add it.
        excel_file_path += ".xlsx"

    for f in msword_file_path:  # loop over the files selected, processing each.
        print(f'\nProcessing {green}"{f}"{white}')
        try:
            process_docx(Docx(f, triage, hashFiles))

        except Exception as docxError:  # If processing a DOCx file raises an error, let the user know, and write it
            # to the error log.
            docxErrorCount += 1  # increment error count by 1.
            filesUnableToProcess.append(f)
            print(f'{red}error processing {f}. {white}Skipping.')
            write_error_log(f'Error trying to process {f}. Skipping.\n'
                            f'Error: {docxError}\n')
        print(f'Finished processing {green}"{f}"{white}. ')

    df = pd.DataFrame(data=doc_summary_worksheet)

    df.to_excel(excel_writer=excel_file_path, sheet_name="Doc_Summary", index=False)

    write_log(f'"Doc_Summary" worksheet written to Excel file.\n\n')

    df = pd.DataFrame(data=metadata_worksheet)

    with pd.ExcelWriter(path=excel_file_path, engine='openpyxl', mode='a') as writer:
        df.to_excel(excel_writer=writer, sheet_name="metadata", index=False)

    write_log(f'"Metadata" worksheet written to Excel.\n\n')

    if not triage:
        df = pd.DataFrame(data=archive_files_worksheet)

        with pd.ExcelWriter(path=excel_file_path, engine='openpyxl', mode='a') as writer:
            df.to_excel(excel_writer=writer, sheet_name="Archive Files", index=False)

        write_log(f'"Archive Files" worksheet written to Excel.\n\n')

        df = pd.DataFrame(data=rsids_worksheet)

        with pd.ExcelWriter(path=excel_file_path, engine='openpyxl', mode='a') as writer:
            df.to_excel(excel_writer=writer, sheet_name="RSIDs", index=False)

        write_log(f'"RSIDs" worksheet written to Excel.\n\n')

    print(f'\n==============================================\n'
          f'Excel output: {green}"{excel_file_path}"{white}\n'
          f'Log file: {green}"{logFile}"{white}')

    write_log("Script finished execution: " + time.strftime("%Y-%m-%d_%H:%M:%S") + '\n')

    if docxErrorCount:  # count greater than 0, meaning there are errors
        print(f'Error log file: {red}"{errorLog}"{white}\n==============================================\n')
        print(f'A total of {red}{docxErrorCount} files{white} could not be processed.')
        input(f'Press {green}Enter{white} to see a list of the files that could not be processed.')
        print(f'File(s) that {red}could not be processed{white}:\n')

        for file in filesUnableToProcess:
            print(f'{red}{file}{white}')
        input(f'Press {green}Enter{white} to exit application.')
