####################################
# Written by Jacques Boucher
# jjrboucher@gmail.com
#
# Version Date: 16 September 2023
#
# Written in Python 3.11
#
########## Description ##########
#
# Script will open a windows dialog to allow you to select a DOCx file.
# The script does not attempt to validate the file.
# A docx file is nothing more than a ZIP file, hence why this script uses the zipfile library.
#
# It will extract the results to a file called docx-artifacts.xlsx as defined by the variable excel_file_path at the start of the main part of the script.
# If the file does not exist, it creates it. If the file does exist, it appends to it.
# The file will be located in the folder where the script is executed from.
# If executing from the GUI by double clicking on the .py file, it should be stored in that same folder.
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
# 2 - It will extract all the unique RSIDs from the file word/settings.xml and write it to a worksheet called rsids_summary.
#     In this worksheet, it will save the following information to a row:
#     "File Name", "Unique RSID", "RSID Root"
#     Where "Unique RSID" is a numerical count of the # of RSIDs in the file.
#
#     What is an RSID (Revision Save ID)? See https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.rsid?view=openxml-2.8.1
#
# 3 - It will extract all the unique RSIDs from the file word/settings.xml and write it to a worksheet called rsids.
#     In this worksheet, it will save the following information to rows (one for each unique RSID):
#     "File Name", "RSID"
#
# 3 - It will extract all known relevant metadata from the files docProps/app.xml and docProps/core.xml and write it to a worksheet called metadata.
#     In this worksheet, it will save the following information to a row:
#     "File Name", "Author", "Created Date","Last Modified By","Modified Date","Last Printed Date","Manager","Company","Revision","Total Editing Time","Pages","Paragraphs","Lines","Words","Characters","Characters With Spaces","Title","Subject","Keywords","Description","Application","App Version","Template","Doc Security","Category","contentStatus"
#
#
########## Dependencies ##########
#
# If running the script on a Linux system, you may need to install python-tk. You can do this with the following command on a Debian (e.g. Ubuntu) system from the temrinal window:
# sudo apt-get install python3-tk
#
# Whether running on Linux, Mac, or Windows, you may need to install some of the libraries if they are not included in your installation of Python 3.
# In particular, you may need to install openpyxl and hashlib. You can do so as follows from a terminal window:
#
# pip3 install openpyxl
# pip3 install hashlib
#
# If any other libraries are missing when trying to execute the script, install those in the same manner.
#
###################################

import os
import zipfile
import tkinter as tk
import hashlib
from tkinter import filedialog
import re
from openpyxl import load_workbook
from openpyxl import Workbook


def extract_rsids_from_xml(xml_content):
    try:
        rsids = []
        pattern = r'<w:rsid(?:[^>]*)/>'
        matches = re.findall(pattern, xml_content)  # Find all RSIDs

        for match in matches:
            rsid_match = re.search(r'<w:rsid w:val="([^"]*)"', match)  # Loops through them
            if rsid_match:
                rsid = rsid_match.group(1)
                rsids.append(rsid)  # Appends it to the list
        try:
            rsidRoot = re.search(r'<w:rsidRoot w:val="([^"]*)"', xml_content).group(1)
        except:
            rsidRoot = ""
        return rsids, rsidRoot

    except Exception as e:
        print(f"An error occurred while extracting RSIDs: {e}")
        return []  # if it can't find any RSID (that should never happen), it returns an empty list.


def extract_from_app_xml(xml_content):
    # extract relevant metadata from app.xml file using a GREP expression
    appXML = {}
    appXML["template"] = re.search(r'<Template>(.*?)</Template>', xml_content)
    appXML["totalTime"] = re.search(r'<TotalTime>(.*?)</TotalTime>', xml_content)
    appXML["pages"] = re.search(r'<Pages>(.*?)</Pages>', xml_content)
    appXML["words"] = re.search(r'<Words>(.*?)</Words>', xml_content)
    appXML["characters"] = re.search(r'<Characters>(.*?)</Characters>', xml_content)
    appXML["application"] = re.search(r'<Application>(.*?)</Application>', xml_content)
    appXML["docSecurity"] = re.search(r'<DocSecurity>(.*?)</DocSecurity>', xml_content)
    appXML["lines"] = re.search(r'<Lines>(.*?)</Lines>', xml_content)
    appXML["paragraphs"] = re.search(r'<Paragraphs>(.*?)</Paragraphs>', xml_content)
    appXML["charactersWithSpaces"] = re.search(r'<CharactersWithSpaces>(.*?)</CharactersWithSpaces>', xml_content)
    appXML["appVersion"] = re.search(r'<AppVersion>(.*?)</AppVersion>', xml_content)
    appXML["manager"] = re.search(r'<Manager>(.*?)</Manager>', xml_content)
    appXML["company"] = re.search(r'<Company>(.*?)</Company>', xml_content)

    for key, value in appXML.items():  # check the resutls of the GREP searches
        if value == None:  # if no hit, assign empty value
            appXML[key] = ""
        else:  # if a hit, extract group(1) from the search hit
            appXML[key] = appXML[key].group(1)

    return appXML


def extract_from_core_xml(xml_content):
    # extract relevant metadata from core.xml file using a GREP expression
    coreXML = {}
    coreXML["title"] = re.search(r'<dc:title>(.*?)</dc:title>', xml_content)
    coreXML["subject"] = re.search(r'<dc:subject>(.*?)</dc:subject>', xml_content)
    coreXML["creator"] = re.search(r'<dc:creator>(.*?)</dc:creator>', xml_content)
    coreXML["keywords"] = re.search(r'<cp:keywords>(.*?)</cp:keywords>', xml_content)
    coreXML["description"] = re.search(r'<dc:description>(.*?)</dc:description>', xml_content)
    coreXML["revision"] = re.search(r'<cp:revision>(.*?)</cp:revision>', xml_content)
    coreXML["created"] = re.search(r'<dcterms:created.*?>(.*?)</dcterms:created>', xml_content)
    coreXML["modified"] = re.search(r'<dcterms:modified.*?>(.*?)</dcterms:modified>', xml_content)
    coreXML["lastModifiedBy"] = re.search(r'<cp:lastModifiedBy>(.*?)</cp:lastModifiedBy>', xml_content)
    coreXML["lastPrinted"] = re.search(r'<cp:lastPrinted>(.*?)</cp:lastPrinted>', xml_content)
    coreXML["category"] = re.search(r'<cp:category>(.*?)</cp:category>', xml_content)
    coreXML["contentStatus"] = re.search(r'<cp:contentStatus>(.*?)</cp:contentStatus>', xml_content)

    for key, value in coreXML.items():  # check the resutls of the GREP searches
        if value == None:  # if no hit, assign empty value
            coreXML[key] = ""
        else:  # if a hit, extract group(1) from the search hit
            coreXML[key] = coreXML[key].group(1)
    return coreXML


def list_of_XML_files(filename_path, file_name):
    with zipfile.ZipFile(filename_path, 'r') as zip_file:
        # list content of the DOCx file
        XMLFiles = []
        for file_info in zip_file.infolist():
            with zipfile.ZipFile(filename_path, 'r') as zip_ref:
                with zip_ref.open(file_info.filename) as xml_file:
                    md5hash = hashlib.md5(xml_file.read()).hexdigest()
            XMLFiles.append([file_name, file_info.filename, file_info.file_size, md5hash])
        return (XMLFiles)


def write_to_excel(xml_files, excel_file_path, file_name, rsids, rsidRoot, creator, created, lastModifiedBy, modified,
                   lastPrinted, manager, company, revision, totalTime, pages, paragraphs, lines, words, characters,
                   charactersWithSpaces, title, subject, keywords, description, application, appVersion, template,
                   docSecurity, category, contentStatus):
    try:
        if os.path.exists(excel_file_path):  # if the file exists, open it.
            workbook = load_workbook(excel_file_path)
        else:  # otherwise, create it
            workbook = Workbook()

        # List of files in DOCx document
        if "XML_files" in workbook.sheetnames:  # if the worksheet XML_files already exists, select it.
            worksheet = workbook["XML_files"]
        else:
            # Create the worksheet "XML_files"
            worksheet = workbook.create_sheet(title="XML_files")
            worksheet.append(["File Name", "XML", "Size (bytes)", "MD5Hash"])

        for fn, xml, size, md5hash in xml_files:  # Loop through all the embedded files
            worksheet.append([fn, xml, size, md5hash])  # Write a row to the spreadsheet for each embedded file.

        print(f"List of XML files along with size and hash appended to worksheet 'XML_files'")

        # Summary worksheet of # of RSIDs in a document
        if "rsids_summary" in workbook.sheetnames:  # if the worksheet rsids_summary already exits, select it.
            worksheet = workbook["rsids_summary"]
        else:
            # Create the worksheet "rsids"
            worksheet = workbook.create_sheet(title="rsids_summary")
            worksheet.append(["File Name", "Unique RSID", "RSID Root"])

        worksheet.append([file_name, len(rsids), rsidRoot])

        print(f"RSIDs summary appended to worksheet 'rsids_summary'")

        # Check if the worksheet "rsids" already exists
        if "rsids" in workbook.sheetnames:  # if the worksheet rsids already exists, select it.
            worksheet = workbook["rsids"]
        else:
            # Create the worksheet "rsids"
            worksheet = workbook.create_sheet(title="rsids")
            worksheet.append(["File Name", "RSID"])

        for rsid in set(rsids):
            worksheet.append([file_name, rsid])

        print(f"Unique RSIDs appended to '{excel_file_path}' in worksheet 'rsids'")

        # Check if the worksheet "metadata" already exists
        if "metadata" in workbook.sheetnames:  # if the worksheet metadata already exists, select it.
            worksheet = workbook["metadata"]
        else:
            # Create the worksheet "metadata"
            worksheet = workbook.create_sheet(title="metadata")
            worksheet.append(
                ["File Name", "Author", "Created Date", "Last Modified By", "Modified Date", "Last Printed Date",
                 "Manager", "Company", "Revision", "Total Editing Time", "Pages",
                 "Paragraphs", "Lines", "Words", "Characters", "Characters With Spaces", "Title", "Subject", "Keywords",
                 "Description", "Application", "App Version", "Template",
                 "Doc Security", "Category", "Content Status"])

        worksheet.append(
            [file_name, creator, created, lastModifiedBy, modified, lastPrinted, manager, company, revision, totalTime,
             pages, paragraphs, lines, words, characters,
             charactersWithSpaces, title, subject, keywords, description, application, appVersion, template,
             docSecurity, category, contentStatus])

        print(f"Metadata appeneded to '{excel_file_path}' in worksheet 'metadata'")

        # Remove the default sheet created by openpyxl
        default_sheet = workbook.active
        if default_sheet.title == "Sheet":
            workbook.remove(default_sheet)

        workbook.save(excel_file_path)  # save the file

    except Exception as e:
        print(f"An error occurred while writing to Excel: {e}")


if __name__ == "__main__":

    # Output file - same path as where the script is run. It will create it if it does not exist, or append to it if it does.
    excel_file_path = "docx-artifacts.xlsx"

    root = tk.Tk()
    root.withdraw()  # Hide the main window

    zip_file_path = filedialog.askopenfilename(title="Select DOCx File", filetypes=[("DOCx Files", "*.docx")])
    if not zip_file_path:
        print("No DOCx file selected. Exiting.")
    else:

        filename = os.path.basename(zip_file_path)

        # list of XML file in DOCx
        XMLFiles = list_of_XML_files(zip_file_path,
                                     filename)  # Executes the function to get a list of all XML files within the DOCx file

        # parse word/settings.xml artifacts
        xml_file_path_within_zip = "word/settings.xml"  # Path of the XML file within the ZIP

        try:
            with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                with zip_ref.open(xml_file_path_within_zip) as xml_file:
                    xml_content = xml_file.read().decode("utf-8")

                    rsids, rsidRoot = extract_rsids_from_xml(
                        xml_content)  # Executes the function to get all unique RSIDs.

        except FileNotFoundError:
            print(f"File '{xml_file_path_within_zip}' not found in the ZIP archive.")
        except Exception as e:
            print(f"An error occurred: {e}")

        # parse docProps/app.xml artifacts
        xml_file_path_within_zip = "docProps/app.xml"  # Path of the XML file within the ZIP

        try:
            with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                with zip_ref.open(xml_file_path_within_zip) as xml_file:
                    xml_content = xml_file.read().decode("utf-8")
                    app_xml_metadata = extract_from_app_xml(
                        xml_content)  # Executes the function to get metadata from app.xml

        except FileNotFoundError:
            print(f"File '{xml_file_path_within_zip}' not found in the ZIP archive.")
        except Exception as e:
            print(f"An error occurred: {e}")

        # parse docProps/core.xml artifacts
        xml_file_path_within_zip = "docProps/core.xml"  # Path of the XML file within the ZIP
        try:
            with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                with zip_ref.open(xml_file_path_within_zip) as xml_file:
                    xml_content = xml_file.read().decode("utf-8")
                    core_xml_metadata = extract_from_core_xml(
                        xml_content)  # Executes the function to get the metadata from core.xml

        except FileNotFoundError:
            print(f"File '{xml_file_path_within_zip}' not found in the ZIP archive.")
        except Exception as e:
            print(f"An error occurred: {e}")

        write_to_excel(XMLFiles, excel_file_path, filename, rsids, rsidRoot, core_xml_metadata["creator"],
                       core_xml_metadata["created"], core_xml_metadata["lastModifiedBy"], core_xml_metadata["modified"],
                       core_xml_metadata["lastPrinted"], app_xml_metadata["manager"], app_xml_metadata["company"],
                       core_xml_metadata["revision"], app_xml_metadata["totalTime"], app_xml_metadata["pages"],
                       app_xml_metadata["paragraphs"], app_xml_metadata["lines"], app_xml_metadata["words"],
                       app_xml_metadata["characters"], app_xml_metadata["charactersWithSpaces"],
                       core_xml_metadata["title"], core_xml_metadata["subject"], core_xml_metadata["keywords"],
                       core_xml_metadata["description"], app_xml_metadata["application"],
                       app_xml_metadata["appVersion"],
                       app_xml_metadata["template"], app_xml_metadata["docSecurity"], core_xml_metadata["category"],
                       core_xml_metadata[
                           "contentStatus"])  # Executes the function to write everything to the Excel file.
