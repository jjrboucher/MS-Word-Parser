from openpyxl import load_workbook
from openpyxl import Workbook
import os


def write_to_excel(excel_filepath, file_name, xml_files, all_rsids, document_summary, rsid_root,
                   all_metadata):
    """
    This function writes the artifacts collected from the MS Word document to an Excel file.
    :param excel_filepath:
    :param file_name:
    :param xml_files:
    :param all_rsids:
    :param document_summary:
    :param rsid_root:
    :param all_metadata:
    :return: nil
    """
    try:
        if os.path.exists(excel_filepath):  # if the file exists, open it.
            workbook = load_workbook(excel_filepath)
        else:  # otherwise, create it
            workbook = Workbook()

        # List of files in DOCx document
        if "XML_files" in workbook.sheetnames:  # if the worksheet XML_files already exists, select it.
            worksheet = workbook["XML_files"]
        else:
            # Create the worksheet "XML_files"
            worksheet = workbook.create_sheet(title="XML_files")
            worksheet.append(["File Name", "XML", "Size (bytes)", "MD5Hash"])

        for msword_file, xml_file, file_size, md5hash in xml_files:  # Loop through all the embedded files
            # Write a row to the spreadsheet for each embedded file.
            worksheet.append([msword_file, xml_file, file_size, md5hash])

        print(f"List of XML files along with size and hash appended to worksheet 'XML_files'")

        # Summary worksheet of # of RSIDs in a document
        if "doc_summary" in workbook.sheetnames:  # if the worksheet doc_summary already exits, select it.
            worksheet = workbook["doc_summary"]
        else:
            # Create the worksheet "doc_summary"
            worksheet = workbook.create_sheet(title="doc_summary")
            worksheet.append(["File Name", "Unique RSIDs", "RSID Root", "<w:p> tags", "<w:r> tags", "<w:t> tags"])

        worksheet.append([file_name, len(all_rsids), rsid_root, document_summary["paragraphs"],
                          document_summary["runs"], document_summary["text"]])

        print(f"Document summary appended to worksheet 'doc_summary'")

        # Check if the worksheet "rsids" already exists
        if "rsids" in workbook.sheetnames:  # if the worksheet rsids already exists, select it.
            worksheet = workbook["rsids"]
        else:
            # Create the worksheet "rsids"
            worksheet = workbook.create_sheet(title="rsids")
            worksheet.append(["File Name", "RSID"])

        for rsid in set(all_rsids):
            worksheet.append([file_name, rsid])

        print(f"Unique RSIDs appended to worksheet 'rsids'")

        # Check if the worksheet "metadata" already exists
        if "metadata" in workbook.sheetnames:  # if the worksheet metadata already exists, select it.
            worksheet = workbook["metadata"]
        else:
            # Create the worksheet "metadata"
            worksheet = workbook.create_sheet(title="metadata")
            headings = list(all_metadata.keys())  # Adds the keys as column headings to a list
            headings.insert(0, "File Name")  # Adds column heading "File Name" at the start of the list
            worksheet.append(headings)  # Writes the headings to the spreadsheet

        metadata = list(all_metadata.values())  # Adds values to the list
        metadata.insert(0, file_name)  # Adds the file name to the start of the list
        worksheet.append(metadata)  # Writes the metadata to the spreadsheet

        print(f"Metadata appended to worksheet 'metadata'")

        # Remove the default sheet created by openpyxl
        default_sheet = workbook.active
        if default_sheet.title == "Sheet":
            workbook.remove(default_sheet)

        workbook.save(excel_filepath)  # save the file

        print(f"Results written to {excel_filepath}.")

    except Exception as function_error:
        print(f"An error occurred while writing to Excel: {function_error}")
