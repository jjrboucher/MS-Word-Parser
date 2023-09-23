import hashlib
import zipfile


def list_of_xml_files(filename_path, file_name):
    print("Processing word/document.xml for list of XML files.")
    with zipfile.ZipFile(filename_path, 'r') as zip_file:
        # list content of the DOCx file
        xml_files = []
        for file_info in zip_file.infolist():
            with zipfile.ZipFile(filename_path, 'r') as zip_ref:
                with zip_ref.open(file_info.filename) as xml_file:
                    md5hash = hashlib.md5(xml_file.read()).hexdigest()
            xml_files.append([file_name, file_info.filename, file_info.file_size, md5hash])
        return xml_files


def extract_content_of_xml(docxfile, xmlfile):

    try:
        with zipfile.ZipFile(docxfile, 'r') as mswordFile:
            with mswordFile.open(xmlfile) as xmlFile:
                return xmlFile.read().decode("utf-8")

    except FileNotFoundError:
        print(f"File '{xmlfile}' not found in the ZIP archive.")
        return ""
    except Exception as e:
        print(f"An error occurred: {e}")
        return ""
