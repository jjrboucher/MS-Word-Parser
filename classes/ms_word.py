import zipfile
import hashlib
import re


class Docx:

    def __init__(self, msword_file):
        """.docx file to pass to the class"""
        self.msword_file = msword_file
        self.core_xml_file = "docProps/core.xml"
        self.core_xml_content = self.load_core_xml()
        self.app_xml_file = "docProps/app.xml"
        self.app_xml_content = self.load_app_xml()
        self.document_xml_file = "word/document.xml"
        self.document_xml_content = self.load_document_xml()

    def load_core_xml(self):
        # load core.xml
        try:
            with zipfile.ZipFile(self.msword_file, 'r') as zipref:
                with zipref.open(self.core_xml_file) as xmlFile:
                    return xmlFile.read().decode("utf-8")
        except FileNotFoundError:
            print(f"File '{self.core_xml_file} not found in the DOCx archive.")
        except Exception as e:
            print(f"An error occurred: {e}")

    def load_app_xml(self):
        # load app.xml
        try:
            with zipfile.ZipFile(self.msword_file, 'r') as zipref:
                with zipref.open(self.app_xml_file) as xmlFile:
                    return xmlFile.read().decode("utf-8")
        except FileNotFoundError:
            print(f"File '{self.app_xml_file} not found in the DOCx archive.")
        except Exception as e:
            print(f"An error occurred: {e}")

        self.load_document_xml()

    def load_document_xml(self):
        # load document.xml
        try:
            with zipfile.ZipFile(self.msword_file, 'r') as zipref:
                with zipref.open(self.document_xml_file) as xmlFile:
                    return xmlFile.read().decode("utf-8")
        except FileNotFoundError:
            print(f"File '{self.document_xml_file}' not found in the ZIP archive.")
        except Exception as e:
            print(f"An error occurred: {e}")

    def filename(self):
        return self.msword_file

    def xml_files(self):
        with zipfile.ZipFile(self.msword_file, 'r') as zip_file:
            # list content of the DOCx file
            xml_files = {}
            for file_info in zip_file.infolist():
                with zipfile.ZipFile(self.msword_file, 'r') as zip_ref:
                    with zip_ref.open(file_info.filename) as xml_file:
                        md5hash = hashlib.md5(xml_file.read()).hexdigest()
                xml_files[file_info.filename] = [file_info.file_size, md5hash]
            return xml_files  # returns dictionary {xml_filename: [file size, file hash]}

    def xml_hash(self, xmlfile):
        return self.xml_files()[xmlfile][1]

    def xml_size(self, xmlfile):
        return self.xml_files()[xmlfile][0]

    def title(self):
        doc_title = re.search(r'<dc:title>(.*?)</dc:title>', self.core_xml_content)
        return "" if doc_title is None else doc_title.group(1)

    def subject(self):
        doc_subject = re.search(r'<dc:subject>(.*?)</dc:subject>', self.core_xml_content)
        return "" if doc_subject is None else doc_subject.group(1)

    def creator(self):
        doc_creator = re.search(r'<dc:creator>(.*?)</dc:creator>', self.core_xml_content)
        return "" if doc_creator is None else doc_creator.group(1)

    def keywords(self):
        doc_keywords = re.search(r'<cp:keywords>(.*?)</cp:keywords>', self.core_xml_content)
        return "" if doc_keywords is None else doc_keywords.group(1)

    def description(self):
        doc_description = re.search(r'<dc:description>(.*?)</dc:description>', self.core_xml_content)
        return "" if doc_description is None else doc_description.group(1)

    def revision(self):
        doc_revision = re.search(r'<cp:revision>(.*?)</cp:revision>', self.core_xml_content)
        return "" if doc_revision is None else doc_revision.group(1)

    def created(self):
        doc_created = re.search(r'<dcterms:created[^>].*?>(.*?)</dcterms:created>', self.core_xml_content)
        return "" if doc_created is None else doc_created.group(1)

    def modified(self):
        doc_modified = re.search(r'<dcterms:modified[^>].*?>(.*?)</dcterms:modified>', self.core_xml_content)
        return "" if doc_modified is None else doc_modified.group(1)

    def last_modified_by(self):
        doc_lastmodifiedby = re.search(r'<cp:lastModifiedBy>(.*?)</cp:lastModifiedBy>', self.core_xml_content)
        return "" if doc_lastmodifiedby is None else doc_lastmodifiedby.group(1)

    def last_printed(self):
        doc_lastprinted = re.search(r'<cp:lastPrinted>(.*?)</cp:lastPrinted>', self.core_xml_content)
        return "" if doc_lastprinted is None else doc_lastprinted.group(1)

    def category(self):
        doc_category = re.search(r'<cp:category>(.*?)</cp:category>', self.core_xml_content)
        return "" if doc_category is None else doc_category.group(1)

    def content_status(self):
        doc_contentstatus = re.search(r'<cp:contentStatus>(.*?)</cp:contentStatus>', self.core_xml_content)
        return "" if doc_contentstatus is None else doc_contentstatus.group(1)

    def template(self):
        doc_template = re.search(r'<Template>(.*?)</Template>', self.app_xml_content)
        return "" if doc_template is None else doc_template.group(1)

    def total_editing_time(self):
        doc_edit_time = re.search(r'<TotalTime>(.*?)</TotalTime>', self.app_xml_content)
        return "" if doc_edit_time is None else doc_edit_time.group(1)

    def pages(self):
        doc_pages = re.search(r'<Pages>(.*?)</Pages>', self.app_xml_content)
        return "" if doc_pages is None else doc_pages.group(1)

    def words(self):
        doc_words = re.search(r'<Words>(.*?)</Words>', self.app_xml_content)
        return "" if doc_words is None else doc_words.group(1)

    def characters(self):
        doc_characters = re.search(r'<Characters>(.*?)</Characters>', self.app_xml_content)
        return "" if doc_characters is None else doc_characters.group(1)

    def application(self):
        doc_application = re.search(r'<Application>(.*?)</Application>', self.app_xml_content)
        return "" if doc_application is None else doc_application.group(1)

    def security(self):
        doc_security = re.search(r'<DocSecurity>(.*?)</DocSecurity>', self.app_xml_content)
        return "" if doc_security is None else doc_security.group(1)

    def lines(self):
        doc_lines = re.search(r'<Lines>(.*?)</Lines>', self.app_xml_content)
        return "" if doc_lines is None else doc_lines.group(1)

    def paragraphs(self):
        doc_paragraphs = re.search(r'<Paragraphs>(.*?)</Paragraphs>', self.app_xml_content)
        return "" if doc_paragraphs is None else doc_paragraphs.group(1)

    def characters_with_spaces(self):
        doc_characters_with_spaces = re.search(r'<CharactersWithSpaces>(.*?)</CharactersWithSpaces>',
                                               self.app_xml_content)
        return "" if doc_characters_with_spaces is None else doc_characters_with_spaces.group(1)

    def app_version(self):
        doc_app_version = re.search(r'<AppVersion>(.*?)</AppVersion>', self.app_xml_content)
        return "" if doc_app_version is None else doc_app_version.group(1)

    def manager(self):
        doc_manager = re.search(r'<Manager>(.*?)</Manager>', self.app_xml_content)
        return "" if doc_manager is None else doc_manager.group(1)

    def company(self):
        doc_company = re.search(r'<Company>(.*?)</Company>', self.app_xml_content)
        return "" if doc_company is None else doc_company.group(1)

    def paragraph_tags(self):
        return len(re.findall(r'</w:p>', self.document_xml_content))

    def runs_tags(self):
        return len(re.findall(r'</w:r>', self.document_xml_content))

    def text_tags(self):
        return len(re.findall(r'</w:t>', self.document_xml_content))

    def __str__(self):
        if self.last_printed() == "":
            printed = f'Document was never printed'
        else:
            printed = f'Printed: {self.last_printed()}'
        return (f'Document: {self.filename()}\n'
                f'Created by: {self.creator()}\n'
                f'Created date: {self.created()}\n'
                f'Last edited by: {self.last_modified_by()}\n'
                f'Edited date: {self.modified()}\n'
                f'{printed}\n'
                f'Total pages: {self.pages()}\n'
                f'Total editing time: {self.total_editing_time()} minute(s).')
