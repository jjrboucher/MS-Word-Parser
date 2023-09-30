import zipfile
import hashlib
import re


class Docx:
    """
    Accepts a docx file. Has the following methods to extract data from core.xml, app.xml, document.xml

    app_version, application, category, characters, characters_with_spaces, company, content_status, created, creator,
    description, filename, keywords, last_modified_by, last_printed, lines, manager, modified, pages, paragraph_tags,
    paragraphs, revision, runs_tags, security, subject, template, text_tags, title, total_editing_time, words,
    xml_files, xml_hash, xml_size
    """

    def __init__(self, msword_file):
        """.docx file to pass to the class"""
        self.msword_file = msword_file
        self.core_xml_file = "docProps/core.xml"
        self.core_xml_content = self.__load_core_xml()
        self.app_xml_file = "docProps/app.xml"
        self.app_xml_content = self.__load_app_xml()
        self.document_xml_file = "word/document.xml"
        self.document_xml_content = self.__load_document_xml()
        self.settings_xml_file = "word/settings.xml"
        self.settings_xml_content = self.__load_settings_xml()
        self.rsidRs = self.__extract_all_rsidr_from_summary_xml()

        self.rsidR_in_document_xml = self.__rsidr_in_document_xml()
        self.rsidRPr = self.__other_rsids_in_document_xml("rsidRPr")
        self.rsidP = self.__other_rsids_in_document_xml("rsidP")
        self.rsidRDefault = self.__other_rsids_in_document_xml("rsidRDefault")

        self.p_tags = len(re.findall(r'</w:t>', self.document_xml_content))
        self.r_tags = len(re.findall(r'</w:r>', self.document_xml_content))
        self.t_tags = len(re.findall(r'</w:t>', self.document_xml_content))

    def __load_core_xml(self):
        # load core.xml
        try:
            with zipfile.ZipFile(self.msword_file, 'r') as zipref:
                with zipref.open(self.core_xml_file) as xmlFile:
                    return xmlFile.read().decode("utf-8")
        except FileNotFoundError:
            print(f"File '{self.core_xml_file} not found in the DOCx archive.")
        except Exception as e:
            print(f"An error occurred: {e}")

    def __load_app_xml(self):
        # load app.xml
        try:
            with zipfile.ZipFile(self.msword_file, 'r') as zipref:
                with zipref.open(self.app_xml_file) as xmlFile:
                    return xmlFile.read().decode("utf-8")
        except FileNotFoundError:
            print(f"File '{self.app_xml_file} not found in the DOCx archive.")
        except Exception as e:
            print(f"An error occurred: {e}")

    def __load_document_xml(self):
        # load document.xml
        try:
            with zipfile.ZipFile(self.msword_file, 'r') as zipref:
                with zipref.open(self.document_xml_file) as xmlFile:
                    return xmlFile.read().decode("utf-8")
        except FileNotFoundError:
            print(f"File '{self.document_xml_file}' not found in the ZIP archive.")
        except Exception as e:
            print(f"An error occurred: {e}")

    def __load_settings_xml(self):
        try:
            with zipfile.ZipFile(self.msword_file, 'r') as zipref:
                with zipref.open(self.settings_xml_file) as xmlFile:
                    return xmlFile.read().decode("utf-8")
        except FileNotFoundError:
            print(f"File '{self.settings_xml_file}' not found in the ZIP archive.")
        except Exception as e:
            print(f"An error occurred: {e}")

    def __extract_all_rsidr_from_summary_xml(self):
        """
        function to extract all RSIDs at the beginning of the class. If you were to put this in the method,
        it would have to do this every time you called the method.
        :return:
        """
        rsids_list = []
        # Find all RSIDs, not rsidRoot. rsidRoot is repeated in rsids
        matches = re.findall(r'<w:rsid w:val="[^>]*/>', self.settings_xml_content)

        for match in matches:  # loops through all matches
            # greps for rsid using a group to extract the actual RSID from the string.
            rsid_match = re.search(r'<w:rsid w:val="([^"]*)"', match)
            if rsid_match:
                rsids_list.append(rsid_match.group(1))  # Appends it to the list
        return "" if len(rsids_list) == 0 else rsids_list

    def __rsidr_in_document_xml(self):
        """
        This function calculates the count of each rsidR in document.xml
        :return:
        """
        rsidr_count = {}
        for rsid in self.rsidRs:
            pattern = rf'w:rsidR="{rsid}"'
            rsidr_count[rsid] = len(re.findall(pattern, self.document_xml_content))
        return rsidr_count

    def __other_rsids_in_document_xml(self, rsid):
        """
        :param rsid tag name (e.g. "rsidRPr", "rsidP", "rsidDefault")
        The function accepts an rsid tag name as a parameter (e.g. rsidRPr, rsidP, rsidDefault).
        It searches document.xml for a pattern to find all instances of that rsid tag.
        It creates a dictionary that contains each unique rsid value as the key, and the count of how many times
        that rsid is in document.xml.
        E.g., {"00123456": 4, "00234567": 0, "00345678":11}

        :return: dictionary where the key is unique RSIDs, and the value is a count of the occurences of that rsid
        in document.xml
        """
        rsids = {}
        pattern = rf'w:{rsid}="........"'
        # Find all rsidRPr in document.xml file
        matches = re.findall(pattern, self.document_xml_content)

        for match in matches:  # loops through all matches
            # greps for rsid using a group to extract the actual RSID from the string.
            group_pattern = rf'w:{rsid}="(........)"'
            rsid_match = re.search(group_pattern, match)
            if rsid_match:
                if rsid_match.group(1) in rsids:
                    rsids[rsid_match.group(1)] += 1  # Appends it to the list
                else:
                    rsids[rsid_match.group(1)] = 1
        return rsids

    def filename(self):
        """
        :return: the filename of the DOCx file passed to the class
        """
        return self.msword_file

    def xml_files(self):
        """
        :return: A dictionary in the following format: {XML filename: [file size, file hash]}
        """
        with zipfile.ZipFile(self.msword_file, 'r') as zip_file:
            # returns XML files in the DOCx
            xml_files = {}
            for file_info in zip_file.infolist():
                with zipfile.ZipFile(self.msword_file, 'r') as zip_ref:
                    with zip_ref.open(file_info.filename) as xml_file:
                        md5hash = hashlib.md5(xml_file.read()).hexdigest()
                xml_files[file_info.filename] = [file_info.file_size, md5hash]
            return xml_files  # returns dictionary {xml_filename: [file size, file hash]}

    def xml_hash(self, xmlfile):
        """
        :param xmlfile:
        :return: the hash of a specified XML file
        """
        return self.xml_files()[xmlfile][1]

    def xml_size(self, xmlfile):
        """
        :param xmlfile:
        :return: the size of a specified XML file
        """
        return self.xml_files()[xmlfile][0]

    def title(self):
        """
        :return: the title metadata in core.xml
        """
        doc_title = re.search(r'<dc:title>(.*?)</dc:title>', self.core_xml_content)
        return "" if doc_title is None else doc_title.group(1)

    def subject(self):
        """
        :return: the subject metadata from core.xml
        """
        doc_subject = re.search(r'<dc:subject>(.*?)</dc:subject>', self.core_xml_content)
        return "" if doc_subject is None else doc_subject.group(1)

    def creator(self):
        """
        :return: the creator metadata from core.xml
        """
        doc_creator = re.search(r'<dc:creator>(.*?)</dc:creator>', self.core_xml_content)
        return "" if doc_creator is None else doc_creator.group(1)

    def keywords(self):
        """
        :return: the keywords metadata from core.xml
        """
        doc_keywords = re.search(r'<cp:keywords>(.*?)</cp:keywords>', self.core_xml_content)
        return "" if doc_keywords is None else doc_keywords.group(1)

    def description(self):
        """
        :return: the description metadata from core.xml
        """
        doc_description = re.search(r'<dc:description>(.*?)</dc:description>', self.core_xml_content)
        return "" if doc_description is None else doc_description.group(1)

    def revision(self):
        """
        :return: the revision # metadata from core.xml
        """
        doc_revision = re.search(r'<cp:revision>(.*?)</cp:revision>', self.core_xml_content)
        return "" if doc_revision is None else doc_revision.group(1)

    def created(self):
        """
        :return: the created date metadata from core.xml
        """
        doc_created = re.search(r'<dcterms:created[^>].*?>(.*?)</dcterms:created>', self.core_xml_content)
        return "" if doc_created is None else doc_created.group(1)

    def modified(self):
        """
        :return: the modified date metadata from core.xml
        """
        doc_modified = re.search(r'<dcterms:modified[^>].*?>(.*?)</dcterms:modified>', self.core_xml_content)
        return "" if doc_modified is None else doc_modified.group(1)

    def last_modified_by(self):
        """
        :return: the last modified by metadata from core.xml
        """
        doc_lastmodifiedby = re.search(r'<cp:lastModifiedBy>(.*?)</cp:lastModifiedBy>', self.core_xml_content)
        return "" if doc_lastmodifiedby is None else doc_lastmodifiedby.group(1)

    def last_printed(self):
        """
        :return: the last printed date metadata from core.xml
        """
        doc_lastprinted = re.search(r'<cp:lastPrinted>(.*?)</cp:lastPrinted>', self.core_xml_content)
        return "" if doc_lastprinted is None else doc_lastprinted.group(1)

    def category(self):
        """
        :return: the category metadata from core.xml
        """
        doc_category = re.search(r'<cp:category>(.*?)</cp:category>', self.core_xml_content)
        return "" if doc_category is None else doc_category.group(1)

    def content_status(self):
        """
        :return: the content status metadata from core.xml
        """
        doc_contentstatus = re.search(r'<cp:contentStatus>(.*?)</cp:contentStatus>', self.core_xml_content)
        return "" if doc_contentstatus is None else doc_contentstatus.group(1)

    def template(self):
        """
        :return: the template metadata from app.xml
        """
        doc_template = re.search(r'<Template>(.*?)</Template>', self.app_xml_content)
        return "" if doc_template is None else doc_template.group(1)

    def total_editing_time(self):
        """
        :return: the total editing time in minutes metadata from app.xml
        """
        doc_edit_time = re.search(r'<TotalTime>(.*?)</TotalTime>', self.app_xml_content)
        return "" if doc_edit_time is None else doc_edit_time.group(1)

    def pages(self):
        """
        :return: the # of pages in the document metadata from app.xml
        Note: the author has observed that in some cases, this is not properly updated within the XML file itself.
        It is not an error in the script. It's an error in the metadata. Opening the document and allowing it to
        fully load and then saving it updates this. But of course, it changes other metadata as well if you do that.
        """
        doc_pages = re.search(r'<Pages>(.*?)</Pages>', self.app_xml_content)
        return "" if doc_pages is None else doc_pages.group(1)

    def words(self):
        """
        :return: the number of words in the document metadata from app.xml
        """
        doc_words = re.search(r'<Words>(.*?)</Words>', self.app_xml_content)
        return "" if doc_words is None else doc_words.group(1)

    def characters(self):
        """
        :return: the number of characters in the document metadata from app.xml
        """
        doc_characters = re.search(r'<Characters>(.*?)</Characters>', self.app_xml_content)
        return "" if doc_characters is None else doc_characters.group(1)

    def application(self):
        """
        :return: the application name that created the document metadata from app.xml
        """
        doc_application = re.search(r'<Application>(.*?)</Application>', self.app_xml_content)
        return "" if doc_application is None else doc_application.group(1)

    def security(self):
        """
        :return: the security metadata from app.xml
        """
        doc_security = re.search(r'<DocSecurity>(.*?)</DocSecurity>', self.app_xml_content)
        return "" if doc_security is None else doc_security.group(1)

    def lines(self):
        """
        :return: the number of lines in the document metadata from app.xml
        """
        doc_lines = re.search(r'<Lines>(.*?)</Lines>', self.app_xml_content)
        return "" if doc_lines is None else doc_lines.group(1)

    def paragraphs(self):
        """
        :return: the number of paragraphs in the document metadata from app.xml
        Note: similar to # of pages, the author has noted in testing that sometimes, this may not be accurate in
        the metadata for some reason. It's not an error in this program. It's an error with the metadata itself
        in the document.
        """
        doc_paragraphs = re.search(r'<Paragraphs>(.*?)</Paragraphs>', self.app_xml_content)
        return "" if doc_paragraphs is None else doc_paragraphs.group(1)

    def characters_with_spaces(self):
        """
        :return: the total characters including spaces in the document metadatafrom app.xml
        """
        doc_characters_with_spaces = re.search(r'<CharactersWithSpaces>(.*?)</CharactersWithSpaces>',
                                               self.app_xml_content)
        return "" if doc_characters_with_spaces is None else doc_characters_with_spaces.group(1)

    def app_version(self):
        """
        :return: the version of the app that created the document metadatafrom app.xml
        """
        doc_app_version = re.search(r'<AppVersion>(.*?)</AppVersion>', self.app_xml_content)
        return "" if doc_app_version is None else doc_app_version.group(1)

    def manager(self):
        """
        :return: the manager metadata from app.xml
        """
        doc_manager = re.search(r'<Manager>(.*?)</Manager>', self.app_xml_content)
        return "" if doc_manager is None else doc_manager.group(1)

    def company(self):
        """
        :return: the company metadata from app.xml
        """
        doc_company = re.search(r'<Company>(.*?)</Company>', self.app_xml_content)
        return "" if doc_company is None else doc_company.group(1)

    def paragraph_tags(self):
        """
        :return: the total number of paragraph tags in document.xml
        """
        return self.p_tags

    def runs_tags(self):
        """
        :return: the total number of runs tags in document.xml
        """
        return self.r_tags

    def text_tags(self):
        """
        :return: the total number of text tags in document.xml
        """
        return self.t_tags

    def rsid_root(self):
        """
        :return: rsidRoot from settings.xml
        """
        root = re.search(r'<w:rsidRoot w:val="([^"]*)"', self.settings_xml_content).group(1)
        return "" if root is None else root

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

    def __str__(self):
        """
        :return: a text string that you can print out to get a summary of the document.
        This can be edited to suit your needs. You can naturally accomplish the same results by calling each of
        the methods in your print statement in the main script.
        """
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
