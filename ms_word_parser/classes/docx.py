import hashlib
import struct
import xml.etree.ElementTree as ET
import zipfile
from zipfile import BadZipFile
from datetime import datetime as dt

try:
    from classes.datastore import DataStore
except ModuleNotFoundError:
    from ms_word_parser.classes.datastore import DataStore

__dtfmt__ = "%Y-%m-%d %H:%M:%S"

class Docx:
    """
    Accepts a docx file. Has the following methods to extract data from core.xml, app.xml, document.xml

    app_version, application, category, characters, characters_with_spaces, company, content_status, created, creator,
    description, filename, keywords, last_modified_by, last_printed, lines, manager, modified, pages, paragraph_tags,
    paragraphs, revision, runs_tags, security, subject, template, text_tags, title, total_editing_time, words,
    xml_files, xml_hash, xml_size
    """

    def __init__(
        self, msword_file, triage=False, hashing=True, store: DataStore = None
    ):
        """
        .docx file to pass to the class
        Triage value can be True or False. If True, will parse less info to execute faster.
        When set to False, it does not try to parse RSID values from document.xml.
        If triage value not passed, it defaults to False and does full parsing.
        The script using this class still ultimately decides what methods it wants to use.
        But if in triage mode, some of the variables will not get assigned any value, thus
        will affect any methods that rely on those variables having a value assigned to them.
        """
        if store is None:
            store = DataStore()
        self.store = store
        if store.ms_word_gui:
            update_status = store.ms_word_gui.update_status
        else:
            update_status = lambda msg, **kwargs: update_cli(msg, store=store, **kwargs)
        self.update_status = update_status
        self.item_files = []
        self.ink_files = []
        self.xml_files = {}
        self.namespaces = {
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "aink": "http://schemas.microsoft.com/office/drawing/2016/ink",
            "b": "http://schemas.openxmlformats.org/officeDocument/2006/bibliography",
            "ct": "http://schemas.microsoft.com/office/2006/metadata/contentType",
            "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            "cprop": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
            "cr": "http://schemas.microsoft.com/office/comments/2020/reactions",
            "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
            "dc": "http://purl.org/dc/elements/1.1/",
            "dcterms": "http://purl.org/dc/terms/",
            "dcmitype": "http://purl.org/dc/dcmitype/",
            "default": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
            "ds": "http://schemas.openxmlformats.org/officeDocument/2006/customXml",
            "inkml": "http://www.w3.org/2003/InkML",
            "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "ma": "http://schemas.microsoft.com/office/2006/metadata/properties/metaAttributes",
            "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
            "o": "urn:schemas-microsoft-com:office:office",
            "oel": "http://schemas.microsoft.com/office/2019/extlst",
            "p": "http://schemas.microsoft.com/office/2006/metadata/properties",
            "pc": "http://schemas.microsoft.com/office/infopath/2007/PartnerControls",
            "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "sc": "Microsoft.SharePoint.Taxonomy.ContentTypeSync",
            "sp": "http://schemas.microsoft.com/sharepoint/v3",
            "v": "urn:schemas-microsoft-com:vml",
            "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
            "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
            "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
            "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
            "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
            "w16du": "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
            "w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
            "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
            "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
            "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
            "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
            "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
            "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
            "xs": "http://www.w3.org/2001/XMLSchema",
            "xsd": "http://www.w3.org/2001/XMLSchema",
            "xsi": "http://www.w3.org/2001/XMLSchema-instance",
        }
        self.has_ink = False
        self.has_comments = False
        self.msword_file = msword_file
        self.hashing = hashing
        self.header_offsets, self.binary_content = self.__find_binary_string()
        self.extra_fields = self.__xml_extra_bytes()
        self.__load_all_xml()
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
        self.tr_tags = x.findall(".//w:tr", self.namespaces)
        self.shapedata = x.findall(".//v:shape", self.namespaces)
        self.drawing_tags = x.findall(".//w:drawing", self.namespaces)
        if self.drawing_tags or self.ink_files:
            self.has_ink = True
        if not triage:  # if not run in triage mode, do full parsing
            self.rsidR_in_document_xml = self.__rsids_in_document_xml("rsidR")
            self.rsidRPr = self.__rsids_in_document_xml("rsidRPr")
            self.rsidP = self.__rsids_in_document_xml("rsidP")
            self.rsidRDefault = self.__rsids_in_document_xml("rsidRDefault")
            self.rsidTr = self.__rsids_in_document_xml("rsidTr")
            self.para_id = self.__rsids_in_document_xml("paraId")
            self.text_id = self.__rsids_in_document_xml("textId")

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.core_xml_content = None
        self.app_xml_content = None
        self.document_xml_content = None
        self.comments_xml_content = None
        self.settings_xml_content = None
        self.people_xml_content = None
        self.extensible_xml_content = None
        self.extended_xml_content = None
        self.comments_ids_content = None
        self.custom_xml_content = None

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
        extras = {}
        truncate_extra_field = 20  # extra field can be several hundred bytes, mostly 0x00. This grabs the first 20.

        for offset in self.header_offsets:
            (
                filename_len,
                extrafield_len,
            ) = struct.unpack("<2H", self.binary_content[offset + 26 : offset + 30])
            filename_start = offset + 30
            filename_end = offset + 30 + filename_len
            if filename_end - filename_start < 256:
                # some DOCx files somehow produce false positives of
                # excessively long filenames and results in an error. This avoids that error.
                filename = self.binary_content[filename_start:filename_end].decode(
                    "ascii"
                )
            extrafield_start = filename_end
            extrafield_end = extrafield_start + extrafield_len
            extrafield = self.binary_content[extrafield_start:extrafield_end]
            extrafield_hex_as_text = []

            for h in extrafield:
                extrafield_hex_as_text.append(f"{h:02x}")

            if not extrafield:
                extras[filename] = [extrafield_len, "nil"]
            elif (
                extrafield_len <= truncate_extra_field
            ):  # field size larger than truncate value
                extras[filename] = [
                    extrafield_len,
                    f"0x{''.join(extrafield_hex_as_text)}",
                ]
            else:
                extras[filename] = [
                    extrafield_len,
                    f"0x{''.join(extrafield_hex_as_text[0:truncate_extra_field])}",
                ]  # adds only
                # the select # of characters as specified in the variable truncate_extra_field. This is so that
                # we don't end up with hundreds of characters in a cell in Excel, as some extra fields can be
                # several hundred values long. But so far, most are 0x00, with only the first few being values other
                # than hex 0x00.

        return extras

    def __load_xml(self, xml_file):
        content = ""
        if (
            xml_file in self.get_xml_files()
        ):  # if the file exists, read it and return its content
            if "comments.xml" in xml_file:
                self.has_comments = True
            with zipfile.ZipFile(self.msword_file, "r") as zipref:
                with zipref.open(xml_file) as xmlFile:
                    content = xmlFile.read()
        return content

    def __load_all_xml(self):
        xml_files = {}
        blank = {
            "MD5": None,
            "Modified Time": None,
            "File Size": None,
            "Zip Compression": None,
            "Zip Create System": None,
            "Zip Create Version": None,
            "Zip Extract Version": None,
            "Zip Flag Bits": None,
            "Zip Extra Fields 1": None,
            "Zip Extra Fields 2": None,
        }
        xml_map = {
            "core_xml_content": "docProps/core.xml",
            "app_xml_content": "docProps/app.xml",
            "document_xml_content": "word/document.xml",
            "comments_xml_content": "word/comments.xml",
            "settings_xml_content": "word/settings.xml",
            "people_xml_content": "word/people.xml",
            "extensible_xml_content": "word/commentsExtensible.xml",
            "extended_xml_content": "word/commentsExtended.xml",
            "comments_ids_content": "word/commentsIds.xml",
            "custom_xml_content": "docProps/custom.xml",
        }

        modified_time = None
        compression_types = {0: "Store (None)", 8: "DEFLATE"}
        try:
            with zipfile.ZipFile(self.msword_file, "r") as zipref:
                zip_filenames = zipref.namelist()
                zip_info = zipref.infolist()
                for xml in zip_info:
                    md5hash = None
                    xml_name = xml.filename
                    if xml_name not in self.extra_fields:
                        xml_name = xml_name.replace("/", "\\")
                    xml_files[xml_name] = blank.copy()
                    if (
                        "customXml/item" in xml_name
                        and "Props" not in xml_name
                        and xml_name not in self.item_files
                    ):
                        self.item_files.append(xml_name)
                    if "ink/ink" in xml_name and xml_name not in self.ink_files:
                        self.ink_files.append(xml_name)
                    if self.hashing:
                        with zipref.open(xml_name) as xml_file:
                            content = xml_file.read()
                            md5hash = self.hash(content)
                    m_time = xml.date_time
                    if m_time not in ((1980, 1, 1, 0, 0, 0), (1980, 0, 0, 0, 0, 0)):
                        modified_time = dt(*m_time).strftime(__dtfmt__)
                    xml_files[xml_name]["MD5"] = md5hash
                    xml_files[xml_name]["Modified Time"] = modified_time
                    xml_files[xml_name]["File Size"] = xml.file_size
                    xml_files[xml_name][
                        "Zip Compression"
                    ] = f'{str(xml.compress_type)}: {compression_types.get(xml.compress_type, "Unidentified")}'
                    xml_files[xml_name]["Zip Create System"] = xml.create_system
                    xml_files[xml_name]["Zip Create Version"] = xml.create_version
                    xml_files[xml_name]["Zip Extract Version"] = xml.extract_version
                    xml_files[xml_name]["Zip Flag Bits"] = f"{xml.flag_bits:#0{6}x}"
                    xml_files[xml_name]["Zip Extra Fields Length"] = self.extra_fields[
                        xml_name
                    ][0]
                    xml_files[xml_name]["Zip Extra Fields Bytes"] = self.extra_fields[
                        xml_name
                    ][1]
                for attrib, file_path in xml_map.items():
                    alt_path = file_path.replace("/", "\\")
                    target = file_path if file_path in zip_filenames else alt_path
                    if target in zip_filenames:
                        if "comments.xml" in target:
                            self.has_comments = True
                        content = zipref.read(target)
                        setattr(self, attrib, content)
                    else:
                        setattr(self, attrib, "")
            self.xml_files = xml_files
        except (BadZipFile, FileNotFoundError) as e:
            raise Exception(f"Error accessing {self.msword_file}: {e}") from e
        return self.xml_files

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
            meta_content = ns_extract.text if ns_extract is not None else None
        else:
            return None
        return meta_content

    def get_people(self):
        if self.people_xml_content != "":
            xml = ET.fromstring(self.people_xml_content)
            list_of_people = []
            all_people = xml.findall(".//w15:person", self.namespaces)
            for person in all_people:
                author = person.get(f"{{{self.namespaces['w15']}}}author")
                if len(person) > 0:
                    providerId = person[0].get(
                        f"{{{self.namespaces['w15']}}}providerId"
                    )
                    userId = person[0].get(f"{{{self.namespaces['w15']}}}userId")
                else:
                    providerId = userId = None
                list_of_people.append([author, providerId, userId])
            return list_of_people
        return None

    def any_comments(self):
        return self.has_comments

    def get_comments(self):
        """
        return the list all_comments that contains the following:
            Comment ID,
            Timestamp,
            Author,
            Initials,
            Text
        :return:
        """

        if not self.has_comments:
            return [None, None, None, None, None]
        xml = ET.fromstring(self.comments_xml_content)
        # Find all comments
        comments = xml.findall(".//w:comment", self.namespaces)
        all_comments = []
        for comment in comments:
            author = comment.get(f"{{{self.namespaces['w']}}}author")
            date_time = comment.get(f"{{{self.namespaces['w']}}}date")
            initials = comment.get(f"{{{self.namespaces['w']}}}initials")
            comment_id = comment.get(f"{{{self.namespaces['w']}}}id")
            comment_paras = comment.findall(".//w:p", self.namespaces)
            text = (
                "\n".join(
                    [
                        t.text
                        for t in comment.findall(".//w:t", self.namespaces)
                        if t.text
                    ]
                )
                .encode("utf-8", "surrogatepass")
                .decode()
            )
            if len(comment_paras) > 0:
                comment_paraId = comment_paras[-1].get(
                    f"{{{self.namespaces['w14']}}}paraId"
                )
            else:
                comment_paraId = None
            all_comments.append(
                [comment_id, comment_paraId, date_time, author, initials, text]
            )
        return all_comments

    def get_comments_ids(self):
        if self.comments_ids_content != "":
            all_comments_ids = []
            xml = ET.fromstring(self.comments_ids_content)
            comments_ids = xml.findall(".//w16cid:commentId", self.namespaces)
            for comment_id in comments_ids:
                paraId = comment_id.get(f"{{{self.namespaces['w16cid']}}}paraId", "")
                durableId = comment_id.get(
                    f"{{{self.namespaces['w16cid']}}}durableId", ""
                )
                all_comments_ids.append([paraId, durableId])
            return all_comments_ids
        return None

    def get_extended_comments(self):
        if self.extended_xml_content != "":
            all_extended_comments = []
            xml = ET.fromstring(self.extended_xml_content)
            extended_comments = xml.findall(".//w15:commentEx", self.namespaces)
            for values in extended_comments:
                paraId = values.get(f"{{{self.namespaces['w15']}}}paraId")
                done = values.get(f"{{{self.namespaces['w15']}}}done")
                paraIdParent = values.get(
                    f"{{{self.namespaces['w15']}}}paraIdParent", "IS_PARENT"
                )
                all_extended_comments.append([paraId, paraIdParent, done])
            return all_extended_comments
        return None

    def get_extensible_comments(self):
        if self.extensible_xml_content != "":
            all_extensible_comments = {}
            xml = ET.fromstring(self.extensible_xml_content)
            extensible_comments = xml.findall(
                ".//w16cex:commentExtensible", self.namespaces
            )
            reaction_types = {0: "Unknown", 1: "Like", 2: "Unknown"}
            for values in extensible_comments:
                uri = "None"
                reactionType = "None"
                userId = userProvider = userName = ""
                durableId = values.get(f"{{{self.namespaces['w16cex']}}}durableId")
                dateUtc = values.get(f"{{{self.namespaces['w16cex']}}}dateUtc")
                extLst = values.findall(".//w16cex:extLst", self.namespaces)
                all_extensible_comments[durableId] = []
                all_extensible_comments[durableId].append(dateUtc)
                if extLst:
                    ext = extLst[0].find("w16:ext", self.namespaces)
                    uri = ext.get(f"{{{self.namespaces['w16']}}}uri")
                    all_extensible_comments[durableId].append(uri)
                    for entry in ext.findall(".//cr:reaction", self.namespaces):
                        reactionType = entry.get("reactionType", "")
                        all_extensible_comments[durableId].append(
                            reaction_types[int(reactionType)]
                        )
                        for reactionInfo in entry.findall(
                            ".//cr:reactionInfo", self.namespaces
                        ):
                            reactionDateUtc = reactionInfo.get("dateUtc", "")
                            user = reactionInfo.find("cr:user", self.namespaces)
                            if user is not None:
                                userId = user.get("userId", "")
                                userProvider = user.get("userProvider", "")
                                userName = user.get("userName", "")
                            all_extensible_comments[durableId].append(
                                [reactionDateUtc, userId, userProvider, userName]
                            )
                else:
                    all_extensible_comments[durableId].append(uri)
                    all_extensible_comments[durableId].append(reactionType)
                    all_extensible_comments[durableId].append(["", "", "", ""])
            return all_extensible_comments
        return None

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
            "rsidTr": self.namespaces["w"],
            "paraId": self.namespaces["w14"],
            "textId": self.namespaces["w14"],
        }
        for entry in (self.p_tags, self.r_tags, self.t_tags, self.tr_tags):
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

    def hash(self, content=None):
        """
        Function that will return the hash of the file itself
        """
        if self.hashing:  # if hashing option was selected
            filehash = hashlib.md5()
            if content is None:
                filehash.update(self.binary_content)
            else:
                filehash.update(content)
            return filehash.hexdigest().upper()
        return None  # if no hashing was selected.

    def get_xml_files(self):
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
        compression_types = {0: "Store (None)", 8: "DEFLATE"}
        md5hash = None
        with zipfile.ZipFile(self.msword_file, "r") as zip_file:
            xml_files = {}
            for file_info in zip_file.infolist():
                if (
                    "customXml/item" in file_info.filename
                    and "Props" not in file_info.filename
                    and file_info.filename not in self.item_files
                ):
                    self.item_files.append(file_info.filename)
                if (
                    "ink/ink" in file_info.filename
                    and file_info.filename not in self.ink_files
                ):
                    self.ink_files.append(file_info.filename)
                with zipfile.ZipFile(self.msword_file, "r") as zip_ref:
                    try:
                        with zip_ref.open(file_info.filename) as xml_file:
                            if self.hashing:  # if hashing option selected
                                md5hash = self.hash(xml_file.read())
                            else:
                                md5hash = "Option Not Selected"  # else return blank for hash value.
                    except BadZipFile:
                        pass
                    except OSError as exc:
                        raise Exception(
                            "Error processing the zip file header - likely offset is incorrect."
                        ) from exc
                m_time = file_info.date_time
                if m_time in ((1980, 1, 1, 0, 0, 0), (1980, 0, 0, 0, 0, 0)):
                    modified_time = None
                else:
                    modified_time = dt(*m_time).strftime(__dtfmt__)
                fname = file_info.filename
                if fname not in self.extra_fields:
                    fname = fname.replace("/", "\\")
                xml_files[file_info.filename] = [
                    md5hash,
                    modified_time,
                    file_info.file_size,
                    f'{str(file_info.compress_type)}: {compression_types.get(file_info.compress_type, "Unidentified")}',
                    file_info.create_system,
                    file_info.create_version,
                    file_info.extract_version,
                    f"{file_info.flag_bits:#0{6}x}",
                    self.extra_fields[fname][0],
                    self.extra_fields[fname][1],
                ]
            return xml_files

    def xml_hash(self, xmlfile: str):
        """
        :param: xmlfile
        :return: the hash of a specified XML file
        """
        return self.xml_files[xmlfile]["MD5"]

    def xml_size(self, xmlfile: str):
        """
        :param: xmlfile
        :return: the size of a specified XML file
        """
        return self.xml_files[xmlfile]["File Size"]

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

    def table_row_tags(self):
        """
        :return: the total number of table row tags in document.xml
        """
        return len(self.tr_tags)

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
        return None if root is None else root

    def get_doc_ids(self):
        """
        :return: the w14, w15, and w16 docId's from settings.xml
        """
        x = ET.fromstring(self.settings_xml_content)
        w14_id = w15_id = w16_id = "None"
        w14_ns = x.find(f"{{{self.namespaces['w14']}}}docId")
        if w14_ns is not None:
            w14_id = w14_ns.get(f"{{{self.namespaces['w14']}}}val", "None")
        w15_ns = x.find(f"{{{self.namespaces['w15']}}}docId")
        if w15_ns is not None:
            w15_id = w15_ns.get(f"{{{self.namespaces['w15']}}}val", "None")
        w16_ns = x.find(f"{{{self.namespaces['w16']}}}docId")
        if w16_ns is not None:
            w16_id = w16_ns.get(f"{{{self.namespaces['w16']}}}val", "None")

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

    def rsidtr_in_document_xml(self):
        """
        return dictionary with unique rsidTr and count of how many times it is found in document.xml
        :return:
        """
        return self.rsidTr

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

    def get_proof_state(self):
        xml = ET.fromstring(self.settings_xml_content)
        proof_state = xml.find(f"{{{self.namespaces['w']}}}proofState")
        spelling = grammar = "None"
        if proof_state is not None:
            spelling = proof_state.get(f"{{{self.namespaces['w']}}}spelling", "None")
            grammar = proof_state.get(f"{{{self.namespaces['w']}}}grammar", "None")

        return [spelling, grammar]

    def get_custom_xml(self):
        if self.custom_xml_content:
            props = {}
            xml = ET.fromstring(self.custom_xml_content)
            for cprop in xml.findall(".//cprop:property", self.namespaces):
                attribs = cprop.attrib
                for attr_name, attr_val in attribs.items():
                    props[attr_name] = attr_val
                for sub_prop in cprop:
                    tag = (
                        sub_prop.tag.split("}", 1)[1]
                        if "}" in sub_prop.tag
                        else sub_prop.tag
                    )
                    value = sub_prop.text
                    props[tag] = value
            return props
        return None

    def get_all_content(self, files):
        if files:
            content = {self.msword_file: {}}
            for file in files:
                content[self.msword_file][file] = {}
                xml_content = self.__load_xml(file)
                if xml_content == "":
                    continue
                if b"<?mso-contentType?>" in xml_content:
                    xml_content = (
                        xml_content.replace(b"<?mso-contentType?>", b"")
                    ).decode("utf-8")
                xml = ET.fromstring(xml_content)
                for element in xml.iter():
                    tag = (
                        element.tag.split("}")[-1]
                        if "}" in element.tag
                        else element.tag
                    )
                    if tag not in content[self.msword_file][file]:
                        content[self.msword_file][file][tag] = []
                    attribs = {}
                    for name, value in element.attrib.items():
                        name = name.split("}", 1)[-1] if "}" in name else name
                        attribs[name] = value
                    text = (element.text or "").strip()
                    if text:
                        attribs["_text"] = text
                    tail = (element.tail or "").strip()
                    if tail:
                        attribs["_tail"] = tail
                    child_tags = list(element)
                    if child_tags:
                        attribs["_children"] = []
                        for child in child_tags:
                            attribs["_children"].append(
                                child.tag.split("}")[-1]
                                if "}" in child.tag
                                else child.tag
                            )
                    content[self.msword_file][file][tag].append(attribs)
            return content
        return None

    def get_ink(self):
        ts_data = []
        for ink_file in self.ink_files:
            load_ink = self.__load_xml(ink_file)
            xml = ET.fromstring(load_ink)
            for element in xml.iter():
                tag = element.tag.split("}")[-1] if "}" in element.tag else element.tag
                if tag == "timestamp":
                    (ts_ns, ts_id), (timestring, ts) = element.attrib.items()
            ts_data.append([ink_file, ts])
        return ts_data

    def adjust_timestamp(self, ts):
        if ts:
            adjusted_timestamp = ts.replace("T", " ").replace("Z", "")
            return adjusted_timestamp.split(".")[0]
        return ""
