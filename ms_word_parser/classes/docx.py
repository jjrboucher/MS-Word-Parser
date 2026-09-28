import hashlib
import re
import struct
from collections import Counter
import xml.etree.ElementTree as ET
import zipfile
from zipfile import BadZipFile
from datetime import datetime as dt
import olefile

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
    xml_files
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
        self.item_files = []
        self.ink_files = []
        self.xml_files = {}
        self.protection_state = {"enabled": False}
        self.is_encrypted = False
        self.irm_info = {}
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
            "Relationships": "http://schemas.openxmlformats.org/package/2006/relationships",
            "sc": "Microsoft.SharePoint.Taxonomy.ContentTypeSync",
            "sp": "http://schemas.microsoft.com/sharepoint/v3",
            "Types": "http://schemas.openxmlformats.org/package/2006/content-types",
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
        self._parsed = {}
        self._zip = None
        self.has_ink = False
        self.has_comments = False
        self.msword_file = msword_file
        self.hashing = hashing
        if not zipfile.is_zipfile(msword_file):
            try:
                is_ole = olefile.isOleFile(msword_file)
            except (FileNotFoundError, OSError) as e:
                raise Exception(f"Error accessing {msword_file}: {e}") from e
            if not is_ole:
                raise Exception(
                    f"Error accessing {msword_file}: File is not a zip file"
                )
            self.is_encrypted = True
            self.__init_encrypted()
            return
        self._zip = zipfile.ZipFile(msword_file, "r")
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
        x = self.__parse(self.document_xml_content)
        self.p_tags = x.findall(".//w:p", self.namespaces)
        self.r_tags = x.findall(".//w:r", self.namespaces)
        self.t_tags = x.findall(".//w:t", self.namespaces)
        self.tr_tags = x.findall(".//w:tr", self.namespaces)
        ## TODO self.shapedata = x.findall(".//v:shape", self.namespaces)
        self.drawing_tags = x.findall(".//w:drawing", self.namespaces)
        if self.drawing_tags or self.ink_files:
            self.has_ink = True
        if not triage:  # if not run in triage mode, do full parsing
            counts = self.__count_rsids()
            self.rsidR_in_document_xml = {r: counts["rsidR"][r] for r in self.rsidRs}
            self.rsidRPr = dict(counts["rsidRPr"])
            self.rsidP = dict(counts["rsidP"])
            self.rsidRDefault = dict(counts["rsidRDefault"])
            self.rsidTr = dict(counts["rsidTr"])
            self.para_id = dict(counts["paraId"])
            self.text_id = dict(counts["textId"])

    def __init_encrypted(self):
        self._parsed[""] = ET.fromstring("<empty/>")
        self.extra_fields = {}
        self.xml_files = {}
        self.rsidRs = ""
        self.rsidR_in_document_xml = {}
        self.rsidRPr = {}
        self.rsidP = {}
        self.rsidRDefault = {}
        self.rsidTr = {}
        self.para_id = {}
        self.text_id = {}
        self.p_tags = []
        self.r_tags = []
        self.t_tags = []
        self.tr_tags = []
        self.drawing_tags = []
        for attrib in (
            "core_xml_content",
            "app_xml_content",
            "document_xml_content",
            "comments_xml_content",
            "settings_xml_content",
            "people_xml_content",
            "extensible_xml_content",
            "extended_xml_content",
            "comments_ids_content",
            "custom_xml_content",
            "content_types_content",
            "xml_rels_content",
        ):
            setattr(self, attrib, "")
        self.ns_lookup = {
            "title": ["", "dc"],
            "subject": ["", "dc"],
            "creator": ["", "dc"],
            "keywords": ["", "cp"],
            "description": ["", "dc"],
            "revision": ["", "cp"],
            "created": ["", "dcterms"],
            "modified": ["", "dcterms"],
            "lastModifiedBy": ["", "cp"],
            "lastPrinted": ["", "cp"],
            "category": ["", "cp"],
            "contentStatus": ["", "cp"],
            "language": ["", "dc"],
            "version": ["", "cp"],
            "Template": ["", "default"],
            "TotalTime": ["", "default"],
            "Pages": ["", "default"],
            "Words": ["", "default"],
            "Characters": ["", "default"],
            "Application": ["", "default"],
            "DocSecurity": ["", "default"],
            "Lines": ["", "default"],
            "Paragraphs": ["", "default"],
            "CharactersWithSpaces": ["", "default"],
            "AppVersion": ["", "default"],
            "Manager": ["", "default"],
            "Company": ["", "default"],
            "SharedDoc": ["", "default"],
            "HyperlinksChanged": ["", "default"],
        }
        self.irm_info = self.__extract_irm_info()

    def __extract_irm_info(self):
        """
        Reads the IRM license out of the OLE container if present.
        """
        stream_path = [
            "\x06DataSpaces",
            "TransformInfo",
            "DRMEncryptedTransform",
            "\x06Primary",
        ]
        try:
            ole = olefile.OleFileIO(self.msword_file)
        except Exception:
            return {}
        try:
            if not ole.exists(stream_path):
                return {}
            data = ole.openstream(stream_path).read()
        except Exception:
            return {}
        finally:
            ole.close()

        start = data.find(b"<XrML")
        end = data.find(b"</XrML>")
        if start == -1 or end == -1:
            return {}
        try:
            license_xml = ET.fromstring(data[start : end + len(b"</XrML>")])
        except ET.ParseError:
            return {}

        def text_of(path):
            element = license_xml.find(path)
            return element.text if element is not None else None

        def xrml_time(path):
            value = text_of(path)
            if value and re.fullmatch(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}", value):
                value += ":00"
            return value

        template_raw = text_of(".//DESCRIPTOR/OBJECT/NAME")
        template_name = template_description = None
        if template_raw:
            match = re.search(
                r"NAME (?P<name>[^:]*):DESCRIPTION (?P<description>.*?);?\s*$",
                template_raw,
                re.S,
            )
            if match:
                template_name = match.group("name").strip()
                template_description = match.group("description").strip()
            else:
                template_name = template_raw.strip()

        server = license_xml.find(
            ".//DISTRIBUTIONPOINT/OBJECT[@type='License-Acquisition-URL']/ADDRESS"
        )
        tenant = license_xml.find(".//SECURITYLEVEL[@name='Tenant-ID']")
        sdk = license_xml.find(".//SECURITYLEVEL[@name='SDK']")

        return {
            "owner": text_of(".//WORK/METADATA/OWNER/OBJECT/NAME"),
            "issuer": text_of("./BODY/ISSUER/OBJECT/NAME"),
            "organization": text_of(".//ISSUEDPRINCIPALS/PRINCIPAL/OBJECT/NAME"),
            "template_name": template_name,
            "template_description": template_description,
            "rms_server": server.text if server is not None else None,
            "tenant_id": tenant.get("value") if tenant is not None else None,
            "sdk_version": sdk.get("value") if sdk is not None else None,
            "issued": xrml_time("./BODY/ISSUEDTIME"),
            "valid_from": xrml_time("./BODY/VALIDITYTIME/FROM"),
            "valid_until": xrml_time("./BODY/VALIDITYTIME/UNTIL"),
        }

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
        self.content_types_content = None
        self.xml_rels_content = None
        self._parsed = {}
        if self._zip is not None:
            self._zip.close()
            self._zip = None

    def __parse(self, content):
        """
        Parse all XML content once and store data for lookup
        """
        tree = self._parsed.get(content)
        if tree is None:
            tree = ET.fromstring(content)
            self._parsed[content] = tree
        return tree

    def __xml_extra_bytes(self):
        """
        ref: https://en.wikipedia.org/wiki/ZIP_(file_format)#Local_file_header

        return: {archive member name: [# of bytes in extra field, truncated bytes]}
        """
        extras = {}
        truncate_extra_field = 20  # extra field can be several hundred bytes, mostly 0x00. This grabs the first 20.
        infolist = self._zip.infolist()
        with open(self.msword_file, "rb") as msword_binary:
            for info in infolist:
                msword_binary.seek(info.header_offset + 26)
                filename_len, extrafield_len = struct.unpack(
                    "<2H", msword_binary.read(4)
                )
                msword_binary.seek(info.header_offset + 30 + filename_len)
                extrafield = msword_binary.read(extrafield_len)
                extrafield_hex_as_text = [f"{h:02x}" for h in extrafield]

                if not extrafield:
                    extras[info.filename] = [extrafield_len, "nil"]
                elif (
                    extrafield_len <= truncate_extra_field
                ):  # field size larger than truncate value
                    extras[info.filename] = [
                        extrafield_len,
                        f"0x{''.join(extrafield_hex_as_text)}",
                    ]
                else:
                    extras[info.filename] = [
                        extrafield_len,
                        f"0x{''.join(extrafield_hex_as_text[0:truncate_extra_field])}",
                    ]
                    # adds only the select # of characters as specified in the variable truncate_extra_field.
                    # This is so that we don't end up with hundreds of characters in a cell in Excel,
                    # as some extra fields can be several hundred values long.
                    # But so far, most are 0x00, with only the first few being values other than hex 0x00.

        return extras

    def __load_xml(self, xml_file):
        try:
            with self._zip.open(xml_file) as xmlFile:
                content = xmlFile.read()
        except KeyError:  # not present in the archive
            return ""
        if "comments.xml" in xml_file:
            self.has_comments = True
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
            "content_types_content": "[Content_Types].xml",
            "xml_rels_content": "word/_rels/document.xml.rels",
        }

        modified_time = None
        compression_types = {0: "Store (None)", 8: "DEFLATE"}
        for attrib in xml_map:
            setattr(self, attrib, "")
        path_to_attrib = {}
        for attrib, file_path in xml_map.items():
            path_to_attrib[file_path] = attrib
            path_to_attrib[file_path.replace("/", "\\")] = attrib

        try:
            zipref = self._zip
            zip_info = [
                info
                for info in zipref.infolist()
                if info.filename.endswith(".xml") or info.filename.endswith(".rels")
            ]
            for xml in zip_info:
                md5hash = None
                xml_name = xml.filename
                xml_files[xml_name] = blank.copy()
                if (
                    "customXml/item" in xml_name
                    and "Props" not in xml_name
                    and xml_name not in self.item_files
                ):
                    self.item_files.append(xml_name)
                if "ink/ink" in xml_name and xml_name not in self.ink_files:
                    self.ink_files.append(xml_name)
                target_attrib = path_to_attrib.get(xml_name)
                if self.hashing or target_attrib:
                    try:
                        with zipref.open(xml_name) as xml_file:
                            content = xml_file.read()
                    except Exception:
                        content = None
                    if self.hashing and content is not None:
                        md5hash = self.hash(content)
                    if target_attrib:
                        if "comments.xml" in xml_name:
                            self.has_comments = True
                        setattr(
                            self, target_attrib, content if content is not None else ""
                        )
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
                if xml_name in self.extra_fields:
                    xml_files[xml_name]["Zip Extra Fields Length"] = self.extra_fields[
                        xml_name
                    ][0]
                    xml_files[xml_name]["Zip Extra Fields Bytes"] = self.extra_fields[
                        xml_name
                    ][1]
                else:
                    xml_name_modified = xml_name.replace("/", "\\")
                    if xml_name_modified in self.extra_fields:
                        xml_files[xml_name]["Zip Extra Fields Length"] = (
                            self.extra_fields[xml_name_modified][0]
                        )
                        xml_files[xml_name]["Zip Extra Fields Bytes"] = (
                            self.extra_fields[xml_name_modified][1]
                        )
                    else:
                        xml_files[xml_name]["Zip Extra Fields Length"] = 0
                        xml_files[xml_name]["Zip Extra Fields Bytes"] = "nil"
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
            content = self.__parse(xmlcontent)
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

    def get_content_types(self):
        entries = []
        if self.content_types_content:
            x = ET.fromstring(self.content_types_content)
            for node in x.findall("Types:Default", self.namespaces):
                entries.append(
                    ("Default", node.get("Extension"), node.get("ContentType"))
                )
            for node in x.findall("Types:Override", self.namespaces):
                entries.append(
                    ("Override", node.get("PartName"), node.get("ContentType"))
                )
        return entries

    def get_xml_rels(self):
        rels = {}
        if self.xml_rels_content:
            x = ET.fromstring(self.xml_rels_content)
            rels = {
                node.get("Target"): [node.get("Id"), node.get("Type").split("/")[-1]]
                for node in x.findall("Relationships:Relationship", self.namespaces)
            }
        return rels

    def __extract_all_rsids_from_settings_xml(self):
        """
        function to extract all RSIDs at the beginning of the class.
        :return:
        """
        rsids = []
        x = self.__parse(self.settings_xml_content)
        rsid_tags = x.findall(".//w:rsid", self.namespaces)
        for tag in rsid_tags:
            rsid_tag = tag.get(f"{{{self.namespaces['w']}}}val", None)
            if rsid_tag:
                rsids.append(rsid_tag)
        return "" if not rsids else rsids

    def __count_rsids(self):
        """
        Single pass over the p, r, t and tr tags in document.xml counting each rsid attribute
        (rsidR, rsidRPr, rsidP, rsidRDefault, rsidTr, paraId, textId).

        :return: {attribute name: Counter of rsid value -> occurrences in document.xml}
        """
        w = self.namespaces["w"]
        w14 = self.namespaces["w14"]
        keys = {
            "rsidR": f"{{{w}}}rsidR",
            "rsidRDefault": f"{{{w}}}rsidRDefault",
            "rsidRPr": f"{{{w}}}rsidRPr",
            "rsidP": f"{{{w}}}rsidP",
            "rsidTr": f"{{{w}}}rsidTr",
            "paraId": f"{{{w14}}}paraId",
            "textId": f"{{{w14}}}textId",
        }
        counts = {name: Counter() for name in keys}
        for entry in (self.p_tags, self.r_tags, self.t_tags, self.tr_tags):
            for item in entry:
                attrib = item.attrib
                if not attrib:
                    continue
                for name, key in keys.items():
                    value = attrib.get(key)
                    if value:
                        counts[name][value] += 1
        return counts

    def hyperlinks(self):
        """
        :return: Hyperlink values in document.xml
        """
        all_hyperlinks = []
        doc = self.__parse(self.document_xml_content)
        for hyperlink in doc.findall(f".//{{{self.namespaces['w']}}}hyperlink"):
            link_text = hyperlink.findall(f".//{{{self.namespaces['w']}}}t")
            hyperlinks = ",".join(link.text for link in link_text if link.text)
            hyperlinks = hyperlinks.replace("http", "hxxp")
            rel_id = hyperlink.get(f"{{{self.namespaces['r']}}}id", None)
            if rel_id:
                all_hyperlinks.append([hyperlinks, rel_id])
        rels = self.get_xml_rels()
        for k, v in rels.items():
            if v[1] == "hyperlink":
                all_hyperlinks.append([k.replace("http", "hxxp"), v[0]])
        formatted_hyperlinks = " | ".join(
            f"{url}: {rel}" for url, rel in all_hyperlinks
        )
        return formatted_hyperlinks

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
                with open(self.msword_file, "rb") as msword_binary:
                    for chunk in iter(lambda: msword_binary.read(1024 * 1024), b""):
                        filehash.update(chunk)
            else:
                filehash.update(content)
            return filehash.hexdigest().upper()
        return None  # if no hashing was selected.

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
        x = self.__parse(self.settings_xml_content)
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
        x = self.__parse(self.settings_xml_content)
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
        xml = self.__parse(self.settings_xml_content)
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

    def get_document_protection(self):
        enabled = False
        xml = self.__parse(self.settings_xml_content)
        protect_state = xml.find(f"{{{self.namespaces['w']}}}documentProtection")
        if protect_state is not None:
            for k, v in protect_state.attrib.items():
                k = k.replace(f"{{{self.namespaces['w']}}}", "")
                if k == "enforcement" and v.lower() in ("1", "true", "on"):
                    self.protection_state["enabled"] = True
                    enabled = True
                self.protection_state[k] = v
        return enabled

    def get_track_changes_status(self):
        """
        Status of Track Changes setting.
        """
        xml = self.__parse(self.settings_xml_content)
        element = xml.find(f"{{{self.namespaces['w']}}}trackChanges")
        if element is None:
            return False
        val = element.get(f"{{{self.namespaces['w']}}}val")
        return True if val is None else val.lower() in ("1", "true", "on")

    def get_track_changes(self):
        """
        Extracts tracked changes revisions from document.xml
        """
        w = self.namespaces["w"]

        def qn(tag):
            return f"{{{w}}}{tag}"

        def text_of(element, text_tag):
            return "".join(t.text or "" for t in element.findall(f".//{qn(text_tag)}"))

        def attrs(element):
            return element.get(qn("author")), element.get(qn("date"))

        if not self.document_xml_content:
            return []
        doc = self.__parse(self.document_xml_content)
        changes = []

        # Insertions, deletions, and moved text
        for tag, label, text_tag in (
            ("ins", "Insertion", "t"),
            ("del", "Deletion", "delText"),
            ("moveTo", "Move To", "t"),
            ("moveFrom", "Move From", "delText"),
        ):
            for element in doc.findall(f".//{qn(tag)}"):
                author, date = attrs(element)
                changes.append((label, author, date, text_of(element, text_tag)))
        for run in doc.findall(f".//{qn('r')}"):
            rpr_change = run.find(f"{qn('rPr')}/{qn('rPrChange')}")
            if rpr_change is not None:
                author, date = attrs(rpr_change)
                changes.append(("Formatting", author, date, text_of(run, "t")))
        for para in doc.findall(f".//{qn('p')}"):
            ppr_change = para.find(f"{qn('pPr')}/{qn('pPrChange')}")
            if ppr_change is not None:
                author, date = attrs(ppr_change)
                changes.append(
                    ("Paragraph Formatting", author, date, text_of(para, "t"))
                )
            mark_rpr = para.find(f"{qn('pPr')}/{qn('rPr')}")
            if mark_rpr is not None:
                for tag, label in (
                    ("ins", "Paragraph Mark Inserted"),
                    ("del", "Paragraph Mark Deleted"),
                ):
                    marker = mark_rpr.find(qn(tag))
                    if marker is not None:
                        author, date = attrs(marker)
                        changes.append((label, author, date, text_of(para, "t")))
                mark_rpr_change = mark_rpr.find(qn("rPrChange"))
                if mark_rpr_change is not None:
                    author, date = attrs(mark_rpr_change)
                    changes.append(
                        ("Paragraph Mark Formatting", author, date, text_of(para, "t"))
                    )
        for table in doc.findall(f".//{qn('tbl')}"):
            tblpr_change = table.find(f"{qn('tblPr')}/{qn('tblPrChange')}")
            if tblpr_change is not None:
                author, date = attrs(tblpr_change)
                changes.append(("Table Formatting", author, date, None))
        for row in doc.findall(f".//{qn('tr')}"):
            trpr_change = row.find(f"{qn('trPr')}/{qn('trPrChange')}")
            if trpr_change is not None:
                author, date = attrs(trpr_change)
                changes.append(("Table Row Formatting", author, date, None))
        for cell in doc.findall(f".//{qn('tc')}"):
            tcpr = cell.find(qn("tcPr"))
            if tcpr is None:
                continue
            tcpr_change = tcpr.find(qn("tcPrChange"))
            if tcpr_change is not None:
                author, date = attrs(tcpr_change)
                changes.append(
                    ("Table Cell Formatting", author, date, text_of(cell, "t"))
                )
            for tag, label in (
                ("cellIns", "Table Cell Inserted"),
                ("cellDel", "Table Cell Deleted"),
            ):
                marker = tcpr.find(qn(tag))
                if marker is not None:
                    author, date = attrs(marker)
                    changes.append((label, author, date, text_of(cell, "t")))
        for sectpr_change in doc.findall(f".//{qn('sectPr')}/{qn('sectPrChange')}"):
            author, date = attrs(sectpr_change)
            changes.append(("Section Formatting", author, date, None))

        return changes

    def get_range_permissions(self):
        """
        Extracts w:permStart/w:permEnd permissions from document.xml.
        """
        w = self.namespaces["w"]

        def qn(tag):
            return f"{{{w}}}{tag}"

        if not self.document_xml_content:
            return []
        doc = self.__parse(self.document_xml_content)

        editors = {}
        texts = {}
        order = []
        open_ids = []

        for element in doc.iter():
            tag = element.tag
            if tag == qn("permStart"):
                pid = element.get(qn("id"))
                editors[pid] = element.get(qn("ed")) or element.get(qn("edGrp"))
                texts[pid] = []
                order.append(pid)
                open_ids.append(pid)
            elif tag == qn("permEnd"):
                pid = element.get(qn("id"))
                if pid in open_ids:
                    open_ids.remove(pid)
            elif tag == qn("p"):
                for pid in open_ids:
                    if texts[pid]:
                        texts[pid].append("\n")
            elif tag == qn("t"):
                for pid in open_ids:
                    texts[pid].append(element.text or "")

        return [(pid, editors[pid], "".join(texts[pid])) for pid in order]

    def get_ink(self):
        ts_data = []
        for ink_file in self.ink_files:
            load_ink = self.__load_xml(ink_file)
            xml = ET.fromstring(load_ink)
            ts = None
            for element in xml.iter():
                tag = element.tag.split("}")[-1] if "}" in element.tag else element.tag
                if tag == "timestamp":
                    for name, value in element.attrib.items():
                        attr_name = name.split("}")[-1] if "}" in name else name
                        if attr_name == "timeString":
                            ts = value
                            break
            ts_data.append([ink_file, ts])
        return ts_data

    def get_irm_info(self):
        """
        Returns the IRM info extracted from the OLE container.
        """
        return self.irm_info

    def adjust_timestamp(self, ts):
        if ts:
            adjusted_timestamp = ts.replace("T", " ").replace("Z", "").split(".")[0]
            if dt.fromisoformat(adjusted_timestamp).year < 2007:
                return None
            return adjusted_timestamp
        return None
