from datetime import datetime as dt


class DataStore:
    """Stores the state of all variables for use in multiple functions."""

    def __init__(self):
        """Main data stores"""
        self.doc_summary_worksheet = {}
        self.metadata_worksheet = {}
        self.archive_files_worksheet = {}
        self.rsids_worksheet = {}
        self.comments_worksheet = {}
        self.people_worksheet = {}
        self.extensible_worksheet = {}
        self.extended_worksheet = {}
        self.comments_ids_worksheet = {}
        self.custom_xml_worksheet = {}
        self.item_worksheet = {}
        self.ink_worksheet = {}
        self.timeline_worksheet = {}
        self.aggregated_worksheet = {}
        self.ink_content = []
        self.filenames = []
        self.item_xml_content = None
        self.excel_file = None
        self.sqlite_file = None
        self.output_path = None
        self.errors_worksheet = {"File Name": [], "Error": []}
        self.timestamp = dt.now().strftime("%Y%m%d_%H%M%S")
        self.basename = f"ms-word-parser-{self.timestamp}"
        self.log_file = f"ms-word-parser-{self.timestamp}.log"
        self.ms_word_gui = None
        self.start_time = None
        self.color_fmt = None
        self.logger = None
        self.sqlite = False
        self.excel = False
        self.timeline = False
        self.type_map = {
            "<w:p> tags": "Int32",
            "<w:r> tags": "Int32",
            "<w:t> tags": "Int32",
            "<w:tr> tags": "Int32",
            "<w14:docId>": "string",
            "<w15:docId>": "string",
            "<w16:docId>": "string",
            "App Version": "string",
            "Application": "string",
            "Archive File": "string",
            "Author": "string",
            "Category": "string",
            "Characters With Spaces": "Int32",
            "Characters": "Int32",
            "Comment ID": "Int32",
            "Comment paraId": "string",
            "Company": "string",
            "Content Status": "string",
            "Content": "string",
            "Count in document.xml": "Int32",
            "Created Date": "datetime64[ns]",
            "dateUtc": "datetime64[ns]",
            "Description": "string",
            "Doc Security": "Int32",
            "Done": "boolean",
            "durableId": "string",
            "File Created Date": "datetime64[ns]",
            "File Modified Date": "datetime64[ns]",
            "File Name": "string",
            "Grammar Check": "string",
            "Has Comments": "boolean",
            "Has Ink": "boolean",
            "Hyperlinks Changed": "string",
            "Hyperlinks": "string",
            "Initials": "string",
            "Ink XML File": "string",
            "Item XML File": "string",
            "Keywords": "string",
            "Language": "string",
            "Last Modified By": "string",
            "Last Printed Date": "datetime64[ns]",
            "Lines": "Int32",
            "Manager": "string",
            "MD5 Hash": "string",
            "Modified Date": "datetime64[ns]",
            "Modified Time (local/UTC/Redmond, Washington)": "datetime64[ns]",
            "Pages": "Int32",
            "Paragraphs": "Int32",
            "paraId Text": "string",
            "paraId": "string",
            "paraIdParent": "string",
            "providerId": "string",
            "reactionDateUtc": "datetime64[ns]",
            "reactionType": "string",
            "Revision": "Int32",
            "RSID Root": "string",
            "RSID Type": "string",
            "RSID Value": "string",
            "Shared Doc": "string",
            "Source": "string",
            "Spell Check": "string",
            "Subject": "string",
            "Template": "string",
            "Timestamp (UTC)": "datetime64[ns]",
            "Timestamp": "datetime64[ns]",
            "Title": "string",
            "Total Editing Time": "string",
            "Type": "string",
            "Uncompressed Size (bytes)": "Int32",
            "Unique rsidR": "Int32",
            "uri": "string",
            "userId": "string",
            "userName": "string",
            "userProvider": "string",
            "Value": "string",
            "Version": "string",
            "Words": "Int32",
            "ZIP Compression Type": "string",
            "ZIP Create System": "Int32",
            "ZIP Created Version": "Int32",
            "ZIP Extra Characters (truncated)": "string",
            "ZIP Extra Flag (len)": "Int32",
            "ZIP Extract Version": "Int32",
            "ZIP Flag Bits (hex)": "string",
        }
        self.sqlite_types = {
            "Int32": "INTEGER",
            "string": "TEXT",
            "boolean": "INTEGER",
            "datetime64[ns]": "TEXT",
        }
        self.triage_files = True
        self.hash_files = False
        self.total = 0
        self.remaining = 0
        self.done = 0

    def reset_vars(self):
        """Reset variables"""
        self.doc_summary_worksheet = {}
        self.metadata_worksheet = {}
        self.archive_files_worksheet = {}
        self.rsids_worksheet = {}
        self.comments_worksheet = {}
        self.people_worksheet = {}
        self.extensible_worksheet = {}
        self.extended_worksheet = {}
        self.comments_ids_worksheet = {}
        self.custom_xml_worksheet = {}
        self.item_worksheet = {}
        self.ink_worksheet = {}
        self.timeline_worksheet = {}
        self.aggregated_worksheet = {}
        self.ink_content = []
        self.item_xml_content = None
        self.excel_file = None
        self.sqlite_file = None
        self.output_path = None
        self.errors_worksheet = {"File Name": [], "Error": []}
        self.timestamp = dt.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = f"ms-word-parser-log-{self.timestamp}.log"
        self.basename = f"ms-word-parser-log-{self.timestamp}"
        self.sqlite = False
        self.excel = False
        self.timeline = False
        self.triage_files = True
        self.hash_files = False
        self.total = 0
        self.remaining = 0
        self.done = 0
