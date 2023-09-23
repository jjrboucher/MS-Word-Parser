import re


def extract_tags_from_document_xml(xmlcontent):
    """
    extract relevant artifacts from document.xml
    :param xmlcontent:
    :return: a dictionary containing three key/value pairs.
    "paragraph": # of paragraph tags
    "runs": # of runs tags
    "text": # of text tags
    """

    print("Processing word/document.xml to count # of <w:p>, <w:r>, and <w:t> tags.")
    document_xml = {"paragraphs": len(re.findall(r'</w:p>', xmlcontent)),
                    "runs": len(re.findall(r'</w:r>', xmlcontent)),
                    "text": len(re.findall(r'</w:t>', xmlcontent))}
    return document_xml
