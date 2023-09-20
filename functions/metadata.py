import re


def core_xml(xmlcontent):
    """Function input: core.xml

        Function return: Returns a dictionary containing the metadata extracted by grep expressions"""

    # extract relevant metadata from core.xml file using a GREP expression
    core_xml_metadata = {"title": re.search(r'<dc:title>(.*?)</dc:title>', xmlcontent),
                         "subject": re.search(r'<dc:subject>(.*?)</dc:subject>', xmlcontent),
                         "creator": re.search(r'<dc:creator>(.*?)</dc:creator>', xmlcontent),
                         "keywords": re.search(r'<cp:keywords>(.*?)</cp:keywords>', xmlcontent),
                         "description": re.search(r'<dc:description>(.*?)</dc:description>', xmlcontent),
                         "revision": re.search(r'<cp:revision>(.*?)</cp:revision>', xmlcontent),
                         "created": re.search(r'<dcterms:created.*?>(.*?)</dcterms:created>', xmlcontent),
                         "modified": re.search(r'<dcterms:modified.*?>(.*?)</dcterms:modified>', xmlcontent),
                         "lastModifiedBy": re.search(r'<cp:lastModifiedBy>(.*?)</cp:lastModifiedBy>', xmlcontent),
                         "lastPrinted": re.search(r'<cp:lastPrinted>(.*?)</cp:lastPrinted>', xmlcontent),
                         "category": re.search(r'<cp:category>(.*?)</cp:category>', xmlcontent),
                         "contentStatus": re.search(r'<cp:contentStatus>(.*?)</cp:contentStatus>', xmlcontent)}

    print("Processing docProps/core.xml for metadata.")
    for key, value in core_xml_metadata.items():  # check the results of the GREP searches
        if value is None:  # if no hit, assign empty value
            core_xml_metadata[key] = ""
        else:  # if a hit, extract group(1) from the search hit and add to the dictionary
            core_xml_metadata[key] = core_xml_metadata[key].group(1)
    return core_xml_metadata  # returns dictionary


def app_xml(xmlcontent):
    """Function input: app.xml

        Function return: Returns a dictionary containing the metadata extracted by grep expressions"""

    # extract relevant metadata from app.xml file using a GREP expression
    app_xml_metadata = {"template": re.search(r'<Template>(.*?)</Template>', xmlcontent),
                        "totalTime": re.search(r'<TotalTime>(.*?)</TotalTime>', xmlcontent),
                        "pages": re.search(r'<Pages>(.*?)</Pages>', xmlcontent),
                        "words": re.search(r'<Words>(.*?)</Words>', xmlcontent),
                        "characters": re.search(r'<Characters>(.*?)</Characters>', xmlcontent),
                        "application": re.search(r'<Application>(.*?)</Application>', xmlcontent),
                        "docSecurity": re.search(r'<DocSecurity>(.*?)</DocSecurity>', xmlcontent),
                        "lines": re.search(r'<Lines>(.*?)</Lines>', xmlcontent),
                        "paragraphs": re.search(r'<Paragraphs>(.*?)</Paragraphs>', xmlcontent),
                        "charactersWithSpaces": re.search(r'<CharactersWithSpaces>(.*?)</CharactersWithSpaces>',
                                                          xmlcontent),
                        "appVersion": re.search(r'<AppVersion>(.*?)</AppVersion>', xmlcontent),
                        "manager": re.search(r'<Manager>(.*?)</Manager>', xmlcontent),
                        "company": re.search(r'<Company>(.*?)</Company>', xmlcontent)}

    print("Processing docProps/app.xml for metadata.")
    for key, value in app_xml_metadata.items():  # check the results of the GREP searches
        if value is None:  # if no hit, assign empty value
            app_xml_metadata[key] = ""
        else:  # if a hit, extract group(1) from the search hit and add to the dictionary
            app_xml_metadata[key] = app_xml_metadata[key].group(1)

    return app_xml_metadata  # returns dictionary
