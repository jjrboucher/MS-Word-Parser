import re


def extract_rsids_from_xml(xmlcontent):
    """
    function to extract rsids and rsidRoot from settings.xml
    :param xmlcontent:
    :return: a list containing all rsids, and rsidRoot
    """
    try:
        all_rsids = []
        pattern = r'<w:rsid w:val="[^>]*/>'
        matches = re.findall(pattern, xmlcontent)  # Find all RSIDs, not rsidRoot. rsidRoot is repeated in rsids

        print("Processing word/settings.xml for RSIDs")
        for match in matches:  # loops through all matches
            # greps for rsid using a group to extract the actual RSID from the string.
            rsid_match = re.search(r'<w:rsid w:val="([^"]*)"', match)
            if rsid_match:
                all_rsids.append(rsid_match.group(1))  # Appends it to the list

        print("Processing word/settings.xml for rsidRoot.")
        rsid_root = re.search(r'<w:rsidRoot w:val="([^"]*)"', xmlcontent)

        if rsid_root is None:
            rsid_root = ""
        else:
            rsid_root = rsid_root.group(1)

        return all_rsids, rsid_root

    except Exception as function_error:
        print(f"An error occurred while extracting RSIDs: {function_error}")
        return []  # if it can't find any RSID (that should never happen), it returns an empty list.
