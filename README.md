<h1>MS-Word-Parser</h1>
<h6>
This script will prompt you for a DOCx file and parse data from it and dump it to an Excel file.

The script does not validate that the file being passed to it is indeed a DOCx. It's up to the user to make sure he/she passes a DOCx.

The script will do the following processing:

1 - It will extract all the unique RSIDs from the file word/settings.xml and write it to a worksheet
    called doc_summary.
    In this worksheet, it will save the following information to a row:
    "File Name", "Unique RSIDs", "RSID Root", "<w:p> tags", "<w:r> tags", "<w:t> tags"
    Where "Unique RSID" is a numerical count of the # of RSIDs in the document.xml file.<br>
    <br>What is an RSID (Revision Save ID)?<br>
    See https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.rsid?view=openxml-2.8.1

2 - It will extract all known relevant metadata from the files docProps/app.xml and docProps/core.xml
    and write it to a worksheet called metadata.
    In this worksheet, it will save the following information to a row:<br><br>
    "File Name", "Author", "Created Date","Last Modified By","Modified Date","Last Printed Date","Manager","Company",
    "Revision","Total Editing Time","Pages","Paragraphs","Lines","Words","Characters","Characters With Spaces",
    "Title","Subject","Keywords","Description","Application","App Version","Template","Doc Security","Category",
    "contentStatus"
    
3 - It will extract a list of all the files within the zip file and save it to a worksheet called XML_files.
    In this worksheet, it will save the following information to a row:<br><br>
    "File Name", "XML", "Modified Time (local)", "Size (bytes)", "MD5Hash"<br><br>
    **NOTE:** The modified time of an XML inside of a compound file will be local time to the system that edited it. If you know
    what system edited it, you can get the time zone from that system. Otherwise, it's not possible to know what time zone that date is expressed in.<br><br>
    time is expressed in without any additional information to establish that.<br>
    If the modified time is blank, it will show "nil" for value. Otherwise, it shows the date/time (UTC) that it was modfiied.
    This should always be a nil value. The only time the author has seen an actual date is when the DOCX was renamed to ZIP,
    opened with WinZip and an XML file was edited within the zip and saved (and ZIP resaved). This results in that XML file
    now having a modified date of when the XML file in the ZIP was modified. This is not normal, and should serve as a red
    flag that someone may have manually edited the content of the ZIP file(s) with a date/time.
    
4 - It will extract all the unique RSIDs from the file word/settings.xml and write it to a worksheet called rsids.
    In this worksheet, it will save the following information to rows (one for each unique RSID):
    "File Name", "RSID Type", "RSID Value", "Count in document.xml"<br><br>
    where RSID Type can be one of the following:<br><br>
    - rsidR<br>
    - rsidRPr<br>
    - rsidrP<br>
    - rsidRDefault<br>
    - paraID<br>
    - textID<br><br>
    And "Count in document.xml" is as the name implies, it's how many times that RSID is present in document.xml.</h6>

<h2>Dependencies</h2>

<h6>If running the script on a Linux system, you may need to install python-tk. You can do this with the following
command on a Debian (e.g. Ubuntu) system from the terminal window:<br>  
    
    sudo apt-get install python3-tk
<br>
Whether running on Linux, Mac, or Windows, you may need to install some of the libraries if they are not included in
your installation of Python 3.
<br>
In particular, you may need to install openpyxl and hashlib.  
    
<br>You can do so as follows from a terminal window while in the folder with the script and requirements.txt file:

    pip3 install -r requirements.txt
<hr>
If any other libraries are missing when trying to execute the script, install those in the same manner.</h6>
