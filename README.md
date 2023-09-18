<h1>MS-Word-Parser</h1>
<h6>
This script will prompt you for a DOCx file and parse data from it and dump it to an Excel file.

The script does not validate that the file being passed to it is indeed a DOCx. It's up to the user to make sure he/she passes a DOCx.

The script will do the following processing:

1 - It will extract a list of all the files within the zip file and save it to a worksheet called XML_files.
    In this worksheet, it will save the following information to a row:
    "File Name", "XML", "Size (bytes)", "MD5Hash"

2 - It will extract all the unique RSIDs from the file word/settings.xml and write it to a worksheet
    called doc_summary.
    In this worksheet, it will save the following information to a row:
    "File Name", "Unique RSIDs", "RSID Root", "<w:p> tags", "<w:r> tags", "<w:t> tags"
    Where "Unique RSID" is a numerical count of the # of RSIDs in the file.

    What is an RSID (Revision Save ID)?
    See https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.rsid?view=openxml-2.8.1

3 - It will extract all the unique RSIDs from the file word/settings.xml and write it to a worksheet called rsids.
    In this worksheet, it will save the following information to rows (one for each unique RSID):
    "File Name", "RSID"

3 - It will extract all known relevant metadata from the files docProps/app.xml and docProps/core.xml
    and write it to a worksheet called metadata.
    In this worksheet, it will save the following information to a row:
    "File Name", "Author", "Created Date","Last Modified By","Modified Date","Last Printed Date","Manager","Company",
    "Revision","Total Editing Time","Pages","Paragraphs","Lines","Words","Characters","Characters With Spaces",
    "Title","Subject","Keywords","Description","Application","App Version","Template","Doc Security","Category",
    "contentStatus"</h6>


<h2>********** Dependencies **********</h2>

<h6>If running the script on a Linux system, you may need to install python-tk. You can do this with the following
command on a Debian (e.g. Ubuntu) system from the terminal window:
sudo apt-get install python3-tk
<br>
Whether running on Linux, Mac, or Windows, you may need to install some of the libraries if they are not included in
your installation of Python 3.
<br>
In particular, you may need to install openpyxl and hashlib.  
    
You can do so as follows from a terminal window:
<hr>
    pip3 install openpyxl<br>
    pip3 install hashlib  
<hr>
If any other libraries are missing when trying to execute the script, install those in the same manner.</h6>
