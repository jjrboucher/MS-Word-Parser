<h1>MS-Word-Parser</h1>
<h6>
This script will prompt you for a DOCx file (or several DOCx files if you wish) and parse data from it and dump it to an Excel file. You can either give it a new file name so it creates a new Excel file, or point to an existing Excel file to have the script add to that Excel file. The one caveate is if running the Python version on a Mac, the author has observed that it won't append to an existing Excel file, rather it will overwrite it. But in Windows, it properly appends to an existing Excel file. This allows you to re-run this script against new documents, and add the results to an existing Excel file.

The script does not validate that the file(s) being passed to it is/are indeed DOCx. It's up to the user to make sure he/she passes a DOCx file(s).

The script will output data to the command window and uses ASCII escape sequences to add colour to the text. Unfortunately, with the EXE it doesn't work. You will see the escape sequences as text which makes the output look messy. Don't worry, the script is still working as expected.

<h4>Triage Mode</h4>
When you first execute the script, it will prompt you to run it in either triage mode or full parsing. Triage mode will only create the first two worksheets (doc_summary and RSIDs). The other two worksheets are more technical, and can generate a lot of data (which means longer processing time) that you may not be interested in unless doing a deep dive.

<h4>Log File</h4>
The script will also create a log file in the same folder as the DOCx file(s) you select to process. The log file is called DOCx_Parser_Log_yyyymmdd_hhmmss.log where yyyymmdd is the date and hhmmss is the time expressed in the computer's local time.

<h4>Excel File</h4>
The script will do the following processing:

1 - It will extract all the unique RSIDs from the file word/settings.xml and write it to a worksheet
    called doc_summary.
    In this worksheet, it will save the following information to a row:
    "File Name", "MD5 Hash", "Unique RSIDs", "RSID Root", "<w:p> tags", "<w:r> tags", "<w:t> tags"
    Where "Unique RSID" is a numerical count of the # of RSIDs settings.xml. The count for the <w:?> tags are count of those tags in document.xml file.<br>
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
    "File Name", "Archive File", "MD5 Hash", "Modified Time (local/UTC/Redmond, Washington)", "Size (bytes)", "ZIP Compression Type", "ZIP Create System", "ZIP Created Version", "ZIP Extract Version", "ZIP Flag Bits (hex)", "ZIP Extra Flag (len)", "ZIP Extra Characters"<br><br>
    **NOTE:** The modified time of a file inside of a compound file will be local time to the system that edited it. If you know
    what system edited it, you can get the time zone from that system. Otherwise, it's not possible to know what time zone that date/time is expressed in.<br>

The columns with information about the files in the ZIP is based on the fact that each file in a ZIP has it's own header (https://en.wikipedia.org/wiki/ZIP_(file_format)#Local_file_header). Most of these values are decoded by the library "zipfile". But the "ZIP Extra Characters" is not extracted by the library. The script manually parses the header to extract this info, and displays the content truncated to 20 values. The column "ZIP Extra Flag (len)" lets you know how many characters are actually in that extra field. Observations to date is that only the first 10 or so characters have a value. The rest has been observed to be 0x00.

If the modified time is blank, it will show "nil" for value. Otherwise, it shows the date/time that it was modfiied.
    This should always be a nil value. The only time the author has seen an actual date is when the DOCX was renamed to ZIP,
    opened with WinZip and an XML file was edited within the zip and saved (and ZIP resaved). This results in that XML file
    now having a modified date of when the XML file in the ZIP was modified. This is not normal, and should serve as a red
    flag that someone may have manually edited the content of the ZIP file(s) that have a date/time. A caveat is that some applications other than MS Word (e.g., export from Pages to Docx) will result in each embedded file bearing the date of the export.
    
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

<h2>Caveat</h2>
I have seen where the # of pages and paragraphs was not correct. This was not a problem with the script, but rather with MS Word. When looking at the metadata via the embedded XML files, they were also not accurate (with those inaccurate values accurately extracted by the script). This inaccuracy was also observed in the Details tab when looking at the properties of the document within Windows Explorer. It was only after opening the document and re-saving it did the count update to the correct values. But of course, this modifies the document which is not desireable. Because of this, it is a good practice to make a copy of the document via Windows Explorer copy/paste, and open that one to see if what is being reported is accurate. <br><br>

This doesn't happen often, but it does happen, so it's important to be aware of this.

<h2>Dependencies</h2>

<h6>If running the script on a Linux system, you may need to install python-tk. You can do this with the following
command on a Debian (e.g. Ubuntu) system from the terminal window:<br>  
    
    sudo apt-get install python3-tk
<br>
Whether running on Linux, Mac, or Windows, you may need to install some of the libraries if they are not included in
your installation of Python 3.
<br>
In particular, you may need to install openpyxl, colorama and hashlib.  
    
<br>You can do so as follows from a terminal window while in the folder with the script and requirements.txt file:

    pip3 install -r requirements.txt
<hr>
If any other libraries are missing when trying to execute the script, install those in the same manner.</h6>

<h2>Executable Version</h2>
If you'd rather run the executable rather than needing Python, grab the .exe file.<br>
<br>
You will get the best experience if you run the executable by opening a command/terminal window in Windows, and executing it from there. If you simply double click on the .exe from Windows File Explorer, it will work fine. But the command window closes as soon as the script ends so you don't get a chance to see the information that the script outputs to the screen, and the coloured text will not work so you'll see ANSI escape sequences. If you run it in the command/terminal window, you can scroll through the output, including seeing the filename and path for the Excel file and log file.
