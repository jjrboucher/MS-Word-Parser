# MS Word Parser

## Overview
#### Purpose
Conduct a forensic analysis of one or more Microsoft Word docx/dotx/dotm files.

#### Installation
MS Word Parser is available for installation via pip as `ms-word-parser`:
```
python3 -m pip install ms-word-parser
```
All requirements will automatically install when using this method.

#### Usage
To use the command-line version, simply run:
```
$ parse-docx

usage: parse-docx [-h] [-e EXCEL] [-g] [--hash] [-r] [-V] [--dir DIR | --files [FILES ...]] [-t | -f]

MS Word Parser 2.0.0

options:
  -h, --help            show this help message and exit
  -e EXCEL, --excel EXCEL
                        output path and filename for the Excel output
  -g, --gui             launch the gui
  --hash                hash the doc zip contents
  -r, --recurse         recursively process files in directory
  -V, --verbose         Output to STDOUT as well as log
  --dir DIR             directory to process
  --files [FILES ...]   individual files to be processed
  -t, --triage          triage mode
  -f, --full            full mode
```

To use the Graphical User Interface (GUI), simply run:
`parse-docx -g`

![v2 0 0-screenshot](https://github.com/jjrboucher/MS-Word-Parser/blob/master/.assets/v2.0.0-screenshot.png)

### Input
You can select one or more MS Word files within a folder or alternatively select a root folder and the script will recursively add all MS Word files it finds from that point in the hierarchy.

### Output

The results will be saved to a Microsoft Excel file. You will be prompted to provide a file name of the Excel file and where you want to save it. The Excel file will have four or more worksheets depending on the processing option you select.
The script will also output a log file in the same folder as the Excel file and will bear the following naming convention: `DOCx_Parser_Log_YYYYMMDD_HHMMSS.log.`

### Processing Options

#### Triage
Triage mode will produce the Doc_Summary, Metadata, Comments, and Excel Tips worksheets only. If you are examining 10K, 20K or more MS Word documents collected as part of your investigation, you are going to want to start with triage mode. This will run faster, as there is less parsing. It does not produce the RSIDs or Archive Files worksheets, which can account for a lot of data.
Conduct your initial review using this triage mode and identify documents that you want to reprocess with the full parsing option.

#### Full
The Full parsing option will produce all of the worksheets covered earlier. If processing a large number of files, it can result in a large Excel document. In a test case with 17,500 files, it produced 25 worksheets of RSID values, and an Excel document that was ½ Gig in size. For this reason, it is recommended that you start with triage and only use the Full parsing on select files identified as a result of the triage exercise.

#### Hash Files
If you select this option, the script will hash each file it processes, as well as hash each file within the DOCx compound file (found in the worksheet “Archive Files”).  You may want to use this option to capture the MD5 hash of each file to attest to the integrity of the file later on, as well as to use to identify duplicate files.

#### GUI Workflow
When you launch the application in GUI mode, you’ll need to select your parsing option (Triage or Full), and if you wish, check off the option to hash the files.
You will need to select the Excel file where you wish to save the processing results, as well as select either a list of files within a single folder, or select a folder that you wish to process recursively.
The Processing Status window will identify how many files were passed to the script for parsing, the number of files that produced a processing error, and the # of files remaining to be processed. Within this window you will see output from the script as it’s processing the files.
#### Stop
If you click on the stop button, it will stop processing. The Excel file will written up to the point of the last file to be processed. The log file will also be written.

#### Reset
If you want to run the script again, you can click on the Reset button and select a new output file and new files to process.

#### Open Log File
Opens the log file.

#### Open Excel File
Opens the Excel file.

#### Open Output Path
Open the output path folder.

## Microsoft Excel File

### Worksheets

#### Doc_Summary
This worksheet will have one row for each file processed. It will contain a summary of the artifacts relating to the documents (e.g., MD5 hash, number of rsidR, RSID Root, number of paragraph tags (w:p), run tags (w:r), text tags (w:t).

#### Metadata
This worksheet will have one row for each file processed. It will contain metadata such as the author of the document, the date created, who last modified it, the last modified date, revision count, editing time, etc. It’s important to note that there have been instances where the metadata relating to the number of pages/lines/characters have been inaccurate. This is not a flaw in the script. Rather, it was found that the metadata within the document was wrong. To correct this, the document had to be opened and resaved. But in doing so, you are changing the last time it was modified, and by whom.

#### Comments
This worksheet will have zero or more rows for each file. If there are any comments in the document, each comment will occupy one row. Any reply to a comment will be in its own row. It will contain the comment ID number, the timestamp of the comment, the author, the author’s initials, and the comment itself.

#### RSIDs
This worksheet will have multiple rows for each file. Each row will be a unique RSID Type and value, as well as how many times that RSID appears in the document. The rsidR value denotes Revision Save Identifier. Each time there is a save action, a new rsidR value is generated and any text from that session will have that rsidR value attached to it. This allows you to identify what text was typed within a given editing session (a given rsidR value). 

Large documents can result in hundreds of rows, as they will contain many different RSID values. Excel limits the number of rows in a spreadsheet to just over 1 million. The script will break up the RSIDs worksheet into multiple worksheets if necessary, with 1 million rows in each. Each RSIDs worksheet will include a number (1, 2, 3, etc.) in the worksheet name.

#### Archive Files
This worksheet will have multiple rows for each file. Microsoft Word docx files are compound files. In fact, they are nothing more than a ZIP compressed file with multiple files within it. This worksheet will extract the name of each file within the compound file and include its name, the modified time of the file within the compound file (which in most cases should be “nil”), the uncompressed size of the file, and various ZIP file attributes for that specific file within the compound file.

#### Excel Tips
This worksheet will contain analysis tips.
