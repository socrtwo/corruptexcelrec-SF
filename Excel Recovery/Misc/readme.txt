coffec v2.0.1 -----   (03 Sep 2009)

Corrupt Office 2007 Extractor Command-line version. To extract text from docx or pptx document. It also can do xlsx to csv.



This package contains following files

1.) readme.txt  (this file)
2.) coffec.exe



Extract all the files to a folder. Run coffec.exe in command-line mode.
The program needs ZipDll.dll and UnzDll.dll files to run successfully. Please copy those files to same folder!



See below:

Usage: coffec [-thx] "file"
Commands:
  -h    help message
  -t    output text or xlsx to csv
  -x    extract xml files from docx



For example:

coffec -t "intro.docx"

This will extract xml files and show you text content of document.xml through STDOUT



What's new

1.) New program name
2.) Extract text from pptx or xlsx to csv



Revision History

version 1.0.2 -----   (11 Jun 2009)

1.) Bug fixed. Adding one line only when continuous Tab found
2.) Bug fixed. Remove unwanted text content that parsing from token <w:instrText>
3.) New icon of executable
4.) Minor change.


version 1.0.1  ----  ( 30 Mar 2009 )

1.) First Built 






