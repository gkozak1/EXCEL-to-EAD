# EXCEL-to-EAD
Software to allow Archivists to convert an EXCEL spreadsheet 
This feature involves a cgi script and HTML pages.  
The HTML pages will be in a directory called htdocs/archivespace and the cgi script is in a directory called cgi-bin.
The general process is as follows:
 1. on the web, a user somes to http://xxx.xx.xx/archivespace
 2. The index.html allows them to upload an EXCEL spreadsheet
 3. then the index.html prompts the user to enter the bib id for this EAD
 4. The cgi script "convert_excel.cgi" is called
 5. "convert_excel.cgi" does the following:
 6. checks the variables passed from index.html
 7. If EXCEL file missing then archivesapce/request_missing.html is called and program stops.
 8. If the filename is malformed then archivespace/request_badfile.html is called and program stops.
 9. If the extension is xlsx or xls then archivespace/request_badext.html is called and program stops.
 10. If the bibid is missing then archivespace/request_badid.html is called and program stops.
 11. If all of the above tests pass then we contact Voyager LMS system with bibid.
 12. If the bibid is not found in Voyager then archivespace/request_badid.html is called and program stops.
 13. Else, we read MARC XML file from Voyager and begin building EAD.
 14. Next, we read the EXCEL spreadsheet, and add to EAD.
 15. Finally, we create EAD file, and give user a prompt so they can download the EAD file.
  
