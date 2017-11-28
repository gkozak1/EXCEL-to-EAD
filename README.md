# EXCEL-to-EAD
Software to allow Archivists to convert an EXCEL spreadsheet 
This feature involves a cgi script and HTML pages.  
The HTML pages will be in a directory called html and the cgi script is in a directory called cgi-bin.
The general process is as follows:
 1. someone comes to the web site and uploads an EXCEL spreadsheet
 2. they enter the bib id for this EAD
 3. The cgi script is called and using MARC XML information is pulled from the Voyager Database
 4. Then information is culled from the EXCEL spreadsheet and finally a XML file containing the EAD is created.
