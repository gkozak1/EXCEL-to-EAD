#!/usr/bin/perl



use CGI ':standard';  
use CGI::Carp qw ( fatalsToBrowser );  


$directorypath = "/cul/web/collections.library.cornell.edu/htdocs/archivespace/upload/";
 

################################

my $files_location;  
my $ID;  
my @fileholder;  
 
$files_location = $directorypath;  
 
$ID = param('ID');  
 
if ($ID eq '') {  
print "Content-type: text/html\n\n";  
print "You must specify a file to download.";  
} else {  
 
open(DLFILE, "<$files_location/$ID") || Error('open', 'file');  
@fileholder = <DLFILE>;  
close (DLFILE) || Error ('close', 'file');  
 
 
#these are the html codes that forces the browser to open for download  
print "Content-Type:application/x-download\n";  
print "Content-Disposition:attachment;filename=$ID\n\n";  
print @fileholder;  
}
