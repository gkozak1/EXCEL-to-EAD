#!/usr/bin/perl -w
#|----------------------------------------------------------------|
#| Perl Script: getmarc.cgi                                       |
#|----------------------------------------------------------------|
#|  2017, George Kozak, gsk5@cornell.edu                          |
#|  Library Systems Cornell University                            |
#|----------------------------------------------------------------|
#| This job will take a Bib ID submitted by a user                |
#| and output a MARC XML file                                     |
#| Original Program created by Frances Webb (called marcsave.pl)  |
#|----------------------------------------------------------------|
use strict;
use warnings;
use CGI;
use LWP::Simple;

#|----------------------------------------------------------------|
#| variables passed from HTML file                                |
#|----------------------------------------------------------------|

my $query = new CGI;
my $bibid = $query->param("bibid");

#|----------------------------------------------------------------|
#| variables used internally in this program                      |
#|----------------------------------------------------------------|
my ($badid);
my ($destDir, $destfile, $url, $record_xml, $writeFile);

$badid = "http://collections.library.cornell.edu/archivespace/request_badid.html"; 
#|----------------------------------------------------------------|
#| if bibid missing, exit out                                     |
#|----------------------------------------------------------------|
if ( !$bibid ) {
    print "Location: $badid\n\n"; 
    exit(0);
    }

#|----------------------------------------------------------------|
#| process bibid to get Marc XML                                  |
#|----------------------------------------------------------------|
$destDir = '/cul/web/collections.library.cornell.edu/htdocs/archivespace/marc';
$destfile = $destDir . "/" . $bibid . ".xml";
$url = 'http://yazproxy.library.cornell.edu:9000/voyager?query=rec.id%3D'. $bibid. '&startRecord=1&maximumRecords=1&recordSchema=opacxml&version=1.1&operation=searchRetrieve';
$record_xml = get($url);
if ($record_xml =~ m,(<record .*</record>),s) {
    $record_xml = $1;
    } 
#|----------------------------------------------------------------|
#| if bibid is bad, exit out                                      |
#|----------------------------------------------------------------|
else {
    print "Location: $badid\n\n"; 
    exit(0);
    }

open $writeFile, '>', $destfile
	or die $!;
     print $writeFile $record_xml, "\n";
     close $writeFile;

#|-----------------------------------------|
#| Now build HMTL page                     |
#|-----------------------------------------| 
print qq{<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"};
print qq{    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">\n};
print qq{\n};
print qq{<html xmlns="http://www.w3.org/1999/xhtml">\n};
print qq{\n};
print qq{<head>\n};
print qq{	<title>Get MARC XML</title>\n};
	print qq{<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />\n};
print qq{	<meta http-equiv="Content-Language" content="en-us" />\n};
print qq{	<link rel="shortcut icon" href="http://collections.library.cornell.edu/archivespace/favicon.ico" type="image/x-icon" />\n};
print qq{	<link rel="stylesheet" type="text/css" media="screen" href="http://collections.library.cornell.edu/archivespace/styles/screen.css" />\n};
print qq{</head>\n};
print qq{\n};
print qq{\n};
print qq{<div id="skipnav">\n};
print qq{	<a href="#content">Skip to main content</a>\n};
print qq{</div>\n};
print qq{\n};
print qq{<hr />\n};
print qq{\n};
print qq{<div id="cu-identity">\n};
print qq{	<div id="cu-logo">\n};
print qq{		<a id="insignia-link" href="http://www.cornell.edu/"><img src="http://collections.library.cornell.edu/archivespace/images/Library_2line_4c.gif" alt="Cornell University Library" width="283" height="75" border="0" /></a>\n};
print qq{		<div id="unit-signature-links">\n};
print qq{			<a id="cornell-link" href="http://www.cornell.edu/">Cornell University</a>\n};
print qq{			<a id="unit-link" href="http://www.library.cornell.edu/">Cornell University Library</a>\n};
print qq{		</div>\n};
print qq{	</div>\n};
print qq{\n};
print qq{	<div id="search-navigation">\n};
print qq{		<ul>\n};
print qq{			<li><a href="http://www.library.cornell.edu/services/searches.html">Search Library</a></li>\n};
print qq{			<li><a href="http://www.cornell.edu/search/">Search Cornell</a></li>\n};
print qq{		</ul>\n};
print qq{	</div>\n};
print qq{\n};
print qq{</div>\n};
print qq{\n};
print qq{\n};
print qq{<hr />\n};
print qq{\n};
print qq{<div id="wrap">\n};
print qq{\n};
print qq{<div id="content">\n};
print qq{  <div id="main">\n};
print qq{    <h1>Get MARC XML</h1>\n};
print qq{    <h2>Conversion Complete!</h2>\n};
print qq{<p> Download MARC XML file from <a href="http://collections.library.cornell.edu/archivespace/marc/$bibid.xml">HERE</a>\n};
print qq{</p>\n};
print qq{        </div>\n};
print qq{    </div>\n};
print qq{\n};
print qq{</div>\n};
print qq{<hr />\n};
print qq{<div id="footer">\n};
print qq{    <div id="footer-content">\n};
print qq{        <p>&copy;2017 <a href="http://www.library.cornell.edu/">Cornell University Library</a></p>\n};
print qq{    </div>\n};
print qq{</div>\n};
print qq{\n};
print qq{\n};
print qq{</body>\n};
print qq{</html>\n};
