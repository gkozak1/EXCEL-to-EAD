#!/usr/bin/perl -w
#|-----------------------------------------|
#| Perl Script: convert_excel.cgi          |
#|-----------------------------------------|
#|  2017, George Kozak, gsk5@cornell.edu   |
#|  Library Systems Cornell University     |
#|-----------------------------------------|
#| This job will take an EXCEL spreadsheet |
#|  submitted by a user                    |
#| and convert it into a EAD XML file      |
#|-----------------------------------------|
#| version dated 12/08/17                  |
#|-----------------------------------------|
use strict;
use warnings;
use CGI;
use CGI::Carp qw ( fatalsToBrowser );
use File::Basename;
use LWP::Simple;
use Spreadsheet::ParseXLSX;
use Roman;

#|----------------------------------|
#| variables passed from HTML file  |
#|----------------------------------|
$CGI::POST_MAX = 1024 * 5000;
my $safe_filename_characters = "a-zA-Z0-9_.-";
my $dir_path = "/cul/web/collections.library.cornell.edu/htdocs/archivespace/";
my $web_path = "http://collections.library.cornell.edu/archivespace/";
my $upload_dir = $dir_path . "upload";

my $query = new CGI;
my $filename = $query->param("txt");
my $bibid = $query->param("bibid");

my $missing = $web_path . "request_missing.html"; 
my $badfile = $web_path . "request_badfile.html"; 
my $badext = $web_path . "request_badext.html"; 
my $badid = $web_path . "request_badid.html"; 

#|---------------------------------------|
#| is filename missing? if so, exit out  |
#|---------------------------------------|
if ( !$filename ) {
    print "Location: $missing\n\n"; 
    exit(0);
   }

#|----------------------------------------------------|
#| parse filename to see if it is structured properly |
#|----------------------------------------------------|
my ( $name, $path, $extension ) = fileparse ( $filename, '..*' );
$filename = $name . $extension;
$filename =~ tr/ /_/;
$filename =~ s/[^$safe_filename_characters]//g;

#|------------------------------------------------|
#| if filename not structured properly, exit out  |
#|------------------------------------------------|
if ( $filename =~ /^([$safe_filename_characters]+)$/ ) {
     $filename = $1;
    }
else {
    print "Location: $badfile\n\n"; 
    exit(0);
   }

#|-----------------------------|
#| if bibid missing, exit out  |
#|-----------------------------|
if ( !$bibid ) {
    print "Location: $badid\n\n"; 
    exit(0);
    }
#|-------------------------------|
#| process bibid to get Marc XML |
#|-------------------------------|
my $destfile = $upload_dir . "/" . $bibid . ".xml";
my $url = 'http://yazproxy.library.cornell.edu:9000/voyager?query=rec.id%3D'. $bibid. '&startRecord=1&maximumRecords=1&recordSchema=opacxml&version=1.1&operation=searchRetrieve';
my $record_xml = get($url);
if ($record_xml =~ m,(<record .*</record>),s) {
    $record_xml = $1;
    } 
#|---------------------------|
#| if bibid is bad, exit out |
#|---------------------------|
else {
    print "Location: $badid\n\n"; 
    exit(0);
    }
my $writeFile;
open $writeFile, '>', $destfile
	or die $!;
     print $writeFile $record_xml, "\n";
     close $writeFile;

#|--------------|
#| upload file  |
#|--------------|
my $upload_filehandle = $query->upload("txt");
open ( UPLOADFILE, ">$upload_dir/$filename" ) or die "$!";
binmode UPLOADFILE;
while ( <$upload_filehandle> ) {
        print UPLOADFILE;
       }

close UPLOADFILE;
#|-------------------------------------|
#| Read the MARC XML file to get data  |
#| for the EAD we are building         |
#|-------------------------------------|

#|-----------------------------------|
#| Marc tags that we are looking for |
#|-----------------------------------|
#| MARC tag 100 - Main Entry -       |
#|                Personal Name      |
#|-----------------------------------|
my $tag100 = 0;
my $t100a = ""; # Personal Name
my $t100c = ""; # Titles associated with a name
my $t100d = ""; # Dates associated with a name
my $t100q = ""; # Fuller form of name

#|-----------------------------------|
#| MARC tag 110 - Main Entry -       |
#|                Corporate Name     |
#|-----------------------------------|
my $tag110 = 0;
my $t110a = ""; # Corporate name
my $t110b = ""; # subordinate unit

#|-----------------------------------|
#| MARC tag 245 -  Title Statement   |
#|-----------------------------------|
my $tag245 = 0;
my $t245a = ""; # Title
my $t245b = ""; # Remainder of title 
my $t245f = ""; # Date
my $t245g = ""; # Bulk dates

#|-------------------------------------|
#| MARC tag 300 - Physical Description |
#|                (repeatable)         |
#|-------------------------------------|
my $tag300 = 0;
my $t300_cnt = 0;
my $t300_array = ""; 
my $t300_3 = ""; # Material specified
my $t300a = ""; # Extent
my $t300f = ""; # Unit

#|------------------------------------------|
#| MARC tag 351 - Organization of Materials |
#|------------------------------------------|
my $tag351 = 0;
my $t351a = ""; # Organization

#|------------------------------------------|
#| MARC tag 500 - General Note (repeatable) |
#|------------------------------------------|
my $tag500 = 0;
my $t500_cnt = 0;
my $t500_array = "";
my $t500a = ""; # General Note

#|--------------------------------------------|
#| MARC tag 506 - Restrictions on Access Note |
#|                (repeatable)                |
#|--------------------------------------------|
my $tag506 = 0;
my $t506_cnt = 0;
my $t506_array = "";
my $t506_3 = ""; # Terms governing access
my $t506a = ""; # Terms governing access

#|--------------------------------------------|
#| MARC tag 520 - Summary, etc. (repeatable)  |
#|--------------------------------------------|
my $tag520 = 0;
my $t520_cnt = 0;
my $t520_array = "";
my $t520a = ""; # Summary, etc.
my $t520b = ""; # Expansion of summary note

#|--------------------------------------------|
#| MARC tag 524 - Preferred Citation          |
#|--------------------------------------------|
my $tag524 = 0;
my $t524a = ""; # Preferred Citation of Described Materials Note

#|-----------------------------------------|
#| MARC tag 538 - Technical information    |
#|-----------------------------------------|
my $tag538 = 0;
my $t538_cnt = 0;
my $t538_array = "";
my $t538a = ""; # notes

#|-----------------------------------------|
#| MARC tag 540 - Terms Governing Use and  |
#|                Reproduction Note        |
#|-----------------------------------------|
my $tag540 = 0;
my $t540_cnt = 0;
my $t540_array = "";
my $t540a = ""; # Terms governing use and reproduction

#|-------------------------------------------|
#| MARC tag 544 - Location of Other Archival |
#|                Materials Note             |
#|-------------------------------------------|
my $tag544 = 0;
my $t544_cnt = 0;
my $t544_array = "";
my $t544a = ""; # Note

#|-------------------------------------------|
#| MARC tag 545 - Biographical or Historical |
#|                Data (repeatable)          |
#|-------------------------------------------|
my $tag545 = 0;
my $t545_cnt = 0;
my $t545_array = "";
my $t545a = ""; # Biographical or historical data

#|-----------------------------------------------|
#| MARC tag 600 -  Subject Added Entry -         |
#|                 Personal Name (repeatable)    |
#|-----------------------------------------------|
my $tag600 = 0;
my $t600_cnt = 0;
my $t600_array = "";
my $t600a = ""; # Personal name
my $t600c = ""; # Titles associated with a name
my $t600d = ""; # Dates associated with a name
my $t600q = ""; # Fuller form of name
my $t600x = ""; # General subdivision
my $t600y = ""; # Chronological subdivision
my $t600z = ""; # Geographical subdivision

#|------------------------------------------------|
#| MARC tag 610 -  Subject Added Entry -          |
#|                 Corporate Name (repeatable)    |
#|------------------------------------------------|
my $tag610 = 0;
my $t610_cnt = 0;
my $t610_array = "";
my $t610a = ""; # Corporate name
my $t610b = ""; # Subordinate unit
my $t610x = ""; # General subdivision
my $t610y = ""; # Chronological subdivision
my $t610z = ""; # Geographical subdivision

#|------------------------------------------------|
#| MARC tag 611 -  Subject Added Entry -          |
#|                 Meeting Name (repeatable)      |
#|------------------------------------------------|
my $tag611 = 0;
my $t611_cnt = 0;
my $t611_array = "";
my $t611a = ""; # Meeting name 
my $t611c = ""; # Location of meeting
my $t611x = ""; # General subdivision
my $t611y = ""; # Chronological subdivision
my $t611z = ""; # Geographical subdivision

#|-----------------------------------------------|
#| MARC tag 630 -  Subject Added Entry -         | 
#|                 Uniform Name (repeatable)     |
#|-----------------------------------------------|
my $tag630 = 0;
my $t630_cnt = 0;
my $t630_array = "";
my $t630a = ""; # Uniform title

#|-----------------------------------------------|
#| MARC tag 650 -  Subject Added Entry -         |
#|                 Topical Term (repeatable)     |
#|-----------------------------------------------|
my $tag650 = 0;
my $t650_cnt = 0;
my $t650_array = "";
my $t650a = ""; # Topical term or geographic name entry element
my $t650x = ""; # General subdivision
my $t650y = ""; # Chronological subdivision
my $t650z = ""; # Geographic subdivision

#|-----------------------------------------------|
#| MARC tag 651 -  Subject Added Entry -         |
#|                 Geographic Name (repeatable)  |
#|-----------------------------------------------|
my $tag651 = 0;
my $t651_cnt = 0;
my $t651_array = "";
my $t651a = ""; # Geographic name
my $t651x = ""; # General subdivision
my $t651y = ""; # Chronological subdivision
my $t651z = ""; # Geographic subdivision

#|-----------------------------------------|
#| MARC tag 655 - Index Term -             |
#|                Genre/Form (repeatable)  |
#|-----------------------------------------|
my $tag655 = 0;
my $t655_cnt = 0;
my $t655_array = "";
my $t655a = ""; # Genre/form data or focus term
my $t655x = ""; # General subdivision
my $t655y = ""; # Chronological subdivision
my $t655z = ""; # Geographic subdivision

#|-----------------------------------------|
#| MARC tag 656 - Bibliographic -          |
#|                Concise (repeatable)     |
#|-----------------------------------------|
my $tag656 = 0;
my $t656_cnt = 0;
my $t656_array = "";
my $t656a = ""; # Occupation
my $t656x = ""; # General subdivision
my $t656y = ""; # Chronological subdivision
my $t656z = ""; # Geographic subdivision

#|-------------------------------------------|
#| MARC tag 700 - Added Entry -              |
#|                Personal Name (repeatable) |
#|-------------------------------------------|
my $tag700 = 0;
my $t700_cnt = 0;
my $t700_array = "";
my $t700a = ""; # Personal name
my $t700c = ""; # Titles and other words associated with a name
my $t700d = ""; # Dates associated with a name
my $t700q = ""; # Fuller form of name

#|--------------------------------------------|
#| MARC tag 710 - Added Entry -               |
#|                Corporate Name (repeatable) |
#|--------------------------------------------|
my $tag710 = 0;
my $t710_cnt = 0;
my $t710_array = "";
my $t710a = ""; # Corporate name
my $t710b = ""; # Subordinate unit

#|-----------------------|
#| Additional variables  |
#|-----------------------|
my ($temp, $extract1, $extract2, $t524_id, $title);
my $line = "";
my ($tab1, $tab2, $tabi);
my $start = 0;
               
open (MARCXML, "$destfile");
while ($line = <MARCXML>) {
      #|--------------------------------------------|
      #| MARC tag 100 - Main Entry -  Personal Name |
      #|--------------------------------------------|
       if ($line =~ /<datafield tag="100"/) {
           $tag100 = 1;
           $start = 1;
           while ($start) {
                  $line = <MARCXML>;
                  #|-----------------------------------|
                  #| Subfield a = Personal Name        |
                  #|-----------------------------------|
                  if ($line =~ /<subfield code="a"/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t100a = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  #|-----------------------------------|
                  #| Subfield q = Fuller form of name  |
                  #|-----------------------------------|
                  if ($line =~ /<subfield code="q"/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t100q = substr($line,($tab1+19),($tab2-($tab1+19)));
                      $t100a = $t100a . " " . $t100q;
                     }
                  #|-----------------------------------|
                  #| Subfield c = Titles for name      |
                  #|-----------------------------------|
                  if ($line =~ /<subfield code="c"/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t100c = substr($line,($tab1+19),($tab2-($tab1+19)));
                    }
                  #|-----------------------------------|
                  #| Subfield d = Dates for name       |
                  #|-----------------------------------|
                  if ($line =~ /<subfield code="d"/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t100d = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                     }
                  }
               }
      #|----------------------------------------------|
      #| Marc tag 110 - Main Entry - Corporate Name   |
      #|----------------------------------------------|
       if ($line =~ /<datafield tag="110"/) {
           $tag110 = 1;
           $start = 1;
           while ($start) {
                  $line = <MARCXML>;
                  #|------------------------------------|
                  #| Subfield a = Corporate Name        |
                  #|------------------------------------|
                  if ($line =~ /<subfield code="a"/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t110a = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  #|------------------------------------------|
                  #| Subfield b = Name of sub unit or meeting |
                  #|------------------------------------------|
                  if ($line =~ /<subfield code="b"/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t110b = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                     }
                  }
               }
      #|---------------------------------------|
      #| MARC tag 245 -  Title Statement       |
      #|---------------------------------------|
        if ($line =~ /<datafield tag="245"/) {
           $tag245 = 1;
           $start = 1;
           while ($start) {
                  $line = <MARCXML>;
                  #|------------------------------------|
                  #| Subfield a = Title                 |
                  #|------------------------------------|
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t245a = substr($line,($tab1+19),($tab2-($tab1+19)));
                      $title = $t245a;
                      }             
                  #|------------------------------------|
                  #| Subfield b = Remainder of title    |
                  #|------------------------------------|
                  if ($line =~ /<subfield code="b">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t245b = substr($line,($tab1+19),($tab2-($tab1+19)));
                      $title = $t245a . $t245b;
                      }             
                  #|-------------------------------------|
                  #| Subfield f = Date                   |
                  #|-------------------------------------|
                  if ($line =~ /<subfield code="f">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t245f = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|-------------------------------------|
                  #| Subfield g = Bulk dates             |
                  #|-------------------------------------|
                  if ($line =~ /<subfield code="g">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t245g = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                     }
                  }
               }
      #|--------------------------------------|
      #| MARC tag 300 - Physical Description  |
      #|                (repeatable)          |
      #|--------------------------------------|
       if ($line =~ /<datafield tag="300"/) {
           $tag300 = 1;
           $start = 1;
           $t300_3 = "";
           while ($start) {
                  $line = <MARCXML>;
                  #|--------------------------------------------|
                  #| Subfield 3 = Materials specified           |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="3">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t300_3 = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|------------------------------------|
                  #| Subfield a = Extent                |
                  #|------------------------------------|
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t300a = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|------------------------------------|
                  #| Subfield f = Unit                  |
                  #|------------------------------------|
                  if ($line =~ /<subfield code="f">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t300f = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                      $t300_cnt++;
                      if ($t300_3 ne "") {
                          $t300_array .= "<tag>" . $t300_3 . " " . $t300a;
                         }
                      else {
                            $t300_array .= "<tag>" . $t300a;
                           }
                      if ($t300f ne "") {
                          $t300_array .= " " . $t300f . "</tag>";
                         }
                      else {
                            $t300_array .= "</tag>";
                           }
                     }
                  }
               }
      #|----------------------------------------------|
      #| MARC tags 351 - Organization of Materials    |
      #|----------------------------------------------|
       if ($line =~ /<datafield tag="351"/) {
           $tag351 = 1;
           $start = 1;
           while ($start) {
                  $line = <MARCXML>;
                  #|----------------------------------------|
                  #| Subfield a = Organization of materials |
                  #|----------------------------------------|
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t351a = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                     }
                  }
               }
      #|----------------------------------------------|
      #| Get Description of item: MARC tags 500 - 545 |
      #|----------------------------------------------|
      #| MARC tag 500 - General Note (repeatable)     |
      #|----------------------------------------------|
       if ($line =~ /<datafield tag="500"/) {
           $tag500 = 1;
           $start = 1;
           while ($start) {
                  $line = <MARCXML>;
                  #|------------------------------------|
                  #| Subfield a = General Note          |
                  #|------------------------------------|
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t500a = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                      $t500_cnt++;
                      $t500_array .= "<tag>" . $t500a . "</tag>";
                     }
                  }
               }
      #|----------------------------------------------|
      #| MARC tag 506 - Restrictions on Access  Note  |
      #|----------------------------------------------|
       if ($line =~ /<datafield tag="506"/) {
           $tag506 = 1;
           $start = 1;
           $t506_3 = "";
           while ($start) {
                  $line = <MARCXML>;
                  #|--------------------------------------------|
                  #| Subfield 3 = Terms governing Access        |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="3">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t506_3 = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  #|-------------------------------------|
                  #| Subfield a = Terms governing access |
                  #|-------------------------------------|
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t506a = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                      $t506_cnt++;
                      if ($t506_3 ne "") {
                          $t506_array .= "<tag>" . $t506_3 . " " . $t506a . "</tag>";
                         }
                      else {
                             $t506_array .= "<tag>" . $t506a . "</tag>";
                           }
                     }
                  }
               }
      #|----------------------------------------------|
      #| MARC tag 520 - Summary, etc. (repeatable)    |
      #|----------------------------------------------|
       if ($line =~ /<datafield tag="520"/) {
           $tag520 = 1;
           $start = 1;
           while ($start) {
                  $line = <MARCXML>;
                  #|--------------------------------------------|
                  #| Subfield a = Summary, etc.                 |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t520a = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                      $t520_cnt++;
                      $t520_array .= "<tag>" . $t520a . "</tag>";
                     }
                  }
               }
      #|----------------------------------------------|
      #| MARC tag 524 - Preferred Citation Note       |
      #|----------------------------------------------|
       if ($line =~ /<datafield tag="524"/) {
           $tag524 = 1;
           $start = 1;
           while ($start) {
                  $line = <MARCXML>;
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $temp = substr($line,($tab1+19),($tab2-($tab1+19)));
                      if ($temp =~ /#/) {
                          $tab1 = index($line,"#",0);
                          $tab2 = index($line,".",0);
                          $t524_id = substr($line,($tab1+1),($tab2-($tab1+1)));
                          $t524a = $temp;
                          }
                      }             
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                     }
                  }
               }
      #|-----------------------------------------|
      #| MARC tag 538 - Technical information    |
      #|-----------------------------------------|
       if ($line =~ /<datafield tag="538"/) {
           $tag538 = 1;
           $start = 1;
           while ($start) {
                  $line = <MARCXML>;
                  #|--------------------------------------------|
                  #| Subfield a = Note                          |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t538a = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                      $t538_cnt++;
                      $t538_array .= "<tag>" . $t538a . "</tag>";
                     }
                  }
               }
      #|------------------------------------------------------|
      #| MARC tag 540 -  Governing Use and Reproduction Note  |
      #|------------------------------------------------------|
       if ($line =~ /<datafield tag="540"/) {
           $tag540 = 1;
           $start = 1;
           while ($start) {
                  $line = <MARCXML>;
                  #|------------------------------------|
                  #| Subfield a = Summary, etc.         |
                  #|------------------------------------|
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $temp = substr($line,($tab1+19),($tab2-($tab1+19)));
                      $t540a .= "<tag>" . $temp . "</tag>";
                      }             
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                      $t540_cnt++;
                      $t540_array .= "<tag>" . $t540a . "</tag>";
                     }
                  }
               }
      #|--------------------------------------------------------|
      #| MARC tag 544 - Location of Other Archival Materials    |
      #|--------------------------------------------------------|
       if ($line =~ /<datafield tag="544"/) {
           $tag544= 1;
           $start = 1;
           while ($start) {
                  $line = <MARCXML>;
                  #|------------------------------------|
                  #| Subfield a = Note                  |
                  #|------------------------------------|
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t544a = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                      $t544_cnt++;
                      $t544_array .= "<tag>" . $t544a . "</tag>";
                     }
                  }
               }
      #|-------------------------------------------|
      #| MARC tag 545 - Biographical or Historical |
      #|                Data (repeatable)          |
      #|-------------------------------------------|
       if ($line =~ /<datafield tag="545"/) {
           $tag545 = 1;
           $start = 1;
           while ($start) {
                  $line = <MARCXML>;
                  #|----------------------------------------------|
                  #| Subfield a = Biographical or historical data |
                  #|----------------------------------------------|
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t545a = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                      $t545_cnt++;
                      $t545_array .= "<tag>" . $t545a . "</tag>";
                     }
                  }
               }
      #|-----------------------------------------------|
      #| MARC tag 600 -  Subject Added Entry -         |
      #|                 Personal Name (repeatable)    |
      #|-----------------------------------------------|
       if ($line =~ /<datafield tag="600"/) {
           $tag600 = 1;
           $start = 1;
           $t600c = ""; 
           $t600d = ""; 
           $t600q = ""; 
           $t600x = ""; 
           $t600y = ""; 
           $t600z = ""; 
           while ($start) {
                  $line = <MARCXML>;
                  #|------------------------------------|
                  #| Subfield a = Personal Name         |
                  #|------------------------------------|
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t600a = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|--------------------------------------------|
                  #| Subfield c = Titles associated with name   |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="c">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t600c = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|--------------------------------------------|
                  #| Subfield d = Dates associated with name    |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="d">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t600d = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|--------------------------------------------|
                  #| Subfield q = Fuller form of name           |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="q">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t600q = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|--------------------------------------------|
                  #| Subfield x  = General subdivision          |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="x">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t600x = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|--------------------------------------------|
                  #| Subfield y  = Chronological subdivision    |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="y">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t600y = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|--------------------------------------------|
                  #| Subfield z  = Geographical subdivision     |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="z">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t600z = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                      $t600_cnt++;
                      if ($t600q ne "") {
                          $t600a = $t600a . " " . $t600q;
                         }
                      if ($t600c ne "") {
                          $t600a = $t600a . " " . $t600c;
                         }
                      if ($t600d ne "") {
                          $t600a = $t600a . " " . $t600d;
                         }
                      if ($t600x ne "") {
                          $t600a = $t600a . " --  " . $t600x;
                         }
                      if ($t600y ne "") {
                          $t600a = $t600a . " --  " . $t600y;
                         }
                      if ($t600z ne "") {
                          $t600a = $t600a . " --  " . $t600z;
                         }
                      $t600_array .= "<tag>" . $t600a . "</tag>";
                     }
                  }
               }
      #|-----------------------------------------------|
      #| MARC tag 610 -  Subject Added Entry -         |
      #|                 Corporate Name (repeatable)   |
      #|-----------------------------------------------|
       if ($line =~ /<datafield tag="610"/) {
           $tag610 = 1;
           $t610b = "";
           $t610x = "";
           $t610y = "";
           $t610z = "";
           $start = 1;
           while ($start) {
                  $line = <MARCXML>;
                  #|------------------------------------|
                  #| Subfield a = Corporate Name        |
                  #|------------------------------------|
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t610a = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|--------------------------------------------|
                  #| Subfield b = Subordinate unit              |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="b">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t610b = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|--------------------------------------------|
                  #| Subfield x  = General subdivision          |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="x">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t610x = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|--------------------------------------------|
                  #| Subfield y  = Chronological subdivision    |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="y">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t610y = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|--------------------------------------------|
                  #| Subfield z  = Geographical subdivision     |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="z">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t610z = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                      $t610_cnt++;
                      if ($t610b ne "") {
                          $t610a = $t610a . " " . $t610b;
                         }
                      if ($t610x ne "") {
                          $t610a = $t610a . " --  " . $t610x;
                         }
                      if ($t610y ne "") {
                          $t610a = $t610a . " --  " . $t610y;
                         }
                      if ($t610z ne "") {
                          $t610a = $t610a . " --  " . $t610z;
                         }
                      $t610_array .= "<tag>" . $t610a . "</tag>";
                     }
                  }
               }
      #|-----------------------------------------------|
      #| MARC tag 611 -  Subject Added Entry -         |
      #|                 Meeting Name (repeatable)     |
      #|-----------------------------------------------|
       if ($line =~ /<datafield tag="611"/) {
           $tag611 = 1;
           $t611c = "";
           $t611x = "";
           $t611y = "";
           $t611z = "";
           $start = 1;
           while ($start) {
                  $line = <MARCXML>;
                  #|------------------------------------|
                  #| Subfield a = Meeting Name          |
                  #|------------------------------------|
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t611a = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|--------------------------------------------|
                  #| Subfield c = Location of Meeting           |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="b">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t611c = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|--------------------------------------------|
                  #| Subfield x  = General subdivision          |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="x">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t611x = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|--------------------------------------------|
                  #| Subfield y  = Chronological subdivision    |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="y">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t611y = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|--------------------------------------------|
                  #| Subfield z  = Geographical subdivision     |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="z">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t611z = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                      $t611_cnt++;
                      if ($t611c ne "") {
                          $t611a = $t611a . " " . $t611c;
                         }
                      if ($t611x ne "") {
                          $t611a = $t611a . " --  " . $t611x;
                         }
                      if ($t611y ne "") {
                          $t611a = $t611a . " --  " . $t611y;
                         }
                      if ($t611z ne "") {
                          $t611a = $t611a . " --  " . $t611z;
                         }
                      $t611_array .= "<tag>" . $t611a . "</tag>";
                     }
                  }
               }
      #|-----------------------------------------------|
      #| MARC tag 630 -  Subject Added Entry -         |
      #|                 Uniform Title (repeatable)    |
      #|-----------------------------------------------|
       if ($line =~ /<datafield tag="630"/) {
           $tag630 = 1;
           $start = 1;
           while ($start) {
                  $line = <MARCXML>;
                  #|------------------------------------|
                  #| Subfield a = Uniform Title         |
                  #|------------------------------------|
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t630a = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  if ($line =~ /<\/datafield/) {
                      $t630_cnt++;
                      $t630_array .= "<tag>" . $t630a . "</tag>";
                      $start = 0;
                     }
                  }
               }
      #|-----------------------------------------------|
      #| MARC tag 650 -  Subject Added Entry -         |
      #|                 Topical Term (repeatable)     |
      #|-----------------------------------------------|
       if ($line =~ /<datafield tag="650"/) {
           $tag650 = 1;
           $t650x = "";
           $t650y = "";
           $t650z = "";
           $start = 1;
           while ($start) {
                  $line = <MARCXML>;
                  #|------------------------------------|
                  #| Subfield a = Topical term          |
                  #|------------------------------------|
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t650a = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|--------------------------------------------|
                  #| Subfield x = General subdivision           |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="x">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t650x = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  #|--------------------------------------------|
                  #| Subfield y = Chronological subdivision     |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="y">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t650y = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  #|--------------------------------------------|
                  #| Subfield z = Geographic subdivision        |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="z">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t650z = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                      $t650_cnt++;
                      if ($t650x ne "") {
                          $t650a = $t650a . " --  " . $t650x;
                         }
                      if ($t650y ne "") {
                          $t650a = $t650a . " --  " . $t650y;
                         }
                      if ($t650z ne "") {
                          $t650a = $t650a . " --  " . $t650z;
                         }
                      $t650_array .= "<tag>" . $t650a . "</tag>";
                     }
                  }
               }
      #|-----------------------------------------------|
      #| MARC tag 651 -  Subject Added Entry -         |
      #|                 Geographic Name (repeatable)  |
      #|-----------------------------------------------|
       if ($line =~ /<datafield tag="651"/) {
           $tag651 = 1;
           $t651x = "";
           $t651y = "";
           $t651z = "";
           $start = 1;
           while ($start) {
                  $line = <MARCXML>;
                  #|------------------------------------|
                  #| Subfield a = Geographic Name       |
                  #|------------------------------------|
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t651a = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|--------------------------------------------|
                  #| Subfield x = General subdivision           |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="x">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t651x = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  #|--------------------------------------------|
                  #| Subfield y = Chronological subdivision     |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="y">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t651y = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  #|--------------------------------------------|
                  #| Subfield z = Geographic subdivision        |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="z">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t651z = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                      $t651_cnt++;
                      if ($t651x ne "") {
                          $t651a = $t651a . " --  " . $t651x;
                         }
                      if ($t651y ne "") {
                          $t651a = $t651a . " --  " . $t651y;
                         }
                      if ($t651z ne "") {
                          $t651a = $t651a . " --  " . $t651z;
                         }
                      $t651_array .= "<tag>" . $t651a . "</tag>";
                     }
                  }
               }
      #|-----------------------------------------------|
      #| MARC tag 655 -  Index Term -                  |
      #|                 Genre/Form  (repeatable)      |
      #|-----------------------------------------------|
       if ($line =~ /<datafield tag="655"/) {
           $tag655 = 1;
           $t655x = "";
           $t655y = "";
           $t655z = "";
           $start = 1;
           while ($start) {
                  $line = <MARCXML>;
                  #|------------------------------------|
                  #| Subfield a = Genre/Form            |
                  #|------------------------------------|
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t655a = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|--------------------------------------------|
                  #| Subfield x = General subdivision           |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="x">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t655x = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  #|--------------------------------------------|
                  #| Subfield y = Chronological subdivision     |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="y">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t655y = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  #|--------------------------------------------|
                  #| Subfield z = Geographic subdivision        |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="z">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t655z = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                      $t655_cnt++;
                      if ($t655x ne "") {
                          $t655a = $t655a . " --  " . $t655x;
                         }
                      if ($t655y ne "") {
                          $t655a = $t655a . " --  " . $t655y;
                         }
                      if ($t655z ne "") {
                          $t655a = $t655a . " --  " . $t655z;
                         }
                      $t655_array .= "<tag>" . $t655a . "</tag>";
                     }
                  }
               }
      #|-----------------------------------------|
      #| MARC tag 656 - Bibliographic -          |
      #|                Concise (repeatable)     |
      #|-----------------------------------------|
       if ($line =~ /<datafield tag="656"/) {
           $tag656 = 1;
           $start = 1;
           $t656x = "";
           $t656y = "";
           $t656z = "";
           while ($start) {
                  $line = <MARCXML>;
                  #|------------------------------------|
                  #| Subfield a = Occupation            |
                  #|------------------------------------|
                  if ($line =~ /<subfield code="a">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t656a = substr($line,($tab1+19),($tab2-($tab1+19)));
                      }             
                  #|--------------------------------------------|
                  #| Subfield x = General subdivision           |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="x">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t656x = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  #|--------------------------------------------|
                  #| Subfield y = Chronological subdivision     |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="y">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t656y = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  #|--------------------------------------------|
                  #| Subfield z = Geographic subdivision        |
                  #|--------------------------------------------|
                  if ($line =~ /<subfield code="z">/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t656z = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                      $t656_cnt++;
                      if ($t656x ne "") {
                          $t656a = $t656a . " --  " . $t656x;
                         }
                      if ($t656y ne "") {
                          $t656a = $t656a . " --  " . $t656y;
                         }
                      if ($t656z ne "") {
                          $t656a = $t656a . " --  " . $t656z;
                         }
                      $t656_array .= "<tag>" . $t656a . "</tag>";
                     }
                  }
               }

      #|-------------------------------------------|
      #| MARC tag 700 - Added Entry -              |
      #|                Personal Name (repeatable) |
      #|-------------------------------------------|
       if ($line =~ /<datafield tag="700"/) {
           $tag700 = 1;
           $start = 1;
           $t700c = "";
           $t700d = "";
           $t700q = "";
           while ($start) {
                  $line = <MARCXML>;
                  #|-----------------------------------|
                  #| Subfield a = Personal Name        |
                  #|-----------------------------------|
                  if ($line =~ /<subfield code="a"/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t700a = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  #|-----------------------------------|
                  #| Subfield c = Titles associated    |
                  #|-----------------------------------|
                  if ($line =~ /<subfield code="c"/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t700c = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  #|-----------------------------------|
                  #| Subfield d = Dates for name       |
                  #|-----------------------------------|
                  if ($line =~ /<subfield code="d"/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t700d = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  #|-----------------------------------|
                  #| Subfield q = Fuller form of name  |
                  #|-----------------------------------|
                  if ($line =~ /<subfield code="q"/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t700q = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                      $t700_cnt++;
                      if ($t700q ne "") {
                          $t700a = $t700a . " " . $t700q;
                         }
                      if ($t700c ne "") {
                          $t700a = $t700a . " " . $t700c;
                         }
                      if ($t700d ne "") {
                          $t700a = $t700a . " " . $t700d;
                         }
                      $t700_array .= "<tag>" . $t700a . "</tag>";
                     }
                  }
               }
      #|-------------------------------------------|
      #| MARC tag 710 - Added Entry -              |
      #|                Corporate Name (repeatable)|
      #|-------------------------------------------|
       if ($line =~ /<datafield tag="710"/) {
           $tag710 = 1;
           $t710b = "";
           $start = 1;
           while ($start) {
                  $line = <MARCXML>;
                  #|-----------------------------------|
                  #| Subfield a = Corporate Name       |
                  #|-----------------------------------|
                  if ($line =~ /<subfield code="a"/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t710a = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  #|-----------------------------------|
                  #| Subfield b = Subordinate name     |
                  #|-----------------------------------|
                  if ($line =~ /<subfield code="b"/) {
                      $tab1 = index($line,"<subfield",0);
                      $tab2 = index($line,"</subfield",0);
                      $t710b = substr($line,($tab1+19),($tab2-($tab1+19)));
                     }
                  if ($line =~ /<\/datafield/) {
                      $start = 0;
                      $t710_cnt++;
                      if ($t710b ne "") {
                          $t710a = $t710a . " " . $t710b;
                         }
                      $t710_array .= "<tag>" . $t710a . "</tag>";
                     }
                  }
               }
       }
close (MARCXML);
#|----------------------------------------------------------------|
#| Start pulling data out of the excel spreadsheet                |
#|----------------------------------------------------------------|
my $parser = Spreadsheet::ParseXLSX->new;
my $workbook = $parser->parse("$upload_dir/$filename");
my ( $series, $c0level, $type, $box, $folder, $unitid, $circa );
my ( $datefrom, $dateto, $express, $ttitle, $access, $creatpers );
my ( $creatcorp, $arrange,  $scope, $process, $persname );
my ( $corpname, $geogname, $subhead, $odd );
my ( $dims, $genre, $extent, $physdes, $userest ); 
my $series_flag = 0;
my $previous_row = 0;
my $skip_row = 0;
my $excel = $bibid . "_excel";
open (EXCEL, ">$upload_dir/$excel");
for my $worksheet ( $workbook->worksheets() ) {
    my ( $row_min, $row_max ) = $worksheet->row_range();
    my ( $col_min, $col_max ) = $worksheet->col_range();
    for my $row ( $row_min .. $row_max ) {
        for my $col ( $col_min .. $col_max ) {
            my $cell = $worksheet->get_cell( $row, $col );
            next unless $cell;
           #|--------------------------------------------|
           #| Row 0 is usually just a row of field       |
           #| descriptions.  We would normally skip it.  |
           #| However, if row 0 is a data field, we will |
           #| need to include it.  To check that, we will|
           #| check if col 0 contains the phrase "Series"|
           #| or "series" or "SERIES".                   |
           #|--------------------------------------------|

           #|-------------------------------|
           #| Series is in column 0 (A)     |
           #|-------------------------------|
             if ($col == 0) {
                 $series = $cell->value();
                 if (($row == 0) &&
                    (($series =~ /eries/) || 
                     ($series =~ /ERIES/))) { 
                      $skip_row = 1;
                     }
                 else {
                       $skip_row = 0;
                      }
                #|--------------------------------------------|
                #| we will skip this column if it is blank    |
                #| or if it appears to be a description field |
                #|--------------------------------------------|
                 if (($series ne "") && (!$skip_row)) {
                      print EXCEL "<series>$series</series>";
                      $series_flag = 1;
                     }
                 }
              #|---------------------------------|
              #| C0Level is in column 1 (B)      |
              #|---------------------------------|
               if ($col == 1) {
                   $c0level = $cell->value();
                  #|--------------------------------------------|
                  #| we will skip this column if it is blank    |
                  #| or if it appears to be a description field |
                  #|--------------------------------------------|
                   if (($c0level ne "") && (!$skip_row)) {
                        print EXCEL "<c0level>$c0level</c0level>";
                       }
                  }
              #|---------------------------------|
              #| Level Type is in column 2 (C)   |
              #|---------------------------------|
                if ($col == 2) {
                    $type = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($type ne "") && (!$skip_row)) {
                         print EXCEL "<type>$type</type>";
                        }
                   }
              #|---------------------------------|
              #| Box is in column 3 (D)          |
              #|---------------------------------|
                if ($col == 3) {
                    $box = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($box ne "") && (!$skip_row)) {
                         print EXCEL "<box>$box</box>";
                       }
                   }
              #|---------------------------------|
              #| Folder is in column 4 (E)       |
              #|---------------------------------|
                if ($col == 4) {
                    $folder = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($folder ne "") && (!$skip_row)) {
                         print EXCEL "<folder>$folder</folder>";
                       }
                   }
              #|---------------------------------|
              #| Item Number is in column 5 (F)  |
              #|---------------------------------|
                if ($col == 5) {
                    $unitid = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($unitid ne "") && (!$skip_row)) {
                         print EXCEL "<unitid>$unitid</unitid>";
                       }
                   }
              #|---------------------------------|
              #| circa date is in column 6  (G)  |
              #|---------------------------------|
                if ($col == 6) {
                    $circa = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($circa ne "") && (!$skip_row)) {
                         print EXCEL "<circa>$circa</circa>";
                        }
                   }
              #|---------------------------------|
              #| date from is in column 7 (H)    |
              #|---------------------------------|
                if ($col == 7) {
                    $datefrom = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($datefrom ne "") && (!$skip_row)) {
                        print EXCEL "<datefrom>$datefrom</datefrom>";
                        }
                   }
              #|---------------------------------|
              #| date to is in column 8 (I)      |
              #|---------------------------------|
                if ($col == 8) {
                    $dateto = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($dateto ne "") && (!$skip_row)) {
                        print EXCEL "<dateto>$dateto</dateto>";
                        }
                   }
              #|---------------------------------|
              #| expression is in column 9 (J)   |
              #|---------------------------------|
                if ($col == 9) {
                    $express = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($express ne "") && (!$skip_row)) {
                        print EXCEL "<express>$express</express>";
                        }
                   }
              #|---------------------------------|
              #| Title is in column 10 (K)       |
              #|---------------------------------|
                if ($col == 10) {
                    $ttitle = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($ttitle ne "") && (!$skip_row)) {
                        print EXCEL "<title>$ttitle</title>";
                        }
                   }
              #|--------------------------------------------|
              #| Creator Corporate Name is in column 11 (L) |
              #|--------------------------------------------|
                if ($col == 11) {
                    $creatcorp = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($creatcorp ne "") && (!$skip_row)) {
                        print EXCEL "<creatcorp>$creatcorp</creatcorp>";
                        }
                    }
              #|-------------------------------------------|
              #| Creator Personal Name is in column 12 (M) |
              #|-------------------------------------------|
                if ($col == 12) {
                    $creatpers = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($creatpers ne "") && (!$skip_row)) {
                        print EXCEL "<creatpers>$creatpers</creatpers>";
                        }
                    }
              #|------------------------------------------|
              #| Dimension is in column 13 (N)            |
              #|------------------------------------------|
                if ($col == 13) {
                    $dims = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($dims ne "") && (!$skip_row)) {
                        print EXCEL "<dims>$dims</dims>";
                        }
                    }
              #|------------------------------------------|
              #| Extent is in column 14 (O)               |
              #|------------------------------------------|
                if ($col == 14) {
                    $extent = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($extent ne "") && (!$skip_row)) {
                        print EXCEL "<extent>$extent</extent>";
                        }
                    }
              #|------------------------------------------|
              #| Physical Description is in column 15 (P) |
              #|------------------------------------------|
                if ($col == 15) {
                    $physdes = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($physdes ne "") && (!$skip_row)) {
                        print EXCEL "<physdes>$physdes</physdes>";
                        }
                    }
              #|------------------------------------------|
              #| Scope is in column 16 (Q)                |
              #|------------------------------------------|
                if ($col == 16) {
                    $scope = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($scope ne "") && (!$skip_row)) {
                        print EXCEL "<scope>$scope</scope>";
                        }
                    }
              #|------------------------------------------|
              #| Arrangement Note is in column 17 (R)     |
              #|------------------------------------------|
                if ($col == 17) {
                    $arrange = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($arrange ne "") && (!$skip_row)) {
                        print EXCEL "<arrange>$arrange</arrange>";
                        }
                    }
              #|------------------------------------------|
              #| Processing Note is in column 18 (S)      |
              #|------------------------------------------|
                if ($col == 18) {
                    $process = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($process ne "") && (!$skip_row)) {
                        print EXCEL "<process>$process</process>";
                        }
                    }
              #|------------------------------------------|
              #| Odd Note is in column 19 (T)             |
              #|------------------------------------------|
                if ($col == 19) {
                    $odd = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($odd ne "") && (!$skip_row)) {
                        print EXCEL "<odd>$odd</odd>";
                        }
                  }
              #|------------------------------------------|
              #| Personal Name is in column 20 (U)        |
              #|------------------------------------------|
                if ($col == 20) {
                    $persname = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($persname ne "") && (!$skip_row)) {
                        print EXCEL "<persname>$persname</persname>";
                        }
                    }
              #|------------------------------------------|
              #| Corporate Name is in column 21 (V)       |
              #|------------------------------------------|
                if ($col == 21) {
                    $corpname = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($corpname ne "") && (!$skip_row)) {
                        print EXCEL "<corpname>$corpname</corpname>";
                        }
                    }
              #|------------------------------------------|
              #| Geographic Name is in column 22 (W)      |
              #|------------------------------------------|
                if ($col == 22) {
                    $geogname = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($geogname ne "") && (!$skip_row)) {
                        print EXCEL "<geogname>$geogname</geogname>";
                        }
                    }
              #|------------------------------------------|
              #| Subject Headings is in column 23 (X)     |
              #|------------------------------------------|
                if ($col == 23) {
                    $subhead = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($subhead ne "") && (!$skip_row)) {
                        print EXCEL "<subject>$subhead</subject>";
                        }
                    }
              #|------------------------------------------|
              #| Genre is in column 24 (Y)                |
              #|------------------------------------------|
                if ($col == 24) {
                    $genre = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($genre ne "") && (!$skip_row)) {
                        print EXCEL "<genre>$genre</genre>";
                        }
                    }
              #|--------------------------------------|
              #| Access Restrictions in column 25 (Z) |
              #|--------------------------------------|
                if ($col == 25) {
                    $access = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($access ne "") && (!$skip_row)) {
                        print EXCEL "<access>$access</access>";
                        }
                   }
              #|--------------------------------------|
              #| Use Restrictions in column 26 (AA)   |
              #|--------------------------------------|
                if ($col == 26) {
                    $userest = $cell->value();
                   #|--------------------------------------------|
                   #| we will skip this column if it is blank    |
                   #| or if it appears to be a description field |
                   #|--------------------------------------------|
                    if (($userest ne "") && (!$skip_row)) {
                        print EXCEL "<userest>$userest</userest>";
                        }
                   }
            }    # For my col
           print EXCEL "\n";
          }      # for my row
        }        # for my worksheet
close (EXCEL);
#|----------------------------------------------------------------|
#| The EAD file will be built here                                |
#|----------------------------------------------------------------|
$tab1 = index($filename,".",0);
my $id = substr($filename,0,$tab1);
#open (OUTPUT, "> :encoding(UTF-8)", "$upload_dir/$id.xml");
open (OUTPUT, ">$upload_dir/$id.xml");
                  
#|----------------------------------------------------------------|
#| Add introductory EAD information                               |
#|----------------------------------------------------------------|
print OUTPUT qq{<?xml version="1.0" encoding="utf-8"?>\n};
print OUTPUT qq{<!DOCTYPE ead PUBLIC "+//ISBN 1-931666-00-8//DTD };
print OUTPUT qq{ead.dtd (Encoded Archival Description (EAD) Version };
print OUTPUT qq{2002)//EN" "../dtds/ead.dtd">\n};
print OUTPUT qq{<?xml-stylesheet type="text/xsl" href="../styles/style.xsl" ?>};
print OUTPUT qq{<ead>\n};
print OUTPUT qq{  <eadheader repositoryencoding="iso15511" };
print OUTPUT qq{relatedencoding="MARC21" countryencoding="iso3166-1" };
print OUTPUT qq{scriptencoding="iso15924" dateencoding="iso8601" };
print OUTPUT qq{langencoding="iso639-2b">\n};
print OUTPUT qq{    <eadid mainagencycode="nic" countrycode="us"};
#|----------------------------------------------------------------|
#| Title and ID of Finding Aid                                    |
#|----------------------------------------------------------------|
print OUTPUT qq{ publicid="-//Cornell University::};
print OUTPUT qq{Cornell University Library::};
print OUTPUT qq{Division of Rare and Manuscript Collections};
print OUTPUT qq{//TEXT(US::NIC::};
print OUTPUT qq{$id};
print OUTPUT qq{::};
$tab1 = index($title,",",0);
if ($tab1 > 0) {
    $temp = substr($title,0,($tab1));
    print OUTPUT qq{$temp};
    }
else {
      print OUTPUT qq{$title};
     }
print OUTPUT qq{)//EN" >};
print OUTPUT qq{$id};
print OUTPUT qq{.xml</eadid>\n};
print OUTPUT qq{  <filedesc>\n};
print OUTPUT qq{    <titlestmt>\n};
print OUTPUT qq{      <titleproper>Guide to the };
print OUTPUT qq{$title};
#|----------------------------------------------------------------|
#| Date for Finding Aid                                           |
#|----------------------------------------------------------------|
if ($t245f ne "") {
    print OUTPUT qq{ <date>$t245f</date>};
    }
print OUTPUT qq{</titleproper>\n};
#|----------------------------------------------------------------|
#| Create sort version of title                                   |
#|----------------------------------------------------------------|
$tab1 = index($title," ",0);
my $sort = "";
my $first = substr($title,0,($tab1));
my $second = substr($title,($tab1+1));
if (($first eq "A") || ($first eq "An") || ($first eq "The")) {
     $sort = $second . ", " . $first;
    }
else {
      $sort = $title;
     }
print OUTPUT qq{      <titleproper type="sort">$sort</titleproper>\n};
print OUTPUT qq{      <author>Compiled by RMC Staff</author>\n};
print OUTPUT qq{    </titlestmt>\n};
print OUTPUT qq{    <publicationstmt>\n};
print OUTPUT qq{      <publisher>Division of Rare and Manuscript Collections, };
print OUTPUT qq{Cornell University Library</publisher>\n};
#|----------------------------------------------------------------|
#| Calculate published date                                       |
#|----------------------------------------------------------------|
my $thedate;
my $month;
my $year;
my $cdate;
open(DATE, "date|");
$thedate = <DATE>;
close(DATE);
$month = substr($thedate,4,3);
$tab1 = index($thedate,"201",0);
$year = substr($thedate,$tab1,4);
if ($month eq "Jan") {
    $cdate = "January " . $year;
    }
elsif ($month eq "Feb") {
       $cdate = "February " . $year;
       }
elsif ($month eq "Mar") {
       $cdate = "March " . $year;
       }
elsif ($month eq "Apr") {
       $cdate = "April " . $year;
       }
elsif ($month eq "May") {
       $cdate = "May " . $year;
       }
elsif ($month eq "Jun") {
       $cdate = "June " . $year;
       }
elsif ($month eq "Jul") {
       $cdate = "July " . $year;
       }
elsif ($month eq "Aug") {
       $cdate = "August " . $year;
       }
elsif ($month eq "Sep") {
       $cdate = "September " . $year;
       }
elsif ($month eq "Oct") {
       $cdate = "October " . $year;
       }
elsif ($month eq "Nov") {
       $cdate = "November " . $year;
       }
elsif ($month eq "Dec") {
       $cdate = "December " . $year;
       }
else {
       $cdate = $year;
     }
print OUTPUT qq{      <date>$cdate</date>\n};
print OUTPUT qq{    </publicationstmt>\n};
print OUTPUT qq{    <notestmt>\n};
print OUTPUT qq{      <note audience="internal">\n};
print OUTPUT qq{        <p>\n};
print OUTPUT "           <subject>{**Subject Code, Use collection ";
print OUTPUT "subject browse list}</subject>\n";
print OUTPUT qq{        </p>\n};
print OUTPUT qq{      </note>\n};
print OUTPUT qq{    </notestmt>\n};
print OUTPUT qq{  </filedesc>\n};
print OUTPUT qq{  <profiledesc>\n};
print OUTPUT qq{    <creation>Finding aid encoded by RMC Staff, };
print OUTPUT qq{<date>$cdate</date></creation>\n};
print OUTPUT qq{  </profiledesc>\n};
print OUTPUT qq{  </eadheader>\n};
print OUTPUT qq{  <frontmatter>\n};
#|----------------------------------------------------------------|
#| Title information                                              |
#|----------------------------------------------------------------|
print OUTPUT qq{           <titlepage>\n};
print OUTPUT qq{                   <titleproper>Guide to the $title};
print OUTPUT qq{<lb/><date type="inclusive" encodinganalog="245\$f">};
print OUTPUT qq{$t245f</date></titleproper>\n};
if ($t245g ne "") {
    print OUTPUT qq{                   <date type="bulk" encodinganalog=};
    print OUTPUT qq{"245\$g">[$t245g]</date></titleproper>\n};
   }
print OUTPUT qq{           <num>Collection Number: $t524_id</num>\n};
print OUTPUT qq{           <publisher>Division of Rare and Manuscript Collections };
print OUTPUT qq{<lb/>Cornell University Library</publisher>\n};
print OUTPUT qq{           <list type="deflist">\n};
print OUTPUT qq{                   <defitem>\n};
print OUTPUT qq{                           <label>Contact Information:</label>\n};
print OUTPUT qq{                           <item> Division of Rare and Manuscript Collections};
print OUTPUT qq{<lb/> 2B Carl A. Kroch Library<lb/> Cornell University<lb/> Ithaca, NY 14853};
print OUTPUT qq{<lb/> (607) 255-3530<lb/> Fax: (607) 255-9524<lb/>};
print OUTPUT qq{<extref href="mailto:rareref\@cornell.edu">rareref\@cornell.edu</extref>};
print OUTPUT qq{<lb/><extref href="http://rmc.library.cornell.edu">};
print OUTPUT qq{http://rmc.library.cornell.edu</extref><lb/></item>\n};
print OUTPUT qq{                   </defitem>\n};
print OUTPUT qq{                   <defitem>\n};
print OUTPUT qq{                           <label>Compiled by:</label>\n};
print OUTPUT qq{                           <item>RMC Staff</item>\n};
print OUTPUT qq{                   </defitem>\n};
print OUTPUT qq{                   <defitem>\n};
print OUTPUT qq{                           <label>Date completed:</label>\n};
print OUTPUT qq{                           <item>$cdate</item>\n};
print OUTPUT qq{                   </defitem>\n};
print OUTPUT qq{                   <defitem>\n};
print OUTPUT qq{                           <label>EAD encoding:</label>\n};
print OUTPUT qq{                           <item>RMC Staff, $cdate</item>\n};
print OUTPUT qq{                   </defitem>\n};
print OUTPUT qq{                   <defitem>\n};
print OUTPUT qq{                           <label>Date modified:</label>\n};
print OUTPUT qq{                           <item>RMC Staff, $cdate</item>\n};
print OUTPUT qq{                   </defitem>\n};
print OUTPUT qq{           </list>\n};
print OUTPUT qq{           <date>&#x00A9; $year Division of Rare and Manuscript Collections, };
print OUTPUT qq{Cornell University Library</date>\n};
print OUTPUT qq{           </titlepage>\n};
print OUTPUT qq{  </frontmatter>\n};
print OUTPUT qq{  <archdesc level="collection">\n};
print OUTPUT qq{    <did>\n};
print OUTPUT qq{            <head id="a1">DESCRIPTIVE SUMMARY</head>\n};
print OUTPUT qq{            <unittitle label="Title:" encodinganalog="MARC 245">$title\n};
if ($t245f ne "") {
    print OUTPUT qq{            <unitdate encodinganalog="MARC 245" type="inclusive">};
    print OUTPUT qq{$t245f</unitdate>\n};
    }
if ($t245g ne "") {
    print OUTPUT qq{            <unitdate encodinganalog="MARC 245" type="bulk">};
    print OUTPUT qq{$t245g</unitdate>\n};
    }
print OUTPUT qq{            </unittitle>\n};
print OUTPUT qq{            <unitid label="Collection Number:">$t524_id</unitid>\n};
#|-----------------------------------------|
#| Get Creator information for Finding Aid |
#|-----------------------------------------|
print OUTPUT qq{            <origination label="Creator:">\n};
#|-----------------------------------------|
#| Check for Corporate Name, if available  |
#|-----------------------------------------|
if ($tag110) {
    print OUTPUT qq{                    <corpname encodinganalog="110" normal="};
    print OUTPUT qq{$t110a };
    if ($t110b ne "") {
        print OUTPUT qq{$t110b };
        }
    print OUTPUT qq{">};
    print OUTPUT qq{$t110a };
    if ($t110b ne "") {
        print OUTPUT qq{$t110b };
        }
    print OUTPUT qq{</corpname>\n};
    }
#|-----------------------------------------|
#| Get Author from tag 100, if available   |
#|-----------------------------------------|
if ($tag100) {
    print OUTPUT qq{                    <persname encodinganalog="100" normal="};
    print OUTPUT qq{$t100a };
    if ($t100d ne "") {
        print OUTPUT qq{$t100d };
        }
    my $lastname = "";
    my $firstname = "";
    $tab1 = index($t100a,",",0);
    $lastname = substr($t100a,0,($tab1));
    $firstname = substr($t100a,($tab1+1));
    print OUTPUT qq{">$firstname $lastname};
    if ($t100d ne "") {
        print OUTPUT qq{, $t100d };
        }
    print OUTPUT qq{                    </persname>\n};
    }
print OUTPUT qq{            </origination>\n};
if ($tag300) {
    $tabi = 0;
    while ($tabi < $t300_cnt) {
           if ($tabi == 0) {
               $tab1 = index($t300_array,"<tag>",0);
               $tab2 = index($t300_array,"</tag>",0);
               }
           else {
                 $tab1 = index($t300_array,"<tag>",($tab2+5));
                 $tab2 = index($t300_array,"</tag>",($tab2+5));
                }
           $temp = substr($t300_array,($tab1+5),($tab2-($tab1+5)));
           print OUTPUT qq{            <physdesc label="Quantity:" };
           print OUTPUT qq{encodinganalog="MARC 300"><extent>$temp</extent></physdesc>\n};
           $tabi++;
           }
    }
print OUTPUT "            <physdesc label=\"Forms of Material:\">";
print OUTPUT "{**Comma Separated List Of Material Formats,End With Period**.}";
print OUTPUT "</physdesc>\n";
print OUTPUT qq{            <repository label="Repository:">};
print OUTPUT qq{Division of Rare and Manuscript Collections, };
print OUTPUT qq{Cornell University Library</repository>\n};
if ($tag520) {
    $tabi = 0;
    while ($tabi < $t520_cnt) {
           if ($tabi == 0) {
               $tab1 = index($t520_array,"<tag>",0);
               $tab2 = index($t520_array,"</tag>",0);
               }
           else {
                 $tab1 = index($t520_array,"<tag>",($tab2+5));
                 $tab2 = index($t520_array,"</tag>",($tab2+5));
                }
           $temp = substr($t520_array,($tab1+5),($tab2-($tab1+5)));
           print OUTPUT qq{            <abstract label="Abstract:">};
           print OUTPUT qq{$temp</abstract>\n};
           $tabi++;
          }
     }
print OUTPUT qq{            <langmaterial label="Language:">};
print OUTPUT qq{Collection material in <language encodinganalog="041" };
print OUTPUT qq{langcode="eng">English</language>\n};
print OUTPUT qq{            </langmaterial>\n};
print OUTPUT qq{    </did>\n};
#|----------------------------------------------------------------|
#| Organizational Information or Biographical Note                |
#|----------------------------------------------------------------|
if ($tag545) {
    print OUTPUT qq{    <bioghist encodinganalog="MARC 545">\n};
    if ($tag110) {
        print OUTPUT qq{       <head id="a2" altrender="organization">};
        print OUTPUT qq{ORGANIZATIONAL HISTORY</head>\n};
       }
    if ($tag100) {
        print OUTPUT qq{            <head id="a2" altrender="biography">};
        print OUTPUT qq{BIOGRAPHICAL NOTE </head>\n};
       }
    $tabi = 0;
    while ($tabi < $t545_cnt) {
           if ($tabi == 0) {
               $tab1 = index($t545_array,"<tag>",0);
               $tab2 = index($t545_array,"</tag>",0);
               }
           else {
                 $tab1 = index($t545_array,"<tag>",($tab2+5));
                 $tab2 = index($t545_array,"</tag>",($tab2+5));
                }
           $temp = substr($t545_array,($tab1+5),($tab2-($tab1+5)));
           print OUTPUT qq{            <p>$temp</p>\n};
           $tabi++;
          }
     print OUTPUT qq{    </bioghist>\n};
   }
#|----------------------------------------------------------------|
#| Collection Description                                         |
#|----------------------------------------------------------------|
if ($tag520) {
    print OUTPUT qq{    <scopecontent encodinganalog="MARC 520">\n};
    print OUTPUT qq{            <head id="a3">COLLECTION DESCRIPTION};
    print OUTPUT qq{</head>\n};
    $tabi = 0;
    while ($tabi < $t520_cnt) {
           if ($tabi == 0) {
               $tab1 = index($t520_array,"<tag>",0);
               $tab2 = index($t520_array,"</tag>",0);
               }
           else {
                 $tab1 = index($t520_array,"<tag>",($tab2+5));
                 $tab2 = index($t520_array,"</tag>",($tab2+5));
                }
           $temp = substr($t520_array,($tab1+5),($tab2-($tab1+5)));
           print OUTPUT qq{                  <p>$temp</p>\n};
           $tabi++;
          }
     print OUTPUT qq{    </scopecontent>\n};
    }
print OUTPUT qq{    <controlaccess>\n};
print OUTPUT qq{            <head id="a7">SUBJECTS</head>\n};
print OUTPUT qq{            <controlaccess>\n};
print OUTPUT qq{                    <head>Names: </head>\n};
if ($tag100) {
    print OUTPUT qq{                    <persname encodinganalog=};
    print OUTPUT qq{"MARC 100">$t100a};
    if ($t100q ne "") {
        print OUTPUT qq{ $t100q};
        }
    if ($t100c ne "") {
        print OUTPUT qq{ $t100c};
        }
    if ($t100d ne "") {
        print OUTPUT qq{ $t100d};
        }
     print OUTPUT qq{                   </persname>\n};
    }
if ($tag110) {
    print OUTPUT qq{                    <corpname encodinganalog=};
    print OUTPUT qq{"MARC 110">$t110a</corpname>\n};
    }
if ($tag600) {
    $tabi = 0;
    while ($tabi < $t600_cnt) {
           if ($tabi == 0) {
               $tab1 = index($t600_array,"<tag>",0);
               $tab2 = index($t600_array,"</tag>",0);
               }
           else {
                 $tab1 = index($t600_array,"<tag>",($tab2+5));
                 $tab2 = index($t600_array,"</tag>",($tab2+5));
                }
           $temp = substr($t600_array,($tab1+5),($tab2-($tab1+5)));
           print OUTPUT qq{                    <persname encodinganalog=};
           print OUTPUT qq{"MARC 600">$temp</persname>\n};
           $tabi++;
          }
    }
if ($tag700) {
    $tabi = 0;
    while ($tabi < $t700_cnt) {
           if ($tabi == 0) {
               $tab1 = index($t700_array,"<tag>",0);
               $tab2 = index($t700_array,"</tag>",0);
               }
           else {
                 $tab1 = index($t700_array,"<tag>",($tab2+5));
                 $tab2 = index($t700_array,"</tag>",($tab2+5));
                }
           $temp = substr($t700_array,($tab1+5),($tab2-($tab1+5)));
           print OUTPUT qq{                    <persname encodinganalog=};
           print OUTPUT qq{"MARC 700">$temp</persname>\n};
           $tabi++;
          }
    }
if ($tag610) {
    $tabi = 0;
    while ($tabi < $t610_cnt) {
           if ($tabi == 0) {
               $tab1 = index($t610_array,"<tag>",0);
               $tab2 = index($t610_array,"</tag>",0);
               }
           else {
                 $tab1 = index($t610_array,"<tag>",($tab2+5));
                 $tab2 = index($t610_array,"</tag>",($tab2+5));
                }
           $temp = substr($t610_array,($tab1+5),($tab2-($tab1+5)));
           print OUTPUT qq{                    <corpname encodinganalog=};
           print OUTPUT qq{"MARC 610">$temp</corpname>\n};
           $tabi++;
          }
    }
if ($tag710) {
    $tabi = 0;
    while ($tabi < $t710_cnt) {
           if ($tabi == 0) {
               $tab1 = index($t710_array,"<tag>",0);
               $tab2 = index($t710_array,"</tag>",0);
               }
           else {
                 $tab1 = index($t710_array,"<tag>",($tab2+5));
                 $tab2 = index($t710_array,"</tag>",($tab2+5));
                }
           $temp = substr($t710_array,($tab1+5),($tab2-($tab1+5)));
           print OUTPUT qq{                    <corpname encodinganalog=};
           print OUTPUT qq{"MARC 710">$temp</corpname>\n};
           $tabi++;
          }
    }
if ($tag611) {
    $tabi = 0;
    while ($tabi < $t611_cnt) {
           if ($tabi == 0) {
               $tab1 = index($t611_array,"<tag>",0);
               $tab2 = index($t611_array,"</tag>",0);
               }
           else {
                 $tab1 = index($t611_array,"<tag>",($tab2+5));
                 $tab2 = index($t611_array,"</tag>",($tab2+5));
                }
           $temp = substr($t611_array,($tab1+5),($tab2-($tab1+5)));
           print OUTPUT qq{                    <corpname encodinganalog=};
           print OUTPUT qq{"MARC 611">$temp</corpname>\n};
           $tabi++;
          }
    }
print OUTPUT qq{    </controlaccess>\n};
if ($tag630) {
    print OUTPUT qq{    <controlaccess>\n};
    print OUTPUT qq{      <head>Titles: </head>\n};
    $tabi = 0;
    while ($tabi < $t630_cnt) {
           if ($tabi == 0) {
               $tab1 = index($t630_array,"<tag>",0);
               $tab2 = index($t630_array,"</tag>",0);
               }
           else {
                 $tab1 = index($t630_array,"<tag>",($tab2+5));
                 $tab2 = index($t630_array,"</tag>",($tab2+5));
                }
           $temp = substr($t630_array,($tab1+5),($tab2-($tab1+5)));
           print OUTPUT qq{      <title encodinganalog="MARC 630">};
           print OUTPUT qq{$temp</title>\n};
           $tabi++;
          }
    print OUTPUT qq{    </controlaccess>\n};
    }
if (($tag650) || ($tag656)) {
    print OUTPUT qq{    <controlaccess>\n};
    print OUTPUT qq{      <head>Subjects: </head>\n};
    if ($tag650) {
        $tabi = 0;
        while ($tabi < $t650_cnt) {
               if ($tabi == 0) {
                   $tab1 = index($t650_array,"<tag>",0);
                   $tab2 = index($t650_array,"</tag>",0);
                   }
               else {
                     $tab1 = index($t650_array,"<tag>",($tab2+5));
                     $tab2 = index($t650_array,"</tag>",($tab2+5));
                    }
               $temp = substr($t650_array,($tab1+5),($tab2-($tab1+5)));
               print OUTPUT qq{      <subject encodinganalog="MARC 650">};
               print OUTPUT qq{$temp</subject>\n};
               $tabi++;
               }
        }
    if ($tag656) {
        $tabi = 0;
        while ($tabi < $t656_cnt) {
               if ($tabi == 0) {
                   $tab1 = index($t656_array,"<tag>",0);
                   $tab2 = index($t656_array,"</tag>",0);
                   }
               else {
                     $tab1 = index($t656_array,"<tag>",($tab2+5));
                     $tab2 = index($t656_array,"</tag>",($tab2+5));
                    }
               $temp = substr($t656_array,($tab1+5),($tab2-($tab1+5)));
               print OUTPUT qq{      <subject encodinganalog="MARC 656">};
               print OUTPUT qq{$temp</subject>\n};
               $tabi++;
               }
        }
    print OUTPUT qq{    </controlaccess>\n};
   }
if ($tag651) {
    print OUTPUT qq{    <controlaccess>\n};
    print OUTPUT qq{      <head>Places: </head>\n};
    $tabi = 0;
    while ($tabi < $t651_cnt) {
           if ($tabi == 0) {
               $tab1 = index($t651_array,"<tag>",0);
               $tab2 = index($t651_array,"</tag>",0);
               }
           else {
                 $tab1 = index($t651_array,"<tag>",($tab2+5));
                 $tab2 = index($t651_array,"</tag>",($tab2+5));
                }
           $temp = substr($t651_array,($tab1+5),($tab2-($tab1+5)));
           print OUTPUT qq{      <geogname encodinganalog="MARC 651">};
           print OUTPUT qq{$temp</geogname>\n};
           $tabi++;
          }
     print OUTPUT qq{    </controlaccess>\n};
    }
if ($tag655) {
    print OUTPUT qq{    <controlaccess>\n};
    print OUTPUT qq{      <head>Form and Genre Terms: </head>\n};
    $tabi = 0;
    while ($tabi < $t655_cnt) {
           if ($tabi == 0) {
               $tab1 = index($t655_array,"<tag>",0);
               $tab2 = index($t655_array,"</tag>",0);
               }
           else {
                 $tab1 = index($t655_array,"<tag>",($tab2+5));
                 $tab2 = index($t655_array,"</tag>",($tab2+5));
                }
           $temp = substr($t655_array,($tab1+5),($tab2-($tab1+5)));
           print OUTPUT qq{      <genreform encodinganalog="MARC 655" };
           print OUTPUT qq{source="aat">$temp</genreform>\n};
           $tabi++;
          }
     print OUTPUT qq{    </controlaccess>\n};
    }

print OUTPUT qq{  </controlaccess>\n};
print OUTPUT qq{    <descgrp>\n};
print OUTPUT qq{           <head id="a8">INFORMATION FOR USERS</head>\n};
if ($tag506) {
    print OUTPUT qq{                  <accessrestrict>\n};
    print OUTPUT qq{                          <head>};
    print OUTPUT qq{Access Restrictions:</head>\n};
    $tabi = 0;
    while ($tabi < $t506_cnt) {
           if ($tabi == 0) {
               $tab1 = index($t506_array,"<tag>",0);
               $tab2 = index($t506_array,"</tag>",0);
               }
           else {
                 $tab1 = index($t506_array,"<tag>",($tab2+5));
                 $tab2 = index($t506_array,"</tag>",($tab2+5));
                }
           $temp = substr($t506_array,($tab1+5),($tab2-($tab1+5)));
    print OUTPUT qq{                                   <p>$temp</p>\n};
           $tabi++;
          }
    print OUTPUT qq{                  </accessrestrict>\n};
    }         
if ($tag540) {
    print OUTPUT qq{                  <userestrict>\n};
    print OUTPUT qq{                          <head>};
    print OUTPUT qq{Restrictions on Use:</head>\n};
    $tabi = 0;
    while ($tabi < $t540_cnt) {
           if ($tabi == 0) {
               $tab1 = index($t540_array,"<tag>",0);
               $tab2 = index($t540_array,"</tag>",0);
               }
           else {
                 $tab1 = index($t540_array,"<tag>",($tab2+5));
                 $tab2 = index($t540_array,"</tag>",($tab2+5));
                }
           $temp = substr($t540_array,($tab1+5),($tab2-($tab1+5)));
    print OUTPUT qq{                                   <p>$temp</p>\n};
           $tabi++;
          }
    print OUTPUT qq{                  </userestrict>\n};
    }    
if ($tag538) {
    print OUTPUT qq{                  <phystech encodinganalog="538">\n};
    print OUTPUT qq{                          <head>};
    print OUTPUT qq{Physical Characteristics and Technical Requirements:};
    print OUTPUT qq{</head>\n};
    $tabi = 0;
    while ($tabi < $t538_cnt) {
           if ($tabi == 0) {
               $tab1 = index($t538_array,"<tag>",0);
               $tab2 = index($t538_array,"</tag>",0);
               }
           else {
                 $tab1 = index($t538_array,"<tag>",($tab2+5));
                 $tab2 = index($t538_array,"</tag>",($tab2+5));
                }
           $temp = substr($t538_array,($tab1+5),($tab2-($tab1+5)));
    print OUTPUT qq{                                   <p>$temp</p>\n};
           $tabi++;
          }
    print OUTPUT qq{                  </userestrict>\n};
    }    
     
if ($tag524) {
    print OUTPUT qq{                  <prefercite>\n};
    print OUTPUT qq{                          <head>Cite As:</head>\n};
    print OUTPUT qq{                                   <p>$t524a\n};
    print OUTPUT qq{                                   </p>\n};
    print OUTPUT qq{                  </prefercite>\n};
    }         
print OUTPUT qq{    </descgrp>\n};
if ($tag544) {
    print OUTPUT qq{             <relatedmaterial>\n};
    print OUTPUT qq{                    <head id="a5">};
    print OUTPUT qq{RELATED MATERIALS</head>\n};
    $tabi = 0;
    while ($tabi < $t544_cnt) {
           if ($tabi == 0) {
               $tab1 = index($t544_array,"<tag>",0);
               $tab2 = index($t544_array,"</tag>",0);
               }
           else {
                 $tab1 = index($t544_array,"<tag>",($tab2+5));
                 $tab2 = index($t544_array,"</tag>",($tab2+5));
                }
           $temp = substr($t544_array,($tab1+5),($tab2-($tab1+5)));
    print OUTPUT qq{                                   <p>$temp</p>\n};
           $tabi++;
          }
    print OUTPUT qq{             </relatedmaterial>\n};
   }
if ($tag500) {
    print OUTPUT qq{             <odd type="notes">\n};
    print OUTPUT qq{                    <head id="a6">NOTES</head>\n};
    $tabi = 0;
    while ($tabi < $t500_cnt) {
           if ($tabi == 0) {
               $tab1 = index($t500_array,"<tag>",0);
               $tab2 = index($t500_array,"</tag>",0);
               }
           else {
                 $tab1 = index($t500_array,"<tag>",($tab2+5));
                 $tab2 = index($t500_array,"</tag>",($tab2+5));
                }
           $temp = substr($t500_array,($tab1+5),($tab2-($tab1+5)));
           print OUTPUT qq{                            <p>$temp</p>\n};
           $tabi++;
          }
     print OUTPUT qq{             </odd>\n};
   }
if ($tag351) {
    print OUTPUT qq{             <arrangement encodinganalog="MARC 351\$a">\n};
    print OUTPUT qq{                    <head id="a4">};
    print OUTPUT qq{COLLECTION ARRANGEMENT</head>\n};
    print OUTPUT qq{                            <p>$t351a\n};
    print OUTPUT qq{                            </p>\n};   
    print OUTPUT qq{             </arrangement>\n};
   }
if ($series_flag) {
    open (EXCEL, "$upload_dir/$excel");
    my $line2= "";
    print OUTPUT qq{             <arrangement>\n};
    print OUTPUT qq{             <head>SERIES LIST</head>\n};
    my $c0_flag = 0;
    my $tt_flag = 0;
    while ($line2 = <EXCEL>) {
          if ($line2 =~ /</) {
             if ($line2 =~ /<series>/) {
                 $tab1 = index($line2,"<series>",0);
                 $tab2 = index($line2,"</series>",0);
                 $series = substr($line2,($tab1+8),($tab2-($tab1+8)));
                }
             if ($line2 =~ /<c0level>/) {
                 $tab1 = index($line2,"<c0level>",0);
                 $tab2 = index($line2,"</c0level>",0);
                 $c0level = substr($line2,($tab1+9),($tab2-($tab1+9)));
                 $c0_flag = 1;
                }
             if ($line2 =~ /<title>/) {
                 $tab1 = index($line2,"<title>",0);
                 $tab2 = index($line2,"</title>",0);
                 $ttitle = substr($line2,($tab1+7),($tab2-($tab1+7)));
                 $tt_flag = 1;
                }
              if (($c0_flag) && ($c0level eq "1")) {
                   print OUTPUT qq{             <p> <emph render="bold">\n};
                   print OUTPUT qq{             <ref target="s};
                   print OUTPUT qq{$series};
                   print OUTPUT qq{" show="replace" actuate="onrequest">Series };
                   my $roman = Roman($series);
                   print OUTPUT qq{$roman};
                   print OUTPUT qq{. };
                   if ($tt_flag) {
                       print OUTPUT qq{$ttitle};
                      }
                   print OUTPUT qq{</ref> </emph> <lb/> </p>\n};
                   $tt_flag = 0;
                   $c0_flag = 0;
                  }
              } 
          }
    close(EXCEL); 
    print OUTPUT qq{             </arrangement>\n};
    }
print OUTPUT qq{    <dsc type="combined">\n};
print OUTPUT qq{      <head>CONTAINER LIST</head>\n};
open (EXCEL, "$upload_dir/$excel");
my $line2= "";
my $c01_flag = 0;
my $c02_flag = 0;
my $c03_flag = 0;
my $c04_flag = 0;
my $c0level_flag = 0;
my $type_flag = 0;
my $box_flag = 0;
my $folder_flag = 0;
my $unitid_flag = 0;
my $circa_flag = 0;
my $datefrom_flag = 0;
my $dateto_flag = 0;
my $express_flag = 0;
my $title_flag = 0;
my $creatcorp_flag = 0;
my $creatpers_flag = 0;
my $dims_flag = 0;
my $genre_flag = 0;
my $extent_flag = 0;
my $physdes_flag = 0;
my $scope_flag = 0;
my $arrange_flag = 0;
my $process_flag = 0;
my $persname_flag = 0;
my $corpname_flag = 0;
my $geogname_flag = 0;
my $subject_flag = 0;
my $access_flag = 0;
my $userest_flag = 0;
my $odd_flag = 0;
while ($line2 = <EXCEL>) {
       if ($line2 =~ /</) {
           if ($line2 =~ /<series>/) {
               $tab1 = index($line2,"<series>",0);
               $tab2 = index($line2,"</series>",0);
               $series = substr($line2,($tab1+8),($tab2-($tab1+8)));
               if ($series ne "") {
                   $series_flag = 1;
                  }
               else {
                     $series_flag = 0;
                    }
              }
           if ($line2 =~ /<c0level>/) {
               $tab1 = index($line2,"<c0level>",0);
               $tab2 = index($line2,"</c0level>",0);
               $c0level = substr($line2,($tab1+9),($tab2-($tab1+9)));
               if ($c0level ne "") {
                   $c0level_flag = 1;
                  }
               else {
                     $c0level_flag = 0;
                    }
              }
           if ($line2 =~ /<type>/) {
               $tab1 = index($line2,"<type>",0);
               $tab2 = index($line2,"</type>",0);
               $type = substr($line2,($tab1+6),($tab2-($tab1+6)));
               if ($type ne "") {
                   $type_flag = 1;
                  }
               else {
                     $type_flag = 0;
                    }
              }
           if ($line2 =~ /<box>/) {
               $tab1 = index($line2,"<box>",0);
               $tab2 = index($line2,"</box>",0);
               $box = substr($line2,($tab1+5),($tab2-($tab1+5)));
               if ($box ne "") {
                   $box_flag = 1;
                  }
               else {
                     $box_flag = 0;
                    }
              }
           if ($line2 =~ /<folder>/) {
               $tab1 = index($line2,"<folder>",0);
               $tab2 = index($line2,"</folder>",0);
               $folder = substr($line2,($tab1+8),($tab2-($tab1+8)));
               if ($folder ne "") {
                   $folder_flag = 1;
                  }
               else {
                     $folder_flag = 0;
                    }
              }
           if ($line2 =~ /<unitid>/) {
               $tab1 = index($line2,"<unitid>",0);
               $tab2 = index($line2,"</unitid>",0);
               $unitid = substr($line2,($tab1+8),($tab2-($tab1+8)));
               if ($unitid ne "") {
                   $unitid_flag = 1;
                  }
               else {
                     $unitid_flag = 0;
                    }
              }
           if ($line2 =~ /<circa>/) {
               $tab1 = index($line2,"<circa>",0);
               $tab2 = index($line2,"</circa>",0);
               $circa = substr($line2,($tab1+7),($tab2-($tab1+7)));
               if ($circa ne "") {
                   $circa_flag = 1;
                  }
               else {
                     $circa_flag = 0;
                    }
              }
           if ($line2 =~ /<datefrom>/) {
               $tab1 = index($line2,"<datefrom>",0);
               $tab2 = index($line2,"</datefrom>",0);
               $datefrom = substr($line2,($tab1+10),($tab2-($tab1+10)));
               if ($datefrom ne "") {
                   $datefrom_flag = 1;
                  }
               else {
                     $datefrom_flag = 0;
                    }
              }
           if ($line2 =~ /<dateto>/) {
               $tab1 = index($line2,"<dateto>",0);
               $tab2 = index($line2,"</dateto>",0);
               $dateto = substr($line2,($tab1+8),($tab2-($tab1+8)));
               if ($dateto ne "") {
                   $dateto_flag = 1;
                  }
               else {
                     $dateto_flag = 0;
                    }
              }
           if ($line2 =~ /<express>/) {
               $tab1 = index($line2,"<express>",0);
               $tab2 = index($line2,"</express>",0);
               $express = substr($line2,($tab1+9),($tab2-($tab1+9)));
               if ($express ne "") {
                   $express_flag = 1;
                  }
               else {
                     $express_flag = 0;
                    }
              }
           if ($line2 =~ /<title>/) {
               $tab1 = index($line2,"<title>",0);
               $tab2 = index($line2,"</title>",0);
               $ttitle = substr($line2,($tab1+7),($tab2-($tab1+7)));
               if ($ttitle ne "") {
                   $title_flag = 1;
                  }
               else {
                     $title_flag = 0;
                    }
              }
           if ($line2 =~ /<creatcorp>/) {
               $tab1 = index($line2,"<creatcorp>",0);
               $tab2 = index($line2,"</creatcorp>",0);
               $creatcorp = substr($line2,($tab1+11),($tab2-($tab1+11)));
               if ($creatcorp ne "") {
                   $creatcorp_flag = 1;
                  }
               else {
                     $creatcorp_flag = 0;
                    }
               }
           if ($line2 =~ /<creatpers>/) {
               $tab1 = index($line2,"<creatpers>",0);
               $tab2 = index($line2,"</creatpers>",0);
               $creatpers = substr($line2,($tab1+11),($tab2-($tab1+11)));
               if ($creatpers ne "") {
                   $creatpers_flag = 1;
                  }
               else {
                     $creatpers_flag = 0;
                    }
               }
           if ($line2 =~ /<dims>/) {
               $tab1 = index($line2,"<dims>",0);
               $tab2 = index($line2,"</dims>",0);
               $dims = substr($line2,($tab1+6),($tab2-($tab1+6)));
               if ($dims ne "") {
                   $dims_flag = 1;
                  }
               else {
                     $dims_flag = 0;
                    }
              }
           if ($line2 =~ /<extent>/) {
               $tab1 = index($line2,"<extent>",0);
               $tab2 = index($line2,"</extent>",0);
               $extent = substr($line2,($tab1+8),($tab2-($tab1+8)));
               if ($extent ne "") {
                   $extent_flag = 1;
                  }
               else {
                     $extent_flag = 0;
                    }
              }
           if ($line2 =~ /<physdes>/) {
               $tab1 = index($line2,"<physdes>",0);
               $tab2 = index($line2,"</physdes>",0);
               $physdes = substr($line2,($tab1+9),($tab2-($tab1+9)));
               if ($physdes ne "") {
                   $physdes_flag = 1;
                  }
               else {
                     $physdes_flag = 0;
                    }
              }
           if ($line2 =~ /<scope>/) {
               $tab1 = index($line2,"<scope>",0);
               $tab2 = index($line2,"</scope>",0);
               $scope = substr($line2,($tab1+7),($tab2-($tab1+7)));
               if ($scope ne "") {
                   $scope_flag = 1;
                  }
               else {
                     $scope_flag = 0;
                    }
               }
           if ($line2 =~ /<arrange>/) {
               $tab1 = index($line2,"<arrange>",0);
               $tab2 = index($line2,"</arrange>",0);
               $arrange = substr($line2,($tab1+9),($tab2-($tab1+9)));
               if ($arrange ne "") {
                   $arrange_flag = 1;
                  }
               else {
                     $arrange_flag = 0;
                    }
              }
           if ($line2 =~ /<process>/) {
               $tab1 = index($line2,"<process>",0);
               $tab2 = index($line2,"</process>",0);
               $process = substr($line2,($tab1+9),($tab2-($tab1+9)));
               if ($process ne "") {
                   $process_flag = 1;
                  }
               else {
                     $process_flag = 0;
                    }
              }
           if ($line2 =~ /<odd>/) {
               $tab1 = index($line2,"<odd>",0);
               $tab2 = index($line2,"</odd>",0);
               $odd = substr($line2,($tab1+5),($tab2-($tab1+5)));
               if ($odd ne "") {
                   $odd_flag = 1;
                  }
               else {
                     $odd_flag = 0;
                    }
              }
           if ($line2 =~ /<persname>/) {
               $tab1 = index($line2,"<persname>",0);
               $tab2 = index($line2,"</persname>",0);
               $persname = substr($line2,($tab1+10),($tab2-($tab1+10)));
               if ($persname ne "") {
                   $persname_flag = 1;
                  }
               else {
                     $persname_flag = 0;
                    }
              }
           if ($line2 =~ /<corpname>/) {
               $tab1 = index($line2,"<corpname>",0);
               $tab2 = index($line2,"</corpname>",0);
               $corpname = substr($line2,($tab1+10),($tab2-($tab1+10)));
               if ($corpname ne "") {
                   $corpname_flag = 1;
                  }
               else {
                     $corpname_flag = 0;
                    }
              }
           if ($line2 =~ /<geogname>/) {
               $tab1 = index($line2,"<geogname>",0);
               $tab2 = index($line2,"</geogname>",0);
               $geogname = substr($line2,($tab1+10),($tab2-($tab1+10)));
               if ($geogname ne "") {
                   $geogname_flag = 1;
                  }
               else {
                     $geogname_flag = 0;
                    }
               }
           if ($line2 =~ /<subject>/) {
               $tab1 = index($line2,"<subject>",0);
               $tab2 = index($line2,"</subject>",0);
               $subhead = substr($line2,($tab1+9),($tab2-($tab1+9)));
               if ($subhead ne "") {
                   $subject_flag = 1;
                  }
               else {
                     $subject_flag = 0;
                    }
               }
           if ($line2 =~ /<genre>/) {
               $tab1 = index($line2,"<genre>",0);
               $tab2 = index($line2,"</genre>",0);
               $genre = substr($line2,($tab1+7),($tab2-($tab1+7)));
               if ($genre ne "") {
                   $genre_flag = 1;
                  }
               else {
                     $genre_flag = 0;
                    }
              }
           if ($line2 =~ /<access>/) {
               $tab1 = index($line2,"<access>",0);
               $tab2 = index($line2,"</access>",0);
               $access = substr($line2,($tab1+8),($tab2-($tab1+8)));
               if ($access ne "") {
                   $access_flag = 1;
                  }
               else {
                     $access_flag = 0;
                    }
              }
           if ($line2 =~ /<userest>/) {
               $tab1 = index($line2,"<userest>",0);
               $tab2 = index($line2,"</userest>",0);
               $userest = substr($line2,($tab1+9),($tab2-($tab1+9)));
               if ($userest ne "") {
                   $userest_flag = 1;
                  }
               else {
                     $userest_flag = 0;
                    }
              }
 #|-----------------------------------------------|
 #| Do we have a C01 Level?                       |
 #|-----------------------------------------------|
           if (($c0level_flag) && ($c0level eq "1")) {
              #|----------------------------------|
              #| Check if we have a previous c01  |
              #|----------------------------------|
               if ($c01_flag) {
                  #|-----------------------------|
                  #| if yes, then do we have a   |
                  #| previous c04?               |
                  #|-----------------------------|
                   if ($c04_flag) {
                      #|-----------------------------------|
                      #| if yes, close out the c04 tag     |
                      #|-----------------------------------|
                       print OUTPUT qq{           </c04>\n};
                       $c04_flag = 0;
                       $c0level_flag = 0;
                      }
                  #|-----------------------------|
                  #| is there a previous c03?    |
                  #|-----------------------------|
                   if ($c03_flag) {
                      #|-----------------------------------|
                      #| if yes, close out the c03 tag     |
                      #|-----------------------------------|
                       print OUTPUT qq{           </c03>\n};
                       $c03_flag = 0;
                       $c0level_flag = 0;
                      }
                  #|-----------------------------|
                  #| is there a previous c02?    |
                  #|-----------------------------|
                   if ($c02_flag) {
                      #|-----------------------------------|
                      #| if yes, close out the c02 tag     |
                      #|-----------------------------------|
                       print OUTPUT qq{           </c02>\n};
                       $c02_flag = 0;
                       $c0level_flag = 0;
                      }
                    print OUTPUT qq{      </c01>\n};
                    $c01_flag = 0;
                    $c0level_flag = 0;
                  }

            #|-----------------------------------------------|
            #| if no previous C01, then start the tag        |
            #| Need to determine, if we are dealing with a   |
            #| "Series", "Subseries" or "file"               |
            #|-----------------------------------------------|
             if ($type_flag) {
                 if (($type eq "SERIES") || ($type eq "Series") || 
                     ($type eq "series")) {
                      print OUTPUT qq{      <c01 level="series">\n};
                      print OUTPUT qq{       <did>\n};
                      $type_flag = 0;
                     }
                 if (($type eq "SUBSERIES") || ($type eq "SubSeries") || 
                     ($type eq "Subseries") || ($type eq "subseries")) {
                      print OUTPUT qq{      <c01 level="subseries">\n};
                      print OUTPUT qq{       <did>\n};
                      $type_flag = 0;
                    }
                 if (($type eq "FILE") || ($type eq "File") || 
                     ($type eq "file")) { 
                      print OUTPUT qq{      <c01 level="file">\n};
                      print OUTPUT qq{       <did>\n};
                      $type_flag = 0;
                    }
                 if (($type eq "ITEM") || ($type eq "Item") || 
                     ($type eq "item")) { 
                      print OUTPUT qq{      <c01 level="item">\n};
                      print OUTPUT qq{       <did>\n};
                      $type_flag = 0;
                    }
                 }
              else {
                    print OUTPUT qq{      <c01 level="file">\n};
                    print OUTPUT qq{       <did>\n};
                   }
              if ($box_flag) {
                  if ($box eq "map-case") {
                      print OUTPUT qq{        <container type="map-case">};
                     }
                  elsif ($box eq "reel") {
                         print OUTPUT qq{        <container type="reel">};
                        }
                  elsif ($box eq "volume") {
                         print OUTPUT qq{        <container type="volume">};
                        }
                  elsif ($box eq "cabinet") {
                         print OUTPUT qq{        <container type="cabinet">};
                        }
                  elsif ($box eq "shelf") {
                         print OUTPUT qq{        <container type="shelf">};
                        }
                  elsif ($box eq "page") {
                         print OUTPUT qq{        <container type="page">};
                        }
                  elsif ($box eq "folio") {
                         print OUTPUT qq{        <container type="folio">};
                        }
                  elsif ($box eq "file-drawer") {
                         print OUTPUT qq{        <container type="file-drawer">};
                        }
                  else {
                        print OUTPUT qq{        <container type="box">};
                       }
                  print OUTPUT qq{$box</container>\n};
                  $box_flag = 0;
                 }
              if ($folder_flag) {
                  print OUTPUT qq{        <container type="folder">};
                  print OUTPUT qq{$folder</container>\n};
                  $folder_flag = 0;
                 }
              if ($unitid_flag) {
                  print OUTPUT qq{        <unitid>};
                  print OUTPUT qq{$unitid};
                  print OUTPUT qq{</unitid>\n};
                  $unitid_flag = 0;
                 }
               if ($circa_flag) {
                   if (($circa eq "FALSE") || ($circa eq "False") || 
                       ($circa eq "false")) {
                        if ((!$datefrom_flag) && ($dateto_flag)) {
                            #|--------------------------------------|
                            #| Indicates Date is known and singular |
                            #|--------------------------------------|
                             print OUTPUT qq{ <unitdate type="inclusive" };
                             print OUTPUT qq{normal="$dateto">};
                             print OUTPUT qq{$dateto</unitdate>\n};
                             $circa_flag = 0;
                             $dateto_flag = 0;
                           }
                        if (($datefrom_flag) && (!$dateto_flag)) {
                             print OUTPUT qq{ <unitdate type="inclusive" };
                             print OUTPUT qq{normal="$datefrom">};
                             print OUTPUT qq{$datefrom</unitdate>\n};
                             $circa_flag = 0;
                             $datefrom_flag = 0;
                           }
                        if (($datefrom_flag) && ($dateto_flag)) {
                           #|----------------------------------|
                           #| Indicates Date is known and      |
                           #| there is a range                 |
                           #|----------------------------------|
                           if (!$express_flag) {
                              #|-------------------------------|
                              #| if no expression, then create |
                              #| date range                    |
                              #|-------------------------------|
                              print OUTPUT qq{ <unitdate type="inclusive" };
                              print OUTPUT qq{normal="$datefrom/$dateto">};
                              print OUTPUT qq{$datefrom-$dateto</unitdate>\n};
                              $circa_flag = 0;
                              $datefrom_flag = 0;
                              $dateto_flag = 0;
                              }
                          else {
                               #|--------------------------------|
                               #| use expression for date range  |
                               #|--------------------------------|
                                if ($express =~ /,/) {
                                    print OUTPUT qq{ <unitdate type="bulk" };
                                   }
                                else {
                                      print OUTPUT qq{ <unitdate type="inclusive" };
                                     }
                                 print OUTPUT qq{normal="$datefrom/$dateto">};
                                 print OUTPUT qq{$express</unitdate>\n};
                                 $circa_flag = 0;
                                 $datefrom_flag = 0;
                                 $dateto_flag = 0;
                                 $express_flag = 0;
                               }
                         }     
                      } 
                   if (($circa eq "TRUE") || ($circa eq "True") || 
                       ($circa eq "true")) {
                       if ((!$datefrom_flag) && ($dateto_flag)) {
                          #|-----------------------------------|
                          #| Indicates that Date is estimated  |
                          #| and singular                      |
                          #|-----------------------------------|
                          print OUTPUT qq{ <unitdate type="inclusive" };
                          print OUTPUT qq{certainty="Circa" normal="$dateto">};
                          print OUTPUT qq{$dateto</unitdate>\n};
                          $circa_flag = 0;
                          $dateto_flag = 0;
                          }
                       if (($datefrom_flag) && (!$dateto_flag)) {
                          print OUTPUT qq{ <unitdate type="inclusive" };
                          print OUTPUT qq{certainty="Circa" normal="$datefrom">};
                          print OUTPUT qq{$datefrom</unitdate>\n};
                          $circa_flag = 0;
                          $datefrom_flag = 0;
                          }
                       if (($datefrom_flag) && ($dateto_flag)) {
                          #|------------------------------------------|
                          #| Indicates Date is known and there is     |
                          #| a range                                  |
                          #|------------------------------------------|
                          print OUTPUT qq{ <unitdate type="inclusive" };
                          print OUTPUT qq{certainty="Circa" normal="$datefrom/$dateto">};
                          print OUTPUT qq{$datefrom-$dateto</unitdate>\n};
                          $circa_flag = 0;
                          $datefrom_flag = 0;
                          $dateto_flag = 0;
                          }
                      }  
                 #|---------------------------------------------|
                 #| circa = Undated                             |
                 #|---------------------------------------------|
                   if (($circa eq "UNDATED") || ($circa eq "Undated") || 
                       ($circa eq "undated")) {
                       print OUTPUT qq{ <unitdate>Undated};
                       print OUTPUT qq{</unitdate>\n};
                       $circa_flag = 0;
                      }
                 #|---------------------------------------------|
                 #| circa = Date Unknown                        |
                 #|---------------------------------------------|
                   if (($circa eq "DATE UNKNOWN") || 
                       ($circa eq "Date Unknown") || 
                       ($circa eq "date unknown")) {
                       print OUTPUT qq{ <unitdate>Date Unknown};
                       print OUTPUT qq{</unitdate>\n};
                       $circa_flag = 0;
                      }
                 }
               if ($title_flag) {
                   print OUTPUT qq{        <unittitle>};
                   print OUTPUT qq{$ttitle};
                   print OUTPUT qq{</unittitle>\n};
                   $title_flag = 0;
                  }
               if ($creatpers_flag) {
                   while ($creatpers =~ /\|/){
                          $tab1 = index($creatpers,"|",0);
                          $temp = substr($creatpers,0,$tab1);
                          print OUTPUT qq{       <origination label="Creator">};
                          print OUTPUT qq{<persname>$temp</persname>};
                          print OUTPUT qq{</origination> \n};
                          $creatpers = substr($creatpers,($tab1+1));
                         }
                         print OUTPUT qq{       <origination label="Creator">};
                         print OUTPUT qq{<persname>$creatpers</persname>};
                         print OUTPUT qq{</origination> \n};
                         $creatpers_flag = 0;
                  }
               if ($creatcorp_flag) {
                   while ($creatcorp =~ /\|/){
                          $tab1 = index($creatcorp,"|",0);
                          $temp = substr($creatcorp,0,$tab1);
                          print OUTPUT qq{       <origination label="Creator">};
                          print OUTPUT qq{<corpname>$temp</corpname>};
                          print OUTPUT qq{</origination> \n};
                          $creatcorp = substr($creatcorp,($tab1+1));
                         }
                         print OUTPUT qq{       <origination label="Creator">};
                         print OUTPUT qq{<corpname>$creatcorp</corpname>};
                         print OUTPUT qq{</origination> \n};
                         $creatcorp_flag = 0;
                  }
               if ($dims_flag) {
                   print OUTPUT qq{       <physdesc>};
                   print OUTPUT qq{<dimensions>$dims</dimensions>};
                   print OUTPUT qq{</physdesc>\n};
                   $dims_flag = 0;
                  }
               if ($extent_flag) {
                   print OUTPUT qq{       <physdesc>};
                   print OUTPUT qq{<extent>$extent</extent>};
                   print OUTPUT qq{</physdesc>\n};
                   $extent_flag = 0;
                  }
               if ($physdes_flag) {
                   print OUTPUT qq{       <physdesc>};
                   print OUTPUT qq{$physdes};
                   print OUTPUT qq{</physdesc>\n};
                   $physdes_flag = 0;
                  }
               print OUTPUT qq{       </did>\n};
               if ($scope_flag) {
                   print OUTPUT qq{       <scopecontent>};
                   print OUTPUT qq{<p>$scope</p>};
                   print OUTPUT qq{</scopecontent>\n};
                   $scope_flag = 0;
                  }
               if ($arrange_flag) {
                   print OUTPUT qq{       <arrangement>};
                   print OUTPUT qq{<p>$arrange</p>};
                   print OUTPUT qq{</arrangement>\n};
                   $arrange_flag = 0;
                   }
               if ($process_flag) {
                   print OUTPUT qq{       <processinfo>};
                   print OUTPUT qq{<p>$process</p>};
                   print OUTPUT qq{</processinfo>\n};
                   $process_flag = 0;
                  }
               if ($odd_flag) {
                   print OUTPUT qq{       <odd>};
                   print OUTPUT qq{<p>$odd</p>};
                   print OUTPUT qq{</odd>\n};
                   $odd_flag = 0;
                  }
 #|--------------------------------------------------------|
 #| The fields that appear here can have multiples.  They  |
 #| are separated by "|".                                  |
 #|--------------------------------------------------------|
               if ($persname_flag) {
                   while ($persname =~ /\|/){
                          $tab1 = index($persname,"|",0);
                          $temp = substr($persname,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<persname>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</persname>};
                          print OUTPUT qq{</controlaccess>\n};
                          $persname = substr($persname,($tab1+1));
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<persname>};
                   print OUTPUT qq{$persname};
                   print OUTPUT qq{</persname>};
                   print OUTPUT qq{</controlaccess>\n};
                   $persname_flag = 0;
                  }
               if ($corpname_flag) {
                   while ($corpname =~ /\|/){
                          $tab1 = index($corpname,"|",0);
                          $temp = substr($corpname,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<corpname>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</corpname>};
                          print OUTPUT qq{</controlaccess>\n};
                          $corpname = substr($corpname,($tab1+1));
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<corpname>};
                   print OUTPUT qq{$corpname};
                   print OUTPUT qq{</corpname>};
                   print OUTPUT qq{</controlaccess>\n};
                   $corpname_flag = 0;
                  }
               if ($geogname_flag) {
                   while ($geogname =~ /\|/){
                          $tab1 = index($geogname,"|",0);
                          $temp = substr($geogname,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<geogname>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</geogname>};
                          print OUTPUT qq{</controlaccess>\n};
                          $geogname = substr($geogname,($tab1+1));
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<geogname>};
                   print OUTPUT qq{$geogname};
                   print OUTPUT qq{</geogname>};
                   print OUTPUT qq{</controlaccess>\n};
                   $geogname_flag = 0;
                  }
               if ($subject_flag) {
                   while ($subhead =~ /\|/){
                          $tab1 = index($subhead,"|",0);
                          $temp = substr($subhead,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<subject>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</subject>};
                          print OUTPUT qq{</controlaccess>\n};
                          $subhead = substr($subhead,($tab1+1)); 
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<subject>};
                   print OUTPUT qq{$subhead};
                   print OUTPUT qq{</subject>};
                   print OUTPUT qq{</controlaccess>\n};
                   $subject_flag = 0;
                  }
               if ($genre_flag) {
                   while ($genre =~ /\|/){
                          $tab1 = index($genre,"|",0);
                          $temp = substr($genre,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<genreform>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</genreform>};
                          print OUTPUT qq{</controlaccess>\n};
                          $genre = substr($genre,($tab1+1));
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<genreform>};
                   print OUTPUT qq{$genre};
                   print OUTPUT qq{</genreform>};
                   print OUTPUT qq{</controlaccess>\n};
                   $genre_flag = 0;
                  }
               if ($access_flag) {
                   print OUTPUT qq{       <accessrestrict><p>$access</p></accessrestrict> \n};
                   $access_flag = 0;
                  }
               if ($userest_flag) {
                   print OUTPUT qq{       <userestrict><p>$userest</p></userestrict> \n};
                   $userest_flag = 0;
                  }
               $c01_flag = 1;
              }
 #|-----------------------------------------------|
 #| Do we have a C02 Level?                       |
 #|-----------------------------------------------|
          if (($c0level_flag) && ($c0level eq "2")) {
             #|----------------------------------|
             #| Check if we have a previous c02  |
             #|----------------------------------|
              if ($c02_flag) {
                 #|--------------------------------------------|
                 #| if yes, do we have a previous c04?         |
                 #|--------------------------------------------|
                  if ($c04_flag) {
                     #|----------------------------------------|
                     #| if yes, then close out the c04 tag     |
                     #|----------------------------------------|
                      print OUTPUT qq{                 </c04>\n};
                      $c04_flag = 0;
                     }
                 #|--------------------------------------------|
                 #| do we have a previous c03?                 |
                 #|--------------------------------------------|
                  if ($c03_flag) {
                     #|----------------------------------------|
                     #| if yes, then close out the c03 tag     |
                     #|----------------------------------------|
                      print OUTPUT qq{              </c03>\n};
                      $c03_flag = 0;
                     }
                  print OUTPUT qq{           </c02>\n};
                  $c02_flag = 0;
                 }
            #|-----------------------------------------------|
            #| if no previous C02, then start the tag        |
            #| Need to determine, if we are dealing with a   |
            #| "Series", "Subseries" or "file" or "item"     |
            #|-----------------------------------------------|
             if ($type_flag) {
                 if (($type eq "SUBSERIES") || ($type eq "SubSeries") || 
                     ($type eq "Subseries") || ($type eq "subseries")) {
                      print OUTPUT qq{      <c02 level="subseries">\n};
                      print OUTPUT qq{       <did>\n};
                      $type_flag = 0;
                    }
                 if (($type eq "FILE") || ($type eq "File") || 
                     ($type eq "file")) { 
                      print OUTPUT qq{      <c02 level="file">\n};
                      print OUTPUT qq{       <did>\n};
                      $type_flag = 0;
                    }
                 if (($type eq "ITEM") || ($type eq "Item") || 
                     ($type eq "item")) { 
                      print OUTPUT qq{      <c02 level="item">\n};
                      print OUTPUT qq{       <did>\n};
                      $type_flag = 0;
                    }
                 }
              else {
                    print OUTPUT qq{      <c02 level="file">\n};
                    print OUTPUT qq{       <did>\n};
                   }
              if ($box_flag) {
                  if ($box eq "map-case") {
                      print OUTPUT qq{        <container type="map-case">};
                     }
                  elsif ($box eq "reel") {
                         print OUTPUT qq{        <container type="reel">};
                        }
                  elsif ($box eq "volume") {
                         print OUTPUT qq{        <container type="volume">};
                        }
                  elsif ($box eq "cabinet") {
                         print OUTPUT qq{        <container type="cabinet">};
                        }
                  elsif ($box eq "shelf") {
                         print OUTPUT qq{        <container type="shelf">};
                        }
                  elsif ($box eq "page") {
                         print OUTPUT qq{        <container type="page">};
                        }
                  elsif ($box eq "folio") {
                         print OUTPUT qq{        <container type="folio">};
                        }
                  elsif ($box eq "file-drawer") {
                         print OUTPUT qq{        <container type="file-drawer">};
                        }
                  else {
                        print OUTPUT qq{        <container type="box">};
                       }
                  print OUTPUT qq{$box</container>\n};
                  $box_flag = 0;
                 }
              if ($folder_flag) {
                  print OUTPUT qq{        <container type="folder">};
                  print OUTPUT qq{$folder</container>\n};
                  $folder_flag = 0;
                 }
              if ($unitid_flag) {
                  print OUTPUT qq{        <unitid>};
                  print OUTPUT qq{$unitid};
                  print OUTPUT qq{</unitid>\n};
                  $unitid_flag = 0;
                 }
               if ($circa_flag) {
                   if (($circa eq "FALSE") || ($circa eq "False") || 
                       ($circa eq "false")) {
                        if ((!$datefrom_flag) && ($dateto_flag)) {
                            #|--------------------------------------|
                            #| Indicates Date is known and singular |
                            #|--------------------------------------|
                             print OUTPUT qq{ <unitdate type="inclusive" };
                             print OUTPUT qq{normal="$dateto">};
                             print OUTPUT qq{$dateto</unitdate>\n};
                             $circa_flag = 0;
                             $dateto_flag = 0;
                           }
                        if (($datefrom_flag) && (!$dateto_flag)) {
                             print OUTPUT qq{ <unitdate type="inclusive" };
                             print OUTPUT qq{normal="$datefrom">};
                             print OUTPUT qq{$datefrom</unitdate>\n};
                             $circa_flag = 0;
                             $datefrom_flag = 0;
                           }
                        if (($datefrom_flag) && ($dateto_flag)) {
                           #|----------------------------------|
                           #| Indicates Date is known and      |
                           #| there is a range                 |
                           #|----------------------------------|
                           if (!$express_flag) {
                              #|-------------------------------|
                              #| if no expression, then create |
                              #| date range                    |
                              #|-------------------------------|
                              print OUTPUT qq{ <unitdate type="inclusive" };
                              print OUTPUT qq{normal="$datefrom/$dateto">};
                              print OUTPUT qq{$datefrom-$dateto</unitdate>\n};
                              $circa_flag = 0;
                              $datefrom_flag = 0;
                              $dateto_flag = 0;
                              }
                          else {
                               #|--------------------------------|
                               #| use expression for date range  |
                               #|--------------------------------|
                                if ($express =~ /,/) {
                                    print OUTPUT qq{ <unitdate type="bulk" };
                                   }
                                else {
                                      print OUTPUT qq{ <unitdate type="inclusive" };
                                     }
                                 print OUTPUT qq{normal="$datefrom/$dateto">};
                                 print OUTPUT qq{$express</unitdate>\n};
                                 $circa_flag = 0;
                                 $datefrom_flag = 0;
                                 $dateto_flag = 0;
                                 $express_flag = 0;
                               }
                         }     
                      } 
                   if (($circa eq "TRUE") || ($circa eq "True") || 
                       ($circa eq "true")) {
                       if ((!$datefrom_flag) && ($dateto_flag)) {
                          #|-----------------------------------|
                          #| Indicates that Date is estimated  |
                          #| and singular                      |
                          #|-----------------------------------|
                          print OUTPUT qq{ <unitdate type="inclusive" };
                          print OUTPUT qq{certainty="Circa" normal="$dateto">};
                          print OUTPUT qq{$dateto</unitdate>\n};
                          $circa_flag = 0;
                          $dateto_flag = 0;
                          }
                       if (($datefrom_flag) && (!$dateto_flag)) {
                          print OUTPUT qq{ <unitdate type="inclusive" };
                          print OUTPUT qq{certainty="Circa" normal="$datefrom">};
                          print OUTPUT qq{$datefrom</unitdate>\n};
                          $circa_flag = 0;
                          $datefrom_flag = 0;
                          }
                       if (($datefrom_flag) && ($dateto_flag)) {
                          #|------------------------------------------|
                          #| Indicates Date is known and there is     |
                          #| a range                                  |
                          #|------------------------------------------|
                          print OUTPUT qq{ <unitdate type="inclusive" };
                          print OUTPUT qq{certainty="Circa" normal="$datefrom/$dateto">};
                          print OUTPUT qq{$datefrom-$dateto</unitdate>\n};
                          $circa_flag = 0;
                          $datefrom_flag = 0;
                          $dateto_flag = 0;
                          }
                      }  
                 #|---------------------------------------------|
                 #| circa = Undated                             |
                 #|---------------------------------------------|
                   if (($circa eq "UNDATED") || ($circa eq "Undated") || 
                       ($circa eq "undated")) {
                       print OUTPUT qq{ <unitdate>Undated};
                       print OUTPUT qq{</unitdate>\n};
                       $circa_flag = 0;
                      }
                 #|---------------------------------------------|
                 #| circa = Date Unknown                        |
                 #|---------------------------------------------|
                   if (($circa eq "DATE UNKNOWN") || 
                       ($circa eq "Date Unknown") || 
                       ($circa eq "date unknown")) {
                       print OUTPUT qq{ <unitdate>Date Unknown};
                       print OUTPUT qq{</unitdate>\n};
                       $circa_flag = 0;
                      }
                 }
               if ($title_flag) {
                   print OUTPUT qq{        <unittitle>};
                   print OUTPUT qq{$ttitle};
                   print OUTPUT qq{</unittitle>\n};
                   $title_flag = 0;
                  }
               if ($creatpers_flag) {
                   while ($creatpers =~ /\|/){
                          $tab1 = index($creatpers,"|",0);
                          $temp = substr($creatpers,0,$tab1);
                          print OUTPUT qq{       <origination label="Creator">};
                          print OUTPUT qq{<persname>$temp</persname>};
                          print OUTPUT qq{</origination> \n};
                          $creatpers = substr($creatpers,($tab1+1));
                         }
                         print OUTPUT qq{       <origination label="Creator">};
                         print OUTPUT qq{<persname>$creatpers</persname>};
                         print OUTPUT qq{</origination> \n};
                         $creatpers_flag = 0;
                  }
               if ($creatcorp_flag) {
                   while ($creatcorp =~ /\|/){
                          $tab1 = index($creatcorp,"|",0);
                          $temp = substr($creatcorp,0,$tab1);
                          print OUTPUT qq{       <origination label="Creator">};
                          print OUTPUT qq{<corpname>$temp</corpname>};
                          print OUTPUT qq{</origination> \n};
                          $creatcorp = substr($creatcorp,($tab1+1));
                         }
                         print OUTPUT qq{       <origination label="Creator">};
                         print OUTPUT qq{<corpname>$creatcorp</corpname>};
                         print OUTPUT qq{</origination> \n};
                         $creatcorp_flag = 0;
                  }
               if ($dims_flag) {
                   print OUTPUT qq{       <physdesc>};
                   print OUTPUT qq{<dimensions>$dims</dimensions>};
                   print OUTPUT qq{</physdesc>\n};
                   $dims_flag = 0;
                  }
               if ($extent_flag) {
                   print OUTPUT qq{       <physdesc>};
                   print OUTPUT qq{<extent>$extent</extent>};
                   print OUTPUT qq{</physdesc>\n};
                   $extent_flag = 0;
                  }
               if ($physdes_flag) {
                   print OUTPUT qq{       <physdesc>};
                   print OUTPUT qq{$physdes};
                   print OUTPUT qq{</physdesc>\n};
                   $physdes_flag = 0;
                  }
               print OUTPUT qq{       </did>\n};
               if ($scope_flag) {
                   print OUTPUT qq{       <scopecontent>};
                   print OUTPUT qq{<p>$scope</p>};
                   print OUTPUT qq{</scopecontent>\n};
                   $scope_flag = 0;
                  }
               if ($arrange_flag) {
                   print OUTPUT qq{       <arrangement>};
                   print OUTPUT qq{<p>$arrange</p>};
                   print OUTPUT qq{</arrangement>\n};
                   $process_flag = 0;
                   }
               if ($process_flag) {
                   print OUTPUT qq{       <processinfo>};
                   print OUTPUT qq{$process</p>};
                   print OUTPUT qq{</processinfo>\n};
                   $process_flag = 0;
                  }
               if ($odd_flag) {
                   print OUTPUT qq{       <odd>};
                   print OUTPUT qq{$odd</p>};
                   print OUTPUT qq{</odd>\n};
                   $odd_flag = 0;
                  }
 #|--------------------------------------------------------|
 #| The fields that appear here can have multiples.  They  |
 #| are separated by "|".                                  |
 #|--------------------------------------------------------|
               if ($persname_flag) {
                   while ($persname =~ /\|/){
                          $tab1 = index($persname,"|",0);
                          $temp = substr($persname,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<persname>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</persname>};
                          print OUTPUT qq{</controlaccess>\n};
                          $persname = substr($persname,($tab1+1));
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<persname>};
                   print OUTPUT qq{$persname};
                   print OUTPUT qq{</persname>};
                   print OUTPUT qq{</controlaccess>\n};
                   $persname_flag = 0;
                  }
               if ($corpname_flag) {
                   while ($corpname =~ /\|/){
                          $tab1 = index($corpname,"|",0);
                          $temp = substr($corpname,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<corpname>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</corpname>};
                          print OUTPUT qq{</controlaccess>\n};
                          $corpname = substr($corpname,($tab1+1));
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<corpname>};
                   print OUTPUT qq{$corpname};
                   print OUTPUT qq{</corpname>};
                   print OUTPUT qq{</controlaccess>\n};
                   $corpname_flag = 0;
                  }
               if ($geogname_flag) {
                   while ($geogname =~ /\|/){
                          $tab1 = index($geogname,"|",0);
                          $temp = substr($geogname,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<geogname>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</geogname>};
                          print OUTPUT qq{</controlaccess>\n};
                          $geogname = substr($geogname,($tab1+1));
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<geogname>};
                   print OUTPUT qq{$geogname};
                   print OUTPUT qq{</geogname>};
                   print OUTPUT qq{</controlaccess>\n};
                   $geogname_flag = 0;

                  }
               if ($subject_flag) {
                   while ($subhead =~ /\|/){
                          $tab1 = index($subhead,"|",0);
                          $temp = substr($subhead,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<subject>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</subject>};
                          print OUTPUT qq{</controlaccess>\n};
                          $subhead = substr($subhead,($tab1+1)); 
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<subject>};
                   print OUTPUT qq{$subhead};
                   print OUTPUT qq{</subject>};
                   print OUTPUT qq{</controlaccess>\n};
                   $subject_flag = 0;
                  }
               if ($genre_flag) {
                   while ($genre =~ /\|/){
                          $tab1 = index($genre,"|",0);
                          $temp = substr($genre,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<genreform>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</genreform>};
                          print OUTPUT qq{</controlaccess>\n};
                          $genre = substr($genre,($tab1+1));
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<genreform>};
                   print OUTPUT qq{$genre};
                   print OUTPUT qq{</genreform>};
                   print OUTPUT qq{</controlaccess>\n};
                   $genre_flag = 0;
                  }
               if ($access_flag) {
                   print OUTPUT qq{       <accessrestrict><p>$access</p></accessrestrict> \n};
                   $access_flag = 0;
                  }
               if ($userest_flag) {
                   print OUTPUT qq{       <userestrict><p>$userest</p></userestrict> \n};
                   $userest_flag = 0;
                  }
               $c02_flag = 1;
              }
 #|-----------------------------------------------|
 #| Do we have a C03 Level?                       |
 #|-----------------------------------------------|
          if (($c0level_flag) && ($c0level eq "3")) {
             #|----------------------------------|
             #| Check if we have a previous c03  |
             #|----------------------------------|
              if ($c03_flag) {
                 #|-------------------------------|
                 #| if yes, check if we have a    |
                 #| previous c04 tag              |
                 #|-------------------------------|
                  if ($c04_flag) {
                     #|----------------------------------------|
                     #| if yes, then close c04 tag             |
                     #|----------------------------------------|
                      print OUTPUT qq{                 </c04>\n};
                      $c04_flag = 0;
                     }
                  print OUTPUT qq{              </c03>\n};
                  $c03_flag = 0;
                 }
            #|-----------------------------------------------|
            #| if no previous C03, then start the tag        |
            #|-----------------------------------------------|
             if ($type_flag) {
                 if (($type eq "SUBSERIES") || ($type eq "SubSeries") || 
                     ($type eq "Subseries") || ($type eq "subseries")) {
                      print OUTPUT qq{      <c03 level="subseries">\n};
                      print OUTPUT qq{       <did>\n};
                      $type_flag = 0;
                    }
                 if (($type eq "FILE") || ($type eq "File") || 
                     ($type eq "file")) { 
                      print OUTPUT qq{      <c03 level="file">\n};
                      print OUTPUT qq{       <did>\n};
                      $type_flag = 0;
                    }
                 if (($type eq "ITEM") || ($type eq "Item") || 
                     ($type eq "item")) { 
                      print OUTPUT qq{      <c03 level="item">\n};
                      print OUTPUT qq{       <did>\n};
                      $type_flag = 0;
                    }
                 }
              else {
                    print OUTPUT qq{      <c03 level="file">\n};
                    print OUTPUT qq{       <did>\n};
                   }
              if ($box_flag) {
                  if ($box eq "map-case") {
                      print OUTPUT qq{        <container type="map-case">};
                     }
                  elsif ($box eq "reel") {
                         print OUTPUT qq{        <container type="reel">};
                        }
                  elsif ($box eq "volume") {
                         print OUTPUT qq{        <container type="volume">};
                        }
                  elsif ($box eq "cabinet") {
                         print OUTPUT qq{        <container type="cabinet">};
                        }
                  elsif ($box eq "shelf") {
                         print OUTPUT qq{        <container type="shelf">};
                        }
                  elsif ($box eq "page") {
                         print OUTPUT qq{        <container type="page">};
                        }
                  elsif ($box eq "folio") {
                         print OUTPUT qq{        <container type="folio">};
                        }
                  elsif ($box eq "file-drawer") {
                         print OUTPUT qq{        <container type="file-drawer">};
                        }
                  else {
                        print OUTPUT qq{        <container type="box">};
                       }
                  print OUTPUT qq{$box</container>\n};
                  $box_flag = 0;
                 }
              if ($folder_flag) {
                  print OUTPUT qq{        <container type="folder">};
                  print OUTPUT qq{$folder</container>\n};
                  $folder_flag = 0;
                 }
              if ($unitid_flag) {
                  print OUTPUT qq{        <unitid>};
                  print OUTPUT qq{$unitid};
                  print OUTPUT qq{</unitid>\n};
                  $unitid_flag = 0;
                 }
               if ($circa_flag) {
                   if (($circa eq "FALSE") || ($circa eq "False") || 
                       ($circa eq "false")) {
                        if ((!$datefrom_flag) && ($dateto_flag)) {
                            #|--------------------------------------|
                            #| Indicates Date is known and singular |
                            #|--------------------------------------|
                             print OUTPUT qq{ <unitdate type="inclusive" };
                             print OUTPUT qq{normal="$dateto">};
                             print OUTPUT qq{$dateto</unitdate>\n};
                             $circa_flag = 0;
                             $dateto_flag = 0;
                           }
                        if (($datefrom_flag) && (!$dateto_flag)) {
                             print OUTPUT qq{ <unitdate type="inclusive" };
                             print OUTPUT qq{normal="$datefrom">};
                             print OUTPUT qq{$datefrom</unitdate>\n};
                             $circa_flag = 0;
                             $datefrom_flag = 0;
                           }
                        if (($datefrom_flag) && ($dateto_flag)) {
                           #|----------------------------------|
                           #| Indicates Date is known and      |
                           #| there is a range                 |
                           #|----------------------------------|
                           if (!$express_flag) {
                              #|-------------------------------|
                              #| if no expression, then create |
                              #| date range                    |
                              #|-------------------------------|
                              print OUTPUT qq{ <unitdate type="inclusive" };
                              print OUTPUT qq{normal="$datefrom/$dateto">};
                              print OUTPUT qq{$datefrom-$dateto</unitdate>\n};
                              $circa_flag = 0;
                              $datefrom_flag = 0;
                              $dateto_flag = 0;
                              }
                          else {
                               #|--------------------------------|
                               #| use expression for date range  |
                               #|--------------------------------|
                                if ($express =~ /,/) {
                                    print OUTPUT qq{ <unitdate type="bulk" };
                                   }
                                else {
                                      print OUTPUT qq{ <unitdate type="inclusive" };
                                     }
                                 print OUTPUT qq{normal="$datefrom/$dateto">};
                                 print OUTPUT qq{$express</unitdate>\n};
                                 $circa_flag = 0;
                                 $datefrom_flag = 0;
                                 $dateto_flag = 0;
                                 $express_flag = 0;
                               }
                         }     
                      } 
                   if (($circa eq "TRUE") || ($circa eq "True") || 
                       ($circa eq "true")) {
                       if ((!$datefrom_flag) && ($dateto_flag)) {
                          #|-----------------------------------|
                          #| Indicates that Date is estimated  |
                          #| and singular                      |
                          #|-----------------------------------|
                          print OUTPUT qq{ <unitdate type="inclusive" };
                          print OUTPUT qq{certainty="Circa" normal="$dateto">};
                          print OUTPUT qq{$dateto</unitdate>\n};
                          $circa_flag = 0;
                          $dateto_flag = 0;
                          }
                       if (($datefrom_flag) && (!$dateto_flag)) {
                          print OUTPUT qq{ <unitdate type="inclusive" };
                          print OUTPUT qq{certainty="Circa" normal="$datefrom">};
                          print OUTPUT qq{$datefrom</unitdate>\n};
                          $circa_flag = 0;
                          $datefrom_flag = 0;
                          }
                       if (($datefrom_flag) && ($dateto_flag)) {
                          #|------------------------------------------|
                          #| Indicates Date is known and there is     |
                          #| a range                                  |
                          #|------------------------------------------|
                          print OUTPUT qq{ <unitdate type="inclusive" };
                          print OUTPUT qq{certainty="Circa" normal="$datefrom/$dateto">};
                          print OUTPUT qq{$datefrom-$dateto</unitdate>\n};
                          $circa_flag = 0;
                          $datefrom_flag = 0;
                          $dateto_flag = 0;
                          }
                      }  
                 #|---------------------------------------------|
                 #| circa = Undated                             |
                 #|---------------------------------------------|
                   if (($circa eq "UNDATED") || ($circa eq "Undated") || 
                       ($circa eq "undated")) {
                       print OUTPUT qq{ <unitdate>Undated};
                       print OUTPUT qq{</unitdate>\n};
                       $circa_flag = 0;
                      }
                 #|---------------------------------------------|
                 #| circa = Date Unknown                        |
                 #|---------------------------------------------|
                   if (($circa eq "DATE UNKNOWN") || 
                       ($circa eq "Date Unknown") || 
                       ($circa eq "date unknown")) {
                       print OUTPUT qq{ <unitdate>Date Unknown};
                       print OUTPUT qq{</unitdate>\n};
                       $circa_flag = 0;
                      }
                 }
               if ($title_flag) {
                   print OUTPUT qq{        <unittitle>};
                   print OUTPUT qq{$ttitle};
                   print OUTPUT qq{</unittitle>\n};
                   $title_flag = 0;
                  }
               if ($creatpers_flag) {
                   while ($creatpers =~ /\|/){
                          $tab1 = index($creatpers,"|",0);
                          $temp = substr($creatpers,0,$tab1);
                          print OUTPUT qq{       <origination label="Creator">};
                          print OUTPUT qq{<persname>$temp</persname>};
                          print OUTPUT qq{</origination> \n};
                          $creatpers = substr($creatpers,($tab1+1));
                         }
                         print OUTPUT qq{       <origination label="Creator">};
                         print OUTPUT qq{<persname>$creatpers</persname>};
                         print OUTPUT qq{</origination> \n};
                         $creatpers_flag = 0;
                  }
               if ($creatcorp_flag) {
                   while ($creatcorp =~ /\|/){
                          $tab1 = index($creatcorp,"|",0);
                          $temp = substr($creatcorp,0,$tab1);
                          print OUTPUT qq{       <origination label="Creator">};
                          print OUTPUT qq{<corpname>$temp</corpname>};
                          print OUTPUT qq{</origination> \n};
                          $creatcorp = substr($creatcorp,($tab1+1));
                         }
                         print OUTPUT qq{       <origination label="Creator">};
                         print OUTPUT qq{<corpname>$creatcorp</corpname>};
                         print OUTPUT qq{</origination> \n};
                         $creatcorp_flag = 0;
                  }
               if ($dims_flag) {
                   print OUTPUT qq{       <physdesc>};
                   print OUTPUT qq{<dimensions>$dims</dimensions>};
                   print OUTPUT qq{</physdesc>\n};
                   $dims_flag = 0;
                  }
               if ($extent_flag) {
                   print OUTPUT qq{       <physdesc>};
                   print OUTPUT qq{<extent>$extent</extent>};
                   print OUTPUT qq{</physdesc>\n};
                   $extent_flag = 0;
                  }
               if ($physdes_flag) {
                   print OUTPUT qq{       <physdesc>};
                   print OUTPUT qq{$physdes};
                   print OUTPUT qq{</physdesc>\n};
                   $physdes_flag = 0;
                  }
               print OUTPUT qq{       </did>\n};
               if ($scope_flag) {
                   print OUTPUT qq{       <scopecontent>};
                   print OUTPUT qq{<p>$scope</p>};
                   print OUTPUT qq{</scopecontent>\n};
                   $scope_flag = 0;
                  }
               if ($arrange_flag) {
                   print OUTPUT qq{       <arrangement>};
                   print OUTPUT qq{<p>$arrange</p>};
                   print OUTPUT qq{</arrangement>\n};
                   $process_flag = 0;
                   }
               if ($process_flag) {
                   print OUTPUT qq{       <processinfo>};
                   print OUTPUT qq{$process</p>};
                   print OUTPUT qq{</processinfo>\n};
                   $process_flag = 0;
                  }
               if ($odd_flag) {
                   print OUTPUT qq{       <odd>};
                   print OUTPUT qq{$odd</p>};
                   print OUTPUT qq{</odd>\n};
                   $odd_flag = 0;
                  }
 #|--------------------------------------------------------|
 #| The fields that appear here can have multiples.  They  |
 #| are separated by "|".                                  |
 #|--------------------------------------------------------|
               if ($persname_flag) {
                   while ($persname =~ /\|/){
                          $tab1 = index($persname,"|",0);
                          $temp = substr($persname,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<persname>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</persname>};
                          print OUTPUT qq{</controlaccess>\n};
                          $persname = substr($persname,($tab1+1));
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<persname>};
                   print OUTPUT qq{$persname};
                   print OUTPUT qq{</persname>};
                   print OUTPUT qq{</controlaccess>\n};
                   $persname_flag = 0;
                  }
               if ($corpname_flag) {
                   while ($corpname =~ /\|/){
                          $tab1 = index($corpname,"|",0);
                          $temp = substr($corpname,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<corpname>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</corpname>};
                          print OUTPUT qq{</controlaccess>\n};
                          $corpname = substr($corpname,($tab1+1));
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<corpname>};
                   print OUTPUT qq{$corpname};
                   print OUTPUT qq{</corpname>};
                   print OUTPUT qq{</controlaccess>\n};
                   $corpname_flag = 0;
                  }
               if ($geogname_flag) {
                   while ($geogname =~ /\|/){
                          $tab1 = index($geogname,"|",0);
                          $temp = substr($geogname,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<geogname>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</geogname>};
                          print OUTPUT qq{</controlaccess>\n};
                          $geogname = substr($geogname,($tab1+1));
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<geogname>};
                   print OUTPUT qq{$geogname};
                   print OUTPUT qq{</geogname>};
                   print OUTPUT qq{</controlaccess>\n};
                   $geogname_flag = 0;

                  }
               if ($subject_flag) {
                   while ($subhead =~ /\|/){
                          $tab1 = index($subhead,"|",0);
                          $temp = substr($subhead,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<subject>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</subject>};
                          print OUTPUT qq{</controlaccess>\n};
                          $subhead = substr($subhead,($tab1+1)); 
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<subject>};
                   print OUTPUT qq{$subhead};
                   print OUTPUT qq{</subject>};
                   print OUTPUT qq{</controlaccess>\n};
                   $subject_flag = 0;
                  }
               if ($genre_flag) {
                   while ($genre =~ /\|/){
                          $tab1 = index($genre,"|",0);
                          $temp = substr($genre,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<genreform>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</genreform>};
                          print OUTPUT qq{</controlaccess>\n};
                          $genre = substr($genre,($tab1+1));
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<genreform>};
                   print OUTPUT qq{$genre};
                   print OUTPUT qq{</genreform>};
                   print OUTPUT qq{</controlaccess>\n};
                   $genre_flag = 0;
                  }
               if ($access_flag) {
                   print OUTPUT qq{       <accessrestrict><p>$access</p></accessrestrict> \n};
                   $access_flag = 0;
                  }
               if ($userest_flag) {
                   print OUTPUT qq{       <userestrict><p>$userest</p></userestrict> \n};
                   $userest_flag = 0;
                  }
               $c03_flag = 1;
              }
 #|-----------------------------------------------|
 #| Do we have a C04 Level?                       |
 #|-----------------------------------------------|
         if (($c0level_flag) && ($c0level eq "4")) {
            #|----------------------------------|
            #| Check if we have a previous c04  |
            #|----------------------------------|
              if ($c04_flag) {
                 #|----------------------------------------|
                 #| if yes, then close c04 tag             |
                 #|----------------------------------------|
                  print OUTPUT qq{                 </c04>\n};
                  $c04_flag = 0;
                 }
            #|-----------------------------------------------|
            #| if no previous C04, then start the tag        |
            #|-----------------------------------------------|
             if ($type_flag) {
                 if (($type eq "FILE") || ($type eq "File") || 
                     ($type eq "file")) { 
                      print OUTPUT qq{      <c04 level="file">\n};
                      print OUTPUT qq{       <did>\n};
                      $type_flag = 0;
                    }
                 if (($type eq "ITEM") || ($type eq "Item") || 
                     ($type eq "item")) { 
                      print OUTPUT qq{      <c04 level="item">\n};
                      print OUTPUT qq{       <did>\n};
                      $type_flag = 0;
                    }
                 }
              else {
                    print OUTPUT qq{      <c04 level="file">\n};
                    print OUTPUT qq{       <did>\n};
                   }
              if ($box_flag) {
                  if ($box eq "map-case") {
                      print OUTPUT qq{        <container type="map-case">};
                     }
                  elsif ($box eq "reel") {
                         print OUTPUT qq{        <container type="reel">};
                        }
                  elsif ($box eq "volume") {
                         print OUTPUT qq{        <container type="volume">};
                        }
                  elsif ($box eq "cabinet") {
                         print OUTPUT qq{        <container type="cabinet">};
                        }
                  elsif ($box eq "shelf") {
                         print OUTPUT qq{        <container type="shelf">};
                        }
                  elsif ($box eq "page") {
                         print OUTPUT qq{        <container type="page">};
                        }
                  elsif ($box eq "folio") {
                         print OUTPUT qq{        <container type="folio">};
                        }
                  elsif ($box eq "file-drawer") {
                         print OUTPUT qq{        <container type="file-drawer">};
                        }
                  else {
                        print OUTPUT qq{        <container type="box">};
                       }
                  print OUTPUT qq{$box</container>\n};
                  $box_flag = 0;
                 }
              if ($folder_flag) {
                  print OUTPUT qq{        <container type="folder">};
                  print OUTPUT qq{$folder</container>\n};
                  $folder_flag = 0;
                 }
              if ($unitid_flag) {
                  print OUTPUT qq{        <unitid>};
                  print OUTPUT qq{$unitid};
                  print OUTPUT qq{</unitid>\n};
                  $unitid_flag = 0;
                 }
               if ($circa_flag) {
                   if (($circa eq "FALSE") || ($circa eq "False") || 
                       ($circa eq "false")) {
                        if ((!$datefrom_flag) && ($dateto_flag)) {
                            #|--------------------------------------|
                            #| Indicates Date is known and singular |
                            #|--------------------------------------|
                             print OUTPUT qq{ <unitdate type="inclusive" };
                             print OUTPUT qq{normal="$dateto">};
                             print OUTPUT qq{$dateto</unitdate>\n};
                             $circa_flag = 0;
                             $dateto_flag = 0;
                           }
                        if (($datefrom_flag) && (!$dateto_flag)) {
                             print OUTPUT qq{ <unitdate type="inclusive" };
                             print OUTPUT qq{normal="$datefrom">};
                             print OUTPUT qq{$datefrom</unitdate>\n};
                             $circa_flag = 0;
                             $datefrom_flag = 0;
                           }
                        if (($datefrom_flag) && ($dateto_flag)) {
                           #|----------------------------------|
                           #| Indicates Date is known and      |
                           #| there is a range                 |
                           #|----------------------------------|
                           if (!$express_flag) {
                              #|-------------------------------|
                              #| if no expression, then create |
                              #| date range                    |
                              #|-------------------------------|
                              print OUTPUT qq{ <unitdate type="inclusive" };
                              print OUTPUT qq{normal="$datefrom/$dateto">};
                              print OUTPUT qq{$datefrom-$dateto</unitdate>\n};
                              $circa_flag = 0;
                              $datefrom_flag = 0;
                              $dateto_flag = 0;
                              }
                          else {
                               #|--------------------------------|
                               #| use expression for date range  |
                               #|--------------------------------|
                                if ($express =~ /,/) {
                                    print OUTPUT qq{ <unitdate type="bulk" };
                                   }
                                else {
                                      print OUTPUT qq{ <unitdate type="inclusive" };
                                     }
                                 print OUTPUT qq{normal="$datefrom/$dateto">};
                                 print OUTPUT qq{$express</unitdate>\n};
                                 $circa_flag = 0;
                                 $datefrom_flag = 0;
                                 $dateto_flag = 0;
                                 $express_flag = 0;
                               }
                         }     
                      } 
                   if (($circa eq "TRUE") || ($circa eq "True") || 
                       ($circa eq "true")) {
                       if ((!$datefrom_flag) && ($dateto_flag)) {
                          #|-----------------------------------|
                          #| Indicates that Date is estimated  |
                          #| and singular                      |
                          #|-----------------------------------|
                          print OUTPUT qq{ <unitdate type="inclusive" };
                          print OUTPUT qq{certainty="Circa" normal="$dateto">};
                          print OUTPUT qq{$dateto</unitdate>\n};
                          $circa_flag = 0;
                          $dateto_flag = 0;
                          }
                       if (($datefrom_flag) && (!$dateto_flag)) {
                          print OUTPUT qq{ <unitdate type="inclusive" };
                          print OUTPUT qq{certainty="Circa" normal="$datefrom">};
                          print OUTPUT qq{$datefrom</unitdate>\n};
                          $circa_flag = 0;
                          $datefrom_flag = 0;
                          }
                       if (($datefrom_flag) && ($dateto_flag)) {
                          #|------------------------------------------|
                          #| Indicates Date is known and there is     |
                          #| a range                                  |
                          #|------------------------------------------|
                          print OUTPUT qq{ <unitdate type="inclusive" };
                          print OUTPUT qq{certainty="Circa" normal="$datefrom/$dateto">};
                          print OUTPUT qq{$datefrom-$dateto</unitdate>\n};
                          $circa_flag = 0;
                          $datefrom_flag = 0;
                          $dateto_flag = 0;
                          }
                      }  
                 #|---------------------------------------------|
                 #| circa = Undated                             |
                 #|---------------------------------------------|
                   if (($circa eq "UNDATED") || ($circa eq "Undated") || 
                       ($circa eq "undated")) {
                       print OUTPUT qq{ <unitdate>Undated};
                       print OUTPUT qq{</unitdate>\n};
                       $circa_flag = 0;
                      }
                 #|---------------------------------------------|
                 #| circa = Date Unknown                        |
                 #|---------------------------------------------|
                   if (($circa eq "DATE UNKNOWN") || 
                       ($circa eq "Date Unknown") || 
                       ($circa eq "date unknown")) {
                       print OUTPUT qq{ <unitdate>Date Unknown};
                       print OUTPUT qq{</unitdate>\n};
                       $circa_flag = 0;
                      }
                 }
               if ($title_flag) {
                   print OUTPUT qq{        <unittitle>};
                   print OUTPUT qq{$ttitle};
                   print OUTPUT qq{</unittitle>\n};
                   $title_flag = 0;
                  }
               if ($creatpers_flag) {
                   while ($creatpers =~ /\|/){
                          $tab1 = index($creatpers,"|",0);
                          $temp = substr($creatpers,0,$tab1);
                          print OUTPUT qq{       <origination label="Creator">};
                          print OUTPUT qq{<persname>$temp</persname>};
                          print OUTPUT qq{</origination> \n};
                          $creatpers = substr($creatpers,($tab1+1));
                         }
                         print OUTPUT qq{       <origination label="Creator">};
                         print OUTPUT qq{<persname>$creatpers</persname>};
                         print OUTPUT qq{</origination> \n};
                         $creatpers_flag = 0;
                  }
               if ($creatcorp_flag) {
                   while ($creatcorp =~ /\|/){
                          $tab1 = index($creatcorp,"|",0);
                          $temp = substr($creatcorp,0,$tab1);
                          print OUTPUT qq{       <origination label="Creator">};
                          print OUTPUT qq{<corpname>$temp</corpname>};
                          print OUTPUT qq{</origination> \n};
                          $creatcorp = substr($creatcorp,($tab1+1));
                         }
                         print OUTPUT qq{       <origination label="Creator">};
                         print OUTPUT qq{<corpname>$creatcorp</corpname>};
                         print OUTPUT qq{</origination> \n};
                         $creatcorp_flag = 0;
                  }
               if ($dims_flag) {
                   print OUTPUT qq{       <physdesc>};
                   print OUTPUT qq{<dimensions>$dims</dimensions>};
                   print OUTPUT qq{</physdesc>\n};
                   $dims_flag = 0;
                  }
               if ($extent_flag) {
                   print OUTPUT qq{       <physdesc>};
                   print OUTPUT qq{<extent>$extent</extent>};
                   print OUTPUT qq{</physdesc>\n};
                   $extent_flag = 0;
                  }
               if ($physdes_flag) {
                   print OUTPUT qq{       <physdesc>};
                   print OUTPUT qq{$physdes};
                   print OUTPUT qq{</physdesc>\n};
                   $physdes_flag = 0;
                  }
               print OUTPUT qq{       </did>\n};
               if ($scope_flag) {
                   print OUTPUT qq{       <scopecontent>};
                   print OUTPUT qq{<p>$scope</p>};
                   print OUTPUT qq{</scopecontent>\n};
                   $scope_flag = 0;
                  }
               if ($arrange_flag) {
                   print OUTPUT qq{       <arrangement>};
                   print OUTPUT qq{<p>$arrange</p>};
                   print OUTPUT qq{</arrangement>\n};
                   $process_flag = 0;
                   }
               if ($process_flag) {
                   print OUTPUT qq{       <processinfo>};
                   print OUTPUT qq{$process</p>};
                   print OUTPUT qq{</processinfo>\n};
                   $process_flag = 0;
                  }
               if ($odd_flag) {
                   print OUTPUT qq{       <odd>};
                   print OUTPUT qq{$odd</p>};
                   print OUTPUT qq{</odd>\n};
                   $odd_flag = 0;
                  }
 #|--------------------------------------------------------|
 #| The fields that appear here can have multiples.  They  |
 #| are separated by "|".                                  |
 #|--------------------------------------------------------|
               if ($persname_flag) {
                   while ($persname =~ /\|/){
                          $tab1 = index($persname,"|",0);
                          $temp = substr($persname,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<persname>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</persname>};
                          print OUTPUT qq{</controlaccess>\n};
                          $persname = substr($persname,($tab1+1));
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<persname>};
                   print OUTPUT qq{$persname};
                   print OUTPUT qq{</persname>};
                   print OUTPUT qq{</controlaccess>\n};
                   $persname_flag = 0;
                  }
               if ($corpname_flag) {
                   while ($corpname =~ /\|/){
                          $tab1 = index($corpname,"|",0);
                          $temp = substr($corpname,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<corpname>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</corpname>};
                          print OUTPUT qq{</controlaccess>\n};
                          $corpname = substr($corpname,($tab1+1));
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<corpname>};
                   print OUTPUT qq{$corpname};
                   print OUTPUT qq{</corpname>};
                   print OUTPUT qq{</controlaccess>\n};
                   $corpname_flag = 0;
                  }
               if ($geogname_flag) {
                   while ($geogname =~ /\|/){
                          $tab1 = index($geogname,"|",0);
                          $temp = substr($geogname,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<geogname>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</geogname>};
                          print OUTPUT qq{</controlaccess>\n};
                          $geogname = substr($geogname,($tab1+1));
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<geogname>};
                   print OUTPUT qq{$geogname};
                   print OUTPUT qq{</geogname>};
                   print OUTPUT qq{</controlaccess>\n};
                   $geogname_flag = 0;

                  }
               if ($subject_flag) {
                   while ($subhead =~ /\|/){
                          $tab1 = index($subhead,"|",0);
                          $temp = substr($subhead,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<subject>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</subject>};
                          print OUTPUT qq{</controlaccess>\n};
                          $subhead = substr($subhead,($tab1+1)); 
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<subject>};
                   print OUTPUT qq{$subhead};
                   print OUTPUT qq{</subject>};
                   print OUTPUT qq{</controlaccess>\n};
                   $subject_flag = 0;
                  }
               if ($genre_flag) {
                   while ($genre =~ /\|/){
                          $tab1 = index($genre,"|",0);
                          $temp = substr($genre,0,$tab1);
                          print OUTPUT qq{       <controlaccess>};
                          print OUTPUT qq{<genreform>};
                          print OUTPUT qq{$temp};
                          print OUTPUT qq{</genreform>};
                          print OUTPUT qq{</controlaccess>\n};
                          $genre = substr($genre,($tab1+1));
                         }
                   print OUTPUT qq{       <controlaccess>};
                   print OUTPUT qq{<genreform>};
                   print OUTPUT qq{$genre};
                   print OUTPUT qq{</genreform>};
                   print OUTPUT qq{</controlaccess>\n};
                   $genre_flag = 0;
                  }
               if ($access_flag) {
                   print OUTPUT qq{       <accessrestrict><p>$access</p></accessrestrict> \n};
                   $access_flag = 0;
                  }
               if ($userest_flag) {
                   print OUTPUT qq{       <userestrict><p>$userest</p></userestrict> \n};
                   $userest_flag = 0;
                  }
               $c04_flag = 1;
              }
          }
      }
close(EXCEL);
#|-----------------------------------------|
#| Close out the XML                       |
#|-----------------------------------------| 
if ($c04_flag) {
    #|----------------------------------------|
    #| if yes, then close c04 tag             |
    #|----------------------------------------|
    print OUTPUT qq{                 </c04>\n};
   }
if ($c03_flag) {
    #|----------------------------------------|
    #| if yes, then close out the c03 tag     |
    #|----------------------------------------|
    print OUTPUT qq{           </c03>\n};
   }
if ($c02_flag) {
    #|----------------------------------------|
    #| if yes, then close out the c02 tag     |
    #|----------------------------------------|
    print OUTPUT qq{       </c02>\n};
   }
if ($c01_flag) {
    #|----------------------------------------|
    #| if yes, then close out the c01 tag     |
    #|----------------------------------------|
    print OUTPUT qq{   </c01>\n};
   }
print OUTPUT qq{    </dsc>\n};
print OUTPUT qq{  </archdesc>\n};
print OUTPUT qq{ </ead>\n};
close (OUTPUT);
#print "\nProcessing complete\n";
#|-----------------------------------------|
#| Now build HMTL page                     |
#|-----------------------------------------| 
print "Content-type: text/html\n\n";
print qq{<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"\n};
print qq{    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">\n};
print qq{\n};
print qq{<html xmlns="http://www.w3.org/1999/xhtml">\n};
print qq{\n};
print qq{<head>\n};
print qq{	<title>Convert EXCEL to XML</title>\n};
	print qq{<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />\n};
print qq{	<meta http-equiv="Content-Language" content="en-us" />\n};
print qq{	<link rel="shortcut icon" href="http://collections.library.cornell.edu/archivespace/favicon.ico" type="image/x-icon" />\n};
print qq{	<link rel="stylesheet" type="text/css" media="screen" href="http://collections.library.cornell.edu/archivespace/styles/screen.css" />\n};
print qq{</head>\n};
print qq{<body>\n};
print qq{\n};
print qq{<div id="skipnav">\n};
print qq{	<a href="#content">Skip to main content</a>\n};
print qq{</div>\n};
print qq{\n};
print qq{<hr />\n};
print qq{\n};
print qq{<div id="cu-identity">\n};
print qq{	<div id="cu-logo">\n};
print qq{		<a id="insignia-link" href="http://www.cornell.edu/"><img src="images/Library_2line_4c.gif" alt="Cornell University Library" width="283" height="75" border="0" /></a>\n};
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
print qq{    <h1>Convert EXCEL to EAD XML</h1>\n};
print qq{    <h2>Conversion Complete!</h2>\n};
print qq{<p>Download the EAD:\n};
print qq{<form action="http://collections.library.cornell.edu/cgi-bin/ead_download.cgi"> \n};
print qq{<p>\n};
print qq{  <select size="1" name="ID"> \n};
print qq{  <option value="$id.xml">EAD: $id.xml </option> \n};
print qq{</p>\n};
print qq{  </select> \n};
print qq{<input type="submit" value="Download" name="B1"> \n};
print qq{</form> \n};
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
