#################################################################################################################
#!perl.exe
#
# FILE:         xml2odxe.pl
#
# DESCRIPTION:  Read info from xml and creates corresponding odx-e file
#
# USAGE:        see help_text
#
# COPYRIGHT:   (c) 2020 Robert Bosch GmbH
# HISTORY:
#
# Date         | Author          | Modification
# 28.07.2017   | M.Schönfelder   | Initial version
# 24.09.2018   | M.Kraemer       | remove foot note from file names, e.g. ASS_ODX_E_FILE_x_File and ASS_CRC_SRK_F_1_File
# 04.04.2019   | M.Kraemer       | remove double spaces and trailing whitespaces in write commands
# 19.02.2020   | M.Kraemer       | add entries with value MANUALLY, requested by MFI1-COSE
# 06.03.2020   | M.Kraemer       | add special handling for device conversion e.g. <BoschPrjNameNr>7 503 751 261 => 7 503 751 322</BoschPrjNameNr>
# 19.06.2020   | M.Kraemer       | exit on Error: whitespaces in DBKeyNames are not allowed; add length check for DBKeyNames (Max. length in dataface: 30 char)
# 15.02.2022   | M.Kraemer       | add support for new Format in SOS A => B => C
# 19.05.2022   | M.Kraemer       | Error correction: Format in SOS A => E, B => E, C => E, D => E -> E 4 times in SOS odx-e (only one time expected)
#
# ToBeDone:
# - error handling
#################################################################################################################
use strict;
use Cwd;
use File::Basename;
use XML::Simple;
use Data::Dumper;
use Getopt::Long;

# Global variables
# Usage
my $help_text = "
Usage:
   perl $0 -h|-xml <sos_xml>

    -h                    : Print this usage text.
    -xml <sos_xml>        : SOS xml containing delivery infos
\n";

# other
my $debug     = 0;
my $cur_dir   = getcwd();
   $cur_dir   =~ s/\//\\/g;
my $base_dir  = dirname($0);
my $sos_xml   = "";
my $sos_odx   = "";

#################################################################################################################
# MAIN
#################################################################################################################

# scan arguments
scan_args();

# create object
my $xml = new XML::Simple;

# read XML file
my $data = $xml->XMLin($sos_xml, forcearray => ['BoschPrjNameNr','SW_Versions', 'Product_Info','Product_Info_SplitCol1']);
#print Dumper($data) if $debug;

# create output file in odx-e format
unless (open(ODXFILE, "> $sos_odx ")) {
  print "Could not create $sos_odx!";
  exit(1);
}

# Create SOS ODX-e
# print header (once per file)
print_sos_odx_header($data->{BoschPrjName},$data->{ProjectInfo}{SW_PM});
# write same info for all partnumbers
my $p = 0;
my %PNs=();
while (defined $data->{BoschPrjNameNr}->[$p]) {
   print "Partnumber: $data->{BoschPrjNameNr}->[$p]\n";

   my $prod_num = $data->{BoschPrjNameNr}->[$p];
   $prod_num  =~ s/-//g;
   $prod_num  =~ s/\s*//g;
   $prod_num  =~ s/\.*//g;
   # ~ print "Partnumber: $prod_num\n";
   my $index = length($prod_num)/2;  # at least 2 PN's mentioned
   if ( $index > 6 ) {
      my @words = split /=>/, $prod_num;
      #print Dumper \@words if $debug;
      $prod_num = @words[$#words];  # use last PN
   }
   if( exists($PNs{$prod_num} ) ) {
      print "   -> $prod_num already available. Don't add to odx-e.\n";
      $PNs{$prod_num} ++;
      ++$p;
      next;
   } else {
      $PNs{$prod_num} = 1;
   }

  print_start_config_data($prod_num);

  # go through all rows of sw version table and print info to odx-e
  my $i = 0; # in case of page splitting, more than one section of SW_Versions
  while (defined $data->{SW_Versions}->[$i]) { 
    foreach my $prod (@{$data->{SW_Versions}->[$i]->{Product_Info}}) {
      next if ( ($prod->{Col3} eq "-") || ($prod->{Col3} eq "") ); # do not add to odx-e if no db key given ? TBD ?
      print "\nDBKey:'", $prod->{Col3}, "'\n"; # print it always #if $debug;
      if ($prod->{Col3} =~ / /) {
         print "ERROR: whitespace in DBKey Name!\nPlease correct xml and/or request new DBKey!";
         exit(1);
      }
      $prod->{Col1} =~ s/ \[.*\]//; # remove footnote marking
      $prod->{Col5} =~ s/ \[.*\]//; # remove footnote marking
      $prod->{Col9} =~ s/  / /g; # remove double spaces in write commands
      $prod->{Col9} =~ s/\s+$//; # remove trailing whitespaces
      #todo: remove start command or split it # $prod->{Col8} =~ s/.*  //; # remove CRC Start commands
      # data record: version
      if (  ($prod->{Col6} ne "") && ($prod->{Col6} ne "n.a.") && ($prod->{Col6} ne "-")
         || ($prod->{Col4} ne "") && ($prod->{Col4} ne "n.a.") && ($prod->{Col4} ne "-") ) {
        my $version = $prod->{Col6};
        $version =~ s/\[.*\]//;
        $version =~ s/\s*//g;
        $version = $prod->{Col4} if (($version eq "-") || ($version eq "n.a.") || ($version eq ""));
        print_begin_sdg_config_record($prod->{Col1}." version comparison",$prod->{Col3},$version,"USER-DEFINED");
        print_db_key("DATA-DB-KEY",$prod->{Col3});
        print_end_sdg_config_record();
        print "Data Version (USER-DEFINED): $prod->{Col1}, $version\n" if $debug;
      }
      # data record: filename
      if ( ($prod->{Col5} ne "") && ($prod->{Col5} ne "n.a.") && ($prod->{Col5} ne "-") ) {
        my $dbkeyName = $prod->{Col3}."_File";
        if ( length($dbkeyName) > 30 ) {
           print "ERROR: generated DBKey Name > 30 char '".$dbkeyName."'!\nPlease request new DBKey!\n";
           exit(1);
        }
        print_begin_sdg_config_record($prod->{Col1}." version filename",$dbkeyName,$prod->{Col5},"USER-DEFINED");
        print_db_key("DATA-DB-KEY",$dbkeyName);
        print_end_sdg_config_record();
        print "Data File (USER-DEFINED): $prod->{Col1}, $prod->{Col5}\n" if $debug;
      }
      # data record: read command
      my $res = check_bin_data($prod->{Col7});
      print "Data Diag (BINARY): not valid for odx-e ($prod->{Col7})\n" if ($res);
      if ( ($prod->{Col8} ne "-") && ($prod->{Col7} ne "n.a.") && ($prod->{Col7} ne "") && (not $res) ) {
        if ( length($prod->{Col3}) > 27 ) {
           print "ERROR: generated DBKey Name > 30 char '".$prod->{Col3}."_[DA|RD|WE]'!\nPlease request new DBKey!\n";
           exit(1);
        }
        print_begin_sdg_config_record($prod->{Col3}." version telegram content",$prod->{Col3}."_DA",$prod->{Col7},"BINARY");
        print_db_key("DATA-DB-KEY",$prod->{Col3}."_DA");
        print_db_key("READ-DIAG-COMM",$prod->{Col8});
        print_db_key("READ-DIAG-COMM-DB-KEY",$prod->{Col3}."_RD");
        print "Data Diag (BINARY): $prod->{Col3}\n\t_DA: $prod->{Col7}\n\t_RD: $prod->{Col8}" if $debug;
        # data record: write command
        if ($prod->{Col9} ne "-") {
          print_db_key("WRITE-DIAG-COMM",$prod->{Col9});
          print_db_key("WRITE-DIAG-COMM-DB-KEY",$prod->{Col3}."_WE");
          print ",\n\t_WE: $prod->{Col9}\n" if $debug;
        }
        print_end_sdg_config_record();
        print "\n" if $debug;
      }
    }
    # Checksum info
    foreach my $prod (@{$data->{SW_Versions}->[$i]->{Product_Info_SplitCol1}}) {
      next if ( ($prod->{Col3} eq "-") || ($prod->{Col3} eq "") ); # do not add to odx-e if no db key given ? TBD ?
      # data record: read command
      print "\nDBKey:'", $prod->{Col3}, "'\n"; # print it always #if $debug;
      if ($prod->{Col3} =~ / /) {
         print "ERROR: whitespace in DBKey Name!\nPlease correct xml and/or request new DBKey!";
         exit(1);
      }
      my $res = check_bin_data($prod->{Col7});
      $prod->{Col1} =~ s/ \[.*\]//; # remove footnote marking
      $prod->{Col5} =~ s/ \[.*\]//; # remove footnote marking
      $prod->{Col9} =~ s/  / /g; # remove double spaces in write commands
      $prod->{Col9} =~ s/\s+$//; # remove trailing whitespaces
      #todo: remove start command or split it # $prod->{Col8} =~ s/.*  //; # remove CRC Start commands
      print "Data Diag (BINARY): not valid for odx-e ($prod->{Col7})\n" if ($res);
      if ( ($prod->{Col8} ne "-") && ($prod->{Col7} ne "n.a.") && ($prod->{Col7} ne "") && (not $res) ) {
        if ( length($prod->{Col3}) > 27 ) {
           print "ERROR: generated DBKey Name > 30 char '".$prod->{Col3}."_[DA|RD|WE]'!\nPlease request new DBKey!\n";
           exit(1);
        }
        print_begin_sdg_config_record($prod->{Col3}." version telegram content",$prod->{Col3}."_DA",$prod->{Col7},"BINARY");
        print_db_key("DATA-DB-KEY",$prod->{Col3}."_DA");
        print_db_key("READ-DIAG-COMM",$prod->{Col8});
        print_db_key("READ-DIAG-COMM-DB-KEY",$prod->{Col3}."_RD");
        print "Data Diag (BINARY): $prod->{Col3}\n\t_DA: $prod->{Col7}\n\t_RD: $prod->{Col8}" if $debug;
        # data record: write command
        if ($prod->{Col9} ne "-") {
          print_db_key("WRITE-DIAG-COMM",$prod->{Col9});
          print_db_key("WRITE-DIAG-COMM-DB-KEY",$prod->{Col3}."_WE");
          print ",\n\t_WE: $prod->{Col9}\n" if $debug;
        }
        print_end_sdg_config_record();
        print "\n" if $debug;
      }
    }
    ++$i;
  }

  print_config_end();
  ++$p;
}
my @PN = keys %PNs;
my @count = values %PNs;
for(my $i = 0; $i < @PN; $i++){
   print "\nSummary of $sos_xml\n" if ($i eq 0);
   print "$PN[$i]: $count[$i] entries in SOS xml -> 1 entry in SOS odx-e\n";
}

print_odx_end();
exit(0);    # exit without error

#################################################################################################################
# Subroutines
#################################################################################################################
#################################################################################################################
# print header of SOS (odx)
#
sub print_sos_odx_header
{
  my $project = shift;
  my $sw_pm   = shift;
  # create special time format for odx-e file
  my $cur_time  = get_local_time_cur();
  my (undef,$date_odx,$time_odx) = split(/,/,$cur_time);
  $date_odx =~ s/\s*//g;
  $time_odx =~ s/\s*//g;
  my ($d,$m,$y) = split(/\./,$date_odx);
  $date_odx = "$y-$m-$d";

  print ODXFILE "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n";
  print ODXFILE "<ODX xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xsi:noNamespaceSchemaLocation=\"odx.xsd\" MODEL-VERSION=\"2.2.0\">
\n";
  print ODXFILE "  <ECU-CONFIG ID=\"GUID\">\n";
  print ODXFILE "    <SHORT-NAME>$project</SHORT-NAME>\n";
  print ODXFILE "    <ADMIN-DATA>\n";
  print ODXFILE "      <LANGUAGE>en-US</LANGUAGE>\n";
  print ODXFILE "      <DOC-REVISIONS>\n";
  print ODXFILE "        <DOC-REVISION>\n";
  print ODXFILE "          <TEAM-MEMBER-REF ID-REF=\"_Alliance_IVI_SOS_TeamMemberID\" />\n";
  print ODXFILE "          <REVISION-LABEL>V01.00</REVISION-LABEL>\n";
  print ODXFILE "          <STATE>REL</STATE>\n";
  print ODXFILE "          <DATE>${date_odx}T$time_odx</DATE>\n";
  print ODXFILE "          <MODIFICATIONS>\n";
  print ODXFILE "            <MODIFICATION>\n";
  print ODXFILE "              <CHANGE>Initial version</CHANGE>\n";
  print ODXFILE "            </MODIFICATION>\n";
  print ODXFILE "          </MODIFICATIONS>\n";
  print ODXFILE "        </DOC-REVISION>\n";
  print ODXFILE "      </DOC-REVISIONS>\n";
  print ODXFILE "    </ADMIN-DATA>\n";
  print ODXFILE "\n";
  print ODXFILE "    <COMPANY-DATAS>\n";
  print ODXFILE "      <COMPANY-DATA ID=\"_Alliance_IVI_SOS_CompanyID\">\n";
  print ODXFILE "        <SHORT-NAME>Robert_Bosch_Car_Multimedia_GmbH</SHORT-NAME>\n";
  print ODXFILE "        <LONG-NAME>Robert Bosch Car Multimedia GmbH</LONG-NAME>\n";
  print ODXFILE "        <TEAM-MEMBERS>\n";
  print ODXFILE "          <TEAM-MEMBER ID=\"_Alliance_IVI_SOS_TeamMemberID\">\n";
  # name from SPL template has following format "name surname (department)"
  my ($pm_name, $pm_surname, $pm_dep) = split(/ /,$sw_pm);
  my $pm_shortname = substr($pm_surname,0,1).substr($pm_name,0,1);
  $pm_dep =~ s/[\(\)]//g;
  print ODXFILE "            <SHORT-NAME>$pm_shortname</SHORT-NAME>\n";
  print ODXFILE "            <LONG-NAME>$pm_surname $pm_name</LONG-NAME>\n";
  print ODXFILE "            <DEPARTMENT>$pm_dep</DEPARTMENT>\n";
  print ODXFILE "            <PHONE>TBD</PHONE>\n";
  print ODXFILE "            <EMAIL>$pm_surname\.$pm_name\@de.bosch.com</EMAIL>\n";
  print ODXFILE "          </TEAM-MEMBER>\n";
  print ODXFILE "        </TEAM-MEMBERS>\n";
  print ODXFILE "      </COMPANY-DATA>\n";
  print ODXFILE "    </COMPANY-DATAS>\n";
  print ODXFILE "\n";
  print ODXFILE "    <SDGS>\n";
  print ODXFILE "      <SDG>\n";
  print ODXFILE "        <SDG-CAPTION ID=\"SDG_TYPE_SW_VER_INFO\">\n";
  print ODXFILE "          <SHORT-NAME>SWVersionInfo</SHORT-NAME>\n";
  print ODXFILE "        </SDG-CAPTION>\n";
  print ODXFILE "      </SDG>\n";
  print ODXFILE "    </SDGS>\n";
  print ODXFILE "\n";
  print ODXFILE "    <CONFIG-DATAS>\n";
}
#################################################################################################################
sub print_start_config_data {

  my $prod_num = shift;
  $prod_num  =~ s/-//g;
  $prod_num  =~ s/\s*//g;
  $prod_num  =~ s/\.*//g;
  my $prod_num_short =  $prod_num;
  my $index = length($prod_num_short)/2;  # at least 2 PN's mentioned
  if ( $index > 6 ) {
      my @words = split /=>/, $prod_num;
      #print Dumper \@words if $debug;
      $prod_num_short = @words[$#words];  # use last PN
  }


  print ODXFILE "      <CONFIG-DATA>\n";
  print ODXFILE "        <SHORT-NAME>$prod_num_short</SHORT-NAME>\n";
  print ODXFILE "        <LONG-NAME>$prod_num</LONG-NAME>\n";
  # probably used for odx-e conformity
  #print ODXFILE "        <DESC></DESC>\n"; # not used in CD config
  print ODXFILE "        <VALID-BASE-VARIANTS>\n";
  print ODXFILE "          <VALID-BASE-VARIANT>\n";
  print ODXFILE "            <BASE-VARIANT-SNREF SHORT-NAME=\"s_name\"></BASE-VARIANT-SNREF>\n";
  print ODXFILE "          </VALID-BASE-VARIANT>\n";
  print ODXFILE "        </VALID-BASE-VARIANTS>\n";
  print ODXFILE "        <CONFIG-RECORDS>\n";
}
#################################################################################################################
sub print_config_end {

  print ODXFILE "        </CONFIG-RECORDS>\n";
  print ODXFILE "      </CONFIG-DATA>\n";
}
#################################################################################################################
sub print_odx_end {

  print ODXFILE "    </CONFIG-DATAS>\n";
  print ODXFILE "  </ECU-CONFIG>\n";
  print ODXFILE "</ODX>\n";
}
#################################################################################################################
sub print_end_sdg_config_record {

  print ODXFILE "              </SDG>\n";
  print ODXFILE "            </SDGS>\n";
  print ODXFILE "          </CONFIG-RECORD>\n";
}
#################################################################################################################
sub print_begin_sdg_config_record {

  my $longname    = shift;
  my $db_key      = shift;
  my $data        = shift;
  my $format      = shift;
  my $shortname   = $longname;
     $shortname =~ s/\s/_/g; # no white spaces for shortname
     $shortname =~ s/\W//g;  # only alphanumeric, digits and underscores allowed

  #print ODXFILE "          <!--- $db_key --->\n";
  print ODXFILE "          <CONFIG-RECORD>\n";
  print ODXFILE "            <SHORT-NAME>$shortname</SHORT-NAME>\n";
  print ODXFILE "            <LONG-NAME>$longname</LONG-NAME>\n";
  # needed for odx-e conformity
  print ODXFILE "            <CONFIG-ID-ITEM>\n";
  print ODXFILE "              <SHORT-NAME>1</SHORT-NAME>\n";
  print ODXFILE "              <BYTE-POSITION>0</BYTE-POSITION>\n";
  print ODXFILE "              <DATA-OBJECT-PROP-REF ID-REF=\"\"></DATA-OBJECT-PROP-REF>\n";
  print ODXFILE "            </CONFIG-ID-ITEM>\n";
  print ODXFILE "            <DIAG-COMM-DATA-CONNECTORS>\n";
  print ODXFILE "              <DIAG-COMM-DATA-CONNECTOR>\n";
  print ODXFILE "                <UNCOMPRESSED-SIZE>0</UNCOMPRESSED-SIZE>\n";
  print ODXFILE "                <SOURCE-START-ADDRESS>00000000</SOURCE-START-ADDRESS>\n";
  print ODXFILE "              </DIAG-COMM-DATA-CONNECTOR>\n";
  print ODXFILE "            </DIAG-COMM-DATA-CONNECTORS>\n";
  print ODXFILE "            <DATA-RECORDS>\n";
  print ODXFILE "              <DATA-RECORD DATAFORMAT=\"$format\">\n";
  print ODXFILE "              <SHORT-NAME>$db_key</SHORT-NAME>\n";
  print ODXFILE "                <DATA>$data</DATA>\n";
  print ODXFILE "              </DATA-RECORD>\n";
  print ODXFILE "            </DATA-RECORDS>\n";
  print ODXFILE "            <SDGS>\n";
  print ODXFILE "              <SDG>\n";
  print ODXFILE "                <SDG-CAPTION-REF ID-REF=\"SDG_TYPE_SW_VER_INFO\" />\n";
}
#################################################################################################################
sub print_db_key {

  my $db_key_var  = shift;
  my $db_key_name = shift;

  print ODXFILE "                <SD SI=\"$db_key_var\">$db_key_name</SD>\n";
}
#################################################################################################################
# scan arguments and assign them to global script variables
# show help text if arguments are not set correctly
#
sub scan_args
{
  local * check_file = sub
  {
    my $file = shift;
    
    if (-e $cur_dir."\\".$file) {
      $file = $cur_dir."\\".$file;
    }
    elsif (-e $base_dir."\\".$file) {
      $file = $base_dir."\\".$file;
    }
    elsif (-e $file) { # nothing to be done, setting ok
    }
    else {
      exit(1);
    }
    return $file;
  };

  my $h         = "";
  
  # -h: help;-xml: xml containing sos infos
  
  my $res = GetOptions (
      'h'       => \$h,
      'v'       => \$debug,
      'xml=s'   => \$sos_xml,
  );

  if ($h) {
    print $help_text;
    exit(0);
  }

  if ($sos_xml) {
    $sos_xml  = check_file($sos_xml);
    $sos_odx  = $sos_xml;
    $sos_odx  =~ s/\.xml$/\.odx-e/;
    print("###############\nStart of script\n###############\nCreate ODX-E file with infos from $sos_xml ...\n\n");
  }
  else {
    print "Argument for xml missing\n";
    print $help_text;
    exit(1);
  }
}
#################################################################################################################
sub get_local_time_cur
{
  my ($sec, $min, $hour, $day, $mon, $year, $wday) = (localtime)[0,1,2,3,4,5,6];
  my @wdays = ("Sun","Mon","Tue","Wed","Thu","Fri","Sat");
  my $cur = sprintf "%s, %02d.%02d.%04d, %02d:%02d:%02d",$wdays[$wday],$day,($mon+1),($year+1900),$hour,$min,$sec;
  return($cur);
}
#################################################################################################################
sub check_bin_data
{
  my $value = shift;
  $value =~ s/\s*//g; # ignore spaces
  return 1 if (($value =~ m/[^\dA-F]/) && not ($value =~ m/MANUALLY/)); # value does not contain valid binary data (hex) and is not set to MANUALLY
  return 0;
}