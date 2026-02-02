#!/usr/bin/perl #-w

#*****************************************************************************#
#FILE: checkIdentsJson.pl
#SW-COMPONENT: Release Tooling
#DESCRIPTION:  Check availablity of config items for given Bosch Part Numbers
#COPYRIGHT: (C) 2023 Robert Bosch GmbH
#
#The reproduction, distribution and utilization of this file as
#well as the communication of its contents to others without express
#authorization is prohibited. Offenders will be held liable for the
#payment of damages. All rights reserved in the event of the grant
#of a patent, utility model or design.
#*****************************************************************************#

my $Filename         = 'checkIdentsJson.pl';
my $Copyright        = '(c) 2023-2025 Robert Bosch GmbH';
my $VERSION          = '1.04';
my $Date             = '13.01.2025';
my $Author           = 'Welz Ralph (XC-CT/ERN7-E)';

#*****************************************************************************#

#------------------------------------------------------------------------------
# DONE
#------------------------------------------------------------------------------
# 1) Define and implement Use Cases:
# a) Output complete tables for all four References for documentation
# b) Output Software Reference for given Software Version (Build Identifier)
# c) Output Hardware Reference for given Hardware Revision (Bosch Part Number)
# d) Output Part Reference for given Hardware Revision and given Software Version
#
# 3) Implement Parameter Handling
#
#------------------------------------------------------------------------------
# TODO
#------------------------------------------------------------------------------
# 2) Test on Linux pending
# 4) Refactor w.r.t. subs
#


#------------------------------------------------------------------------------
# Perl modules used within the script
#------------------------------------------------------------------------------
my $lib;
BEGIN {
  $0   =~ m|(.+\\).+\E?| if ( $^O =~ /mswin/i );
  $0   =~ m|(.+\/).+\E?| if ( $^O =~ /linux/i );
  $lib =  $1;
}
#use lib $lib;                  # not essential, just in case for script local modules		#Nisha commented
use strict;                    # needed
use Cwd;                       # not needed, at least yet
use File::Path;                # needed
use File::Basename;            # needed
use Path::Tiny qw(path);       # needed
use XML::Simple qw(:strict);   # needed
use JSON::MaybeXS;             # needed
use Data::Dumper;              # not essential, but usefull for debugging
use Getopt::Long;              # needed
use Archive::Extract;          # needed
use Benchmark;                 # not essential, just some little benchmarking

#*****************************************************************************#

my $Info = "
Info:
  $Filename $VERSION $Date
  by  $Author
  $Copyright
\n";

my $Synopsis = "
Synopsis:
  Output part reference, hardware reference and software reference by given hardware version and software verson.

  Returns verbal error messages for possible findings.
  Exits with 0 in case of no findings.
  Exits with 1 in case of finding
  Exits with 2 in case of -h
\n";

my $Usage = "
Usage:
  perl $Filename [-h] [-i <idents.json>] [-p <hardware versions,...> -s <software version]

       -h                 : Print this usage text.
       -i <idents.json>   : Full or relative path to a CCS2 idents.json file
       -p <partnums>      : Hardware Versions (Bosch Part Numbers) : comma-separeted list, dot-notation or non-dot-notation
       -s <software>      : Software Version  (Build Identifier)   : ISH.nn.nn.nn or KUM.nn.nn.nn
\n";

my $Example = "
Example:
  perl $Filename -h
  perl $Filename -i idents.json -p 7.515.752.123,7.515.752.456,7.515.752.789 -s ISH.11.04.20
\n";

#*****************************************************************************#

#------------------------------------------------------------------------------
# global variables
#------------------------------------------------------------------------------

# configure behaviour and verbosity of this script
my %VLEVEL                   = ('DEBUG', 1, 'INFO', 2, 'WARNING', 3, 'ERROR', 4);
my $vlevel                   = $VLEVEL{INFO};   # initial verbosity setting before config is read

# other useful variables
my $path_delim = "";
   $path_delim = "\\" if ( $^O =~ /mswin/i );
   $path_delim = "/"  if ( $^O =~ /linux/i );

# main global variables
my $tempdir    = "";
my $id_file    = "";
my $sw_version = "";
my $log_filename = "";
my @partnums   = ();


#==============================================================================
# main script
# (1) exit(0) only in case of success, exit(1) always in case of error, exit(2) for usage
#==============================================================================
my $script_start_time = Benchmark->new();

# parse and evaluate arguments
parse_args();
print ("\n log_filename : $log_filename");
open (STDOUT,'>',$log_filename) or die "\n$log_filename file could not be opened\n";


# read JSON file and decode it into a Perl hash/array structure
my $json_raw = path($id_file)->slurp_utf8;
my $json_obj = JSON::MaybeXS->new();
my $json_dat = $json_obj->decode($json_raw);
#print_info( Dumper($json_dat) );


# create a hash for F013
#   key (condition) : software version (Build Identifier with Wildcard)
#   value			: software reference F013
# TODO: check for doublettes in software version, which should not happen
my %f013;
for ( @{ $json_dat }) {
  my $code   = $json_dat->[0]->{code};
  my $name   = $json_dat->[0]->{name};
  if ( ($code eq "F013") and ($name eq "secondary_operational_reference") ) {
	foreach ( @{$json_dat->[0]->{values}} ) {
	  $f013{$_->{conditions}->[0]->{sw_version}} = $_->{value};
	}
  }
}
if ( $vlevel <= $VLEVEL{DEBUG} ) {
  foreach my $k ( sort keys %f013 ) { print("F013: key: $k val: $f013{$k}\n") }
  #print Dumper (%f013);
};


# create a hash for F191
#   key (condition) : hardware revision  (Bosch Part Number)
#   value			: hardware reference F191
# TODO: check for doublettes in hardware revision, which should not happen
# TODO: condition is currently simplified
my %f191;
for ( @{ $json_dat }) {
  my $code   = $json_dat->[1]->{code};
  my $name   = $json_dat->[1]->{name};
  print_info("code $code");
  print_info("name $name");
  if ( ($code eq "F191") and ($name eq "vehicle_manufacturer_ecu_hardware_number") ) {
	foreach ( @{$json_dat->[1]->{values}} ) {
	  $f191{$_->{conditions}->[0]->{hw_revision}} = $_->{value};
	}
  }
}
if ( $vlevel <= $VLEVEL{DEBUG} ) {
  foreach my $k ( sort keys %f191 ) { print("F191: key: $k val: $f191{$k}\n") };
  #print Dumper (%f191);
}


# create a hash for F1A1
#   key (condition) : software reference F013 AND hardware reference F191
#   value			: part reference F1A1
# TODO: check for doublettes in TBD, which should not happen, but there is the XXX
# TODO: condition is currently simplified
my %f1a1;
for ( @{ $json_dat }) {
  my $code   = $json_dat->[2]->{code};
  my $name   = $json_dat->[2]->{name};
  print_info("code $code");
  print_info("name $name");
  if ( ($code eq "F1A1") and ($name eq "vehicle_manufacturer_spare_part_number_n") ) {
	foreach ( @{$json_dat->[2]->{values}} ) {
	  my $condition = $_->{conditions}->[1]->{did}->{value} . "___" . $_->{conditions}->[0]->{did}->{value};
	  $f1a1{$condition } = $_->{value};
	}
  }
}
if ( $vlevel <= $VLEVEL{DEBUG} ) {
  foreach my $k ( sort keys %f1a1 ) { print("F1A1: key: $k val: $f1a1{$k}\n") };
  #print Dumper (%f1a1);
}


# create a hash for F187 (old reference)
#   key (condition) : software reference F013 AND hardware reference F191
#   value			: part reference F187
# TODO: check for doublettes in TBD, which should not happen, but there is the XXX
# TODO: condition is currently simplified
my %f187;
for ( @{ $json_dat }) {
  my $code   = $json_dat->[3]->{code};
  my $name   = $json_dat->[3]->{name};
  print_info("code $code");
  print_info("name $name");
  if ( ($code eq "F187") and ($name eq "vehicle_manufacturer_spare_part_number_r") ) {
	foreach ( @{$json_dat->[3]->{values}} ) {
	  my $condition = $_->{conditions}->[1]->{did}->{value} . "___" . $_->{conditions}->[0]->{did}->{value};
	  $f187{$condition } = $_->{value};
	}
  }
}
if ( $vlevel <= $VLEVEL{DEBUG} ) {
  foreach my $k ( sort keys %f187 ) { print("F187: key: $k val: $f187{$k}\n") };
  #print Dumper (%f187);
}


# UC 1a) Output complete tables for all four references for documentation
# print("Software Reference:                                SW Ver\n");
# print("===================                                ------\n");
# foreach my $k ( sort keys %f013 ) { print("Reference F013 has value $f013{$k} under condition $k\n") };
# print("\n");
# print("Hardware Reference:                                HW Rev\n");
# print("===================                                ------\n");
# foreach my $k ( sort keys %f191 ) { print("Reference F191 has value $f191{$k} under condition $k\n") };
# print("\n");
# print("Part Reference (new):                              HW Ref       SW Ref\n");
# print("=====================                              ------       ------\n");
# foreach my $k ( sort keys %f1a1 ) { print("Reference F1A1 has value $f1a1{$k} under condition $k\n") };
# print("\n");
# print("Part Reference (old):                              HW Ref       SW Ref\n");
# print("=====================                              ------       ------\n");
# foreach my $k ( sort keys %f187 ) { print("Reference F187 has value $f187{$k} under condition $k\n") };
# print("\n");


# UC 1b) Output Software Reference for given Software Version (Build Identifier)
my $sv = $sw_version;
my @sr = ();
foreach my $k ( sort keys %f013 ) {
  my $k_ = $k;
  $k_ =~ s|\*||;
  if ( $sv =~ m|\Q$k_\E| ) {
    push(@sr,$f013{$k});
  }
}
#print Dumper (@hr);

my $sr_ = @sr;
print("Software Reference:                                SW Ver\n");
print("===================                                ------\n");
if ( $sr_ == 1 ) {
  print("Reference F013 has value $sr[0] under condition $sv\n");
  print("\n");
}
else {
  print_error("Number of Software References for the given Software Version is \'$sr_ to 1'\ instead of \'1 to 1\'!") and exit(1);
}

foreach my $hv ( @partnums ) {
  # UC 1c) Output Hardware Reference for given Hardware Revision (Bosch Part Number)
  my @hr = ();
  foreach my $k ( sort keys %f191 ) {
    if ( $k eq $hv ) {
      push(@hr,$f191{$k});
    }
  }
  #print Dumper (@hr);

  my $hr_ = @hr;
  
  
  
  print("Hardware Reference:                                HW Rev\n"); 
  print("===================                                ------\n");
  if ( $hr_ == 1 ) {
    print("Reference F191 has value $hr[0] under condition $hv\n");
    print("\n");
  }
  else {
    print_error("Number of Hardware References for the given Hardware Version is \'$hr_ to 1'\ instead of \'1 to 1\'!") and exit(1);
  }


  # UC 1d) Output Part Reference for given Hardware Revision and given Software Version
  my $hr = $hr[0];
  my $sr = $sr[0];
  my @pr = ();
  foreach my $k ( sort keys %f1a1 ) {
    if ( ($k =~ m|\Q$hr\E|) and ($k =~ m|\Q$sr\E|) ) {
      push(@pr,$k);
    }
  }
  if ( $#pr ) { # create 'artificial condition'
	my $k = $hr."___".$sr;
	$f1a1{$k} = '<none>    ';
    push(@pr,$k);
  }

  my $pr_ = @pr;
  #print Dumper (@pr);
  print("Part Reference (new):                              HW Ref       SW Ref\n");
  print("=====================                              ------       ------\n");
  if ( $pr_ > 0 ) {
    foreach my $k ( @pr ) { print("Reference F1A1 has value $f1a1{$k} under condition $k\n") };
    print("\n");
  }
  else {
    print_error("Number of Part References for the given Hardware Version is \'0'\ but must be some positive number!") and exit(1);
  }
}


# some benchmarking
my $script_end_time = Benchmark->new();
my $script_execution_time = timediff($script_end_time, $script_start_time);
print_info("Execution time: " . timestr($script_execution_time)) if $vlevel <= $VLEVEL{INFO};


# report success
print_info("SUCCESS");
exit(0);


#------------------------------------------------------------------------------
# Normalize partnumber to the format d.ddd.ddd.ddd[-xx]
# return 0 in case of error
#------------------------------------------------------------------------------
sub normalize_pn {

  my $pn = shift;

  if ( $pn =~ m|(\d{1}\.{0,1}\d{3}\.{0,1}\d{3}\.{0,1}\d{3}(\-[0-9BP][0-9]){0,1})| ) {
    $pn =~ s/[^0-9BP]//g;  # extract all digits from string, including B for BreadBoards and P for PreProd, the normalize
    if    (length($pn) == 10) { $pn = substr($pn,0,1).".".substr($pn,1,3).".".substr($pn,4,3).".".substr($pn,7,3); }
    elsif (length($pn) == 12) { $pn = substr($pn,0,1).".".substr($pn,1,3).".".substr($pn,4,3).".".substr($pn,7,3)."-".substr($pn,10,2); }
    else                      { $pn = 0; }
  }
  else { $pn = 0; }

  return $pn;
}


#------------------------------------------------------------------------------
# parse arguments given to script - uses Getopt::Long
#------------------------------------------------------------------------------
sub parse_args
{
  my $h        = "";
  my $ident    = "";
  my $partnums = "";
  my $software = "";
  my $res      = GetOptions (
        'h'    => \$h,
        'i=s'  => \$ident,
        'p=s'  => \$partnums,
        's=s'  => \$software,
  );

  if ( not $res ) {
    print_error("Please check your command line. An unknown option or switch was given!\n");
    print_info("$Info,$Synopsis,$Usage,$Example\n") and exit(2);
  }

  if ( $h ) {
    print_info("$Info,$Synopsis,$Usage,$Example\n") and exit(2);
  }

  if (-f $ident) {
    $id_file = $ident; # set global variable
  }
  else {
    print_error("No valid idents.json file given: $ident!\n");
    print_info("$Info,$Synopsis,$Usage,$Example\n") and exit(2);
  }

  if ( $partnums ) {
    my @pn = split(/,/, $partnums);
    foreach my $pn (@pn) {
      my $pn_ = normalize_pn($pn);
	     $pn_ =~ s|\.||g;
      if ($pn_) {
        push(@partnums,$pn_); # set global variable
      }
      else {
        print_error("Given item is not a valid Bosch Part Number: $pn!\n");
        print_info("$Info,$Synopsis,$Usage,$Example\n") and exit(2);
      }
    }
  }
  else {
    print_error("No valid (list of) Bosch Part Number(s) given: $partnums!\n");
    print_info("$Info,$Synopsis,$Usage,$Example\n") and exit(2);
  }

  if ( $software ) {
    if ( $software =~ m!(ISH|KUM|NGO|HND)\.\d{2}\.\d{2}\.\d{2}! ) {
      $sw_version = $software; # set global variable
      $log_filename = $software.".txt";
      #print ("\n log_filename : $log_filename");
    }
    else {
      print_error("Given item is not a valid Software Version: $software!\n");
      print_info("$Info,$Synopsis,$Usage,$Example\n") and exit(2);
    }
  }
  else {
    print_error("No valid Software Version given: $software!\n");
    print_info("$Info,$Synopsis,$Usage,$Example\n") and exit(2);
  }
}


#==============================================================================
# status printing on script - mostly borrowed from old cc42.pm
#==============================================================================

#------------------------------------------------------------------------------
# print infos
# 1) the proper formatting of the info is (mostly) up to the using script
#------------------------------------------------------------------------------
BEGIN {
my $ihs = $0;
   $ihs =~ s/.*(\\|\/)//g;
   $ihs =  " ## Info from $ihs: ";
my $ihl =  ' ' x length($ihs);
sub print_info
{
  my $text = shift;

  $text =~ s/\n/\n$ihl/g;
  #print STDOUT "\n$ihs$text\n";
  return "\n$ihs$text\n";
}}


#------------------------------------------------------------------------------
# print warnings
# 1) the proper formatting of the warning is (mostly) up to the using script
#------------------------------------------------------------------------------
BEGIN {
my $whs = $0;
   $whs =~ s/.*(\\|\/)//g;
   $whs =  " ## Warning in $whs: ";
my $whl =  ' ' x length($whs);
sub print_warning
{
  my $text = shift;

  $text =~ s/\n/\n$whl/g;
  print STDOUT "\n$whs$text\n";
  return "\n$whs$text\n";
}}


#------------------------------------------------------------------------------
# print errors
# 1) the proper formatting of the error is (mostly) up to the using script
#------------------------------------------------------------------------------
BEGIN {
my $ehs = $0;
   $ehs =~ s/.*(\\|\/)//g;
   $ehs =  " ## Error in $ehs: ";
my $ehl =  ' ' x length($ehs);
sub print_error
{
  my $text = shift;

  $text =~ s/\n/\n$ehl/g;
  print STDOUT "\n$ehs$text\n";
  return "\n$ehs$text\n";
}}


# EOF
