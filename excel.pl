#!/usr/bin/perl -w
use strict;
use Data::Dumper;
use Number::Format qw (format_number);
my $srr_xl = "WOODARD___CURRAN_SRR_Report_7_3_2016_13_57_51.xlsx";
my $hc_xl = "WOODARD___CURRAN_HC_Excel_7_3_2016_14_01_42.xlsx";
my $word_template = "7mode.doc";
use Spreadsheet::ParseXLSX;
use FindBin '$Bin';
use Win32::OLE;
use Win32::OLE::Const 'Microsoft.Word';    # wd  constants
#use Win32::OLE::Const 'Microsoft Office';  # mso constants
use Win32::OLE::Const 'Microsoft Office 14.0 Object Library';

use English;
#========================================================================
# 
#========================================================================
sub get_values($$$$);
sub search($$$$);
sub get_data_from_row($$$&@);

sub get_section_data($$$&@);
sub hash_parser($$@);

my %hosts;

my $parser = Spreadsheet::ParseXLSX->new;
my $workbook = $parser->parse($srr_xl);


#for my $worksheet ( $workbook->worksheets() ) {
#  print $worksheet->get_name() . "\n"
#}

my $sysinv = $workbook->worksheet("System Inventory");
#print Dumper ($sysinv);
#print $sysinv->get_name() . "\n";

my ($min, $max) = $sysinv->row_range();

print "Min is $min, max is $max\n";
for (my $i=$min; $i<=$max; ++$i ) {
  my $cell = $sysinv->get_cell($i, 0);
  if ($cell && $cell->value() eq "Filer") {
    my ($model, $host, $serial, $os) = get_values ($sysinv, $i, 1, 4);
    $hosts{$host} = {"H:H"=>$host, "H:M"=>$model, "H:S"=>$serial, "H:V"=>$os};
    #print "Data: $model $host $serial $os\n";
  }

}  


my $firmware = $workbook->worksheet("OS and Firmware") or die;

#get_section_data ($sysinv, "Operating System Review", 4, \&hash_parser, \%hosts, undef, "CUROS", "LATESTOS");

get_section_data ($firmware, "Operating System Review", 4, \&hash_parser, \%hosts, undef, "H:OS", "H:L");
get_section_data ($firmware, "Motherboard Firmware Review", 5, \&hash_parser, \%hosts, undef, "H:MBV", "H:MBL", "MBUPTD");

#========================================================================
# 
#========================================================================
sub patch_latest($$)
{
   my ($hash, $tag) = @_;
   for my $key (keys %$hash) {
     my $val = $hash->{$key}->{$tag};
     if ($val) {
       $val = (split (",", $val))[-1];
       $hash->{$key}->{$tag} = $val
     }

   }
}
   
patch_latest (\%hosts, "H:MBL");
   
my %shelves;
#get_section_data ($firmware, "Shelf Firmware Review", 8, \&array_parser, \%shelves, undef, "type", "module", "count", "V", "L", "UPTD");

my %disks;
#get_section_data ($firmware, "Disk Firmware Review", 6, \&array_parser, \%disks, undef, "model", "count", "V", "L", "UPTD");

my %aggregates;

my %tables = (
  H => \%hosts,
  D => \%disks,
  S => \%shelves,
  A => \%aggregates,
);

#========================================================================
# 
#========================================================================
sub disk_key_parser
{
   my $data = shift;
   my ($host, undef, $model) = @$data;
   ($host) = split ("/",$host);
   $data->[0] = $host;     # Replace the hostname in the data array
   return "$host$SUBSEP$model";
}

sub shelf_key_parser
{
   my $data = shift;
   my ($host, undef, $type, $module) = @$data;
   ($host) = split ("/",$host);
   $data->[0] = $host;     # Replace the hostname in the data array
   $data->[8] = "$type/$module";
#   $data->[6] = (split(",", $data->[6]))[-1];      #use only the last recommended firmware;
   return "$host$SUBSEP$type/$module";
}

get_section_data ($firmware, "Shelf Firmware Review", 9, \&custom_key_parser, \%shelves, \&shelf_key_parser, "S:H", "S:C", "S:T", "S:M", "S:MC", "S:V", "S:L", "UPTD", "S:TM");
patch_latest (\%shelves, "S:L");

get_section_data ($firmware, "Disk Firmware Review", 6, \&custom_key_parser, \%disks, \&disk_key_parser, "D:H", undef, "D:M", "D:C", "D:V", "D:L", "D:UPTD");


my $hc_workbook = $parser->parse($hc_xl);
my $capacity = $hc_workbook->worksheet("Predictive Capacity") or die;

#========================================================================
# 
#========================================================================
sub format_num($)
{
   my $num = shift;
   if ($num > 1000) {
      return format_number ($num, 0)
   }
   return format_number ($num, 2);
}

#========================================================================
# 
#========================================================================
sub aggregate_parser($$)
{
  my ($data, $target) = @_; 
  
  return if ($data->[0] ne 'aggregate');     

  print STDERR "Processing aggregate data: " . join (";", @$data) . "\n";
  # Remove some old formatting characters
  for (my $i=0; $i<scalar(@$data); ++$i) {
    $data->[$i] =~ s/^\[.*\]//; 
  }

  my (undef, $inst, $used, $capacity, $growthperday, $cap, $date) = @$data;

  my $key = $inst;
  
  my $t = $target->{$key} ||= {};
  if ($date eq 'n/a') {
    $date = $cap;
  }
  
  $t->{'A:NAME'} = $key;
  $t->{'A:C'} = format_num($capacity);
  $t->{'A:USED'} = format_num($used);
  $t->{'A:RC'} = format_num($capacity - $used);
  $t->{'A:PU'} = format_num(int (($used * 100) / $capacity + 0.5));
  $t->{'A:G'} = format_num($growthperday);
  $t->{'A:EST'} = $date;
     
}

sub aggregate_key_parser
{
   my $data = shift;
   my ($t, $inst, $used, $capacity, $growthperday, $cap, $date) = @$data;
   
   return if ($t ne 'aggregate');     
   print STDERR "Processing aggregate data: " . join (";", @$data) . "\n";
   # Remove some old formatting characters
   for (my $i=0; $i<scalar(@$data); ++$i) {
     $data->[$i] =~ s/^\[.*\]//; 
   }
   return $inst;
}

get_data_from_row ($capacity, 6, 7, \&aggregate_parser, \%aggregates);

#get_data_from_row ($capacity, 6, 7, \&custom_key_parser, \%aggregates, \&aggregate_key_parser, undef, "A:NAME", "A:USED", "A:CAP", "A:G", "A:D");


print Dumper \%hosts;
print Dumper \%shelves;
print Dumper \%disks;
print Dumper \%aggregates;


exit(0);
# Now generate the output report based on the template.

my $word = CreateObject Win32::OLE 'Word.Application' or die $!;
$word->{'Visible'} = 1;

my $BaseDir=$Bin;
my $document = $word->Documents->Open("$BaseDir/$word_template");

my $tables = $word->ActiveDocument->Tables;

for my $table (Win32::OLE::in($tables))
{
   print STDERR "Processing table...\n";

   my $rows = $table->Rows->{Count};
   print STDERR "Count of rows is $rows\n";

   for my $r (1..$rows) {
      my $text = $table->Cell($r,1)->Range->{Text};
      $text =~ s/[[:cntrl:]]+//g;
      print "$r: [$text]\n";
      if ($text =~ /^<([\w:]+)>/) {
        my $tag = $1;
        print "Processing this line with tag $tag...\n";
        #$table->Cell($r,1)->Range->{Text} = "Line 1";
        #$table->Rows->Add($table->Rows($r));
        #$table->Cell($r,1)->Range->{Text} = "Line 2";    
        }
      }
}
#$table->Cell(3,1)->Range->{Text} = "Test string";

   
#save_doc_as($document, 'c:\temp\winword\ttt.doc');

## uncomment the following two if word should shut down
# close_doc($document);
# $word->Quit;





















  
#========================================================================
#
#========================================================================
sub get_values($$$$)
{
  my ($w, $r, $c, $num) = @_;
  
  return unless ( $w );
  my @ret;
  my $merged;
  for (my $i=$c; $num; ++$i) {
    my $value;
    my $cell=$w->get_cell($r,$i);
    if ($cell) {
      $value = $cell->value();
      # There is some strange formatted return for numeric data. Fix it
      if ($value =~ /^\[.*\]/) {
        $value = $cell->unformatted(); 
      }
      
      if ($cell->is_merged() && !$value) {
          #print STDERR "Cell $r, $i is merged and ignored, (value is [$value])\n";
          next;
      }
    }
    --$num;
    push @ret, $value
  }
  
  #print STDERR "Returning " . join (", ", @ret);
  return @ret;
}

#========================================================================
#
#========================================================================
sub search($$$$)
{
  my ($w, $r, $c, $text) = @_;
  
  return unless ($w);
  
  my ($min, $max) = $w->row_range();

  for (my $i=$r; $i<=$max; ++$i ) {
    my $cell = $w->get_cell($i, $c);
    if ($cell && $cell->value() eq $text) {
      return $i;
    }
  }
  return undef;
}  
  

#========================================================================
# 
#========================================================================
sub get_data_from_row($$$&@)
{
  my ($w, $r, $num, $parser_sub, @parser_params) = @_; 
   
    while (1) {
      my @data = get_values($w, $r, 0, $num+1);
      last unless ($data[0]);
      &$parser_sub (\@data, @parser_params);
      ++$r;
    }  
}

#========================================================================
# 
#========================================================================
sub get_section_data($$$&@)
{
  my ($w, $section, $num, $parser_sub, @parser_params) = @_; 
  
  die "parser_sub should be a sub ref" unless (ref($parser_sub) eq "CODE");

  my $r = search ($w, 0, 0, $section);
  if ($r) {
    get_data_from_row ($w, $r+2, $num, \&$parser_sub, @parser_params);
  }
}


#========================================================================
# 
#========================================================================
sub hash_parser($$@)
{
  my ($data, $target, @tags) = @_; 
  
  #print STDERR "Processing: \n" . Dumper ($data) . "\n";
  my $key = $data->[0];
  ($key) = split ("/", $key);
  my $t = $target->{$key} ||= {};
  my $pos = 0;
  foreach my $tag (@tags) {
    ++$pos;
    $t->{$tag} = $data->[$pos] if ($tag);

  }
  #$target->{$key} = $t;
     
}

#========================================================================
# 
#========================================================================
sub array_parser($$@)
{
  my ($data, $target, @tags) = @_; 
  
  print STDERR "Processing: \n" . Dumper ($data) . "\n";
  #print STDERR "Tags are: " . join (",", @tags) . "\n";
  my $key = $data->[0];
  ($key) = split ("/", $key);
  my $t = {};
  my $pos = 0;
  foreach my $tag (@tags) {
    ++$pos;
    $t->{$tag} = $data->[$pos] if ($tag);
    print STDERR "$tag = $data->[$pos] \n";

  }
  $target->{$key} ||= [];
  push @{$target->{$key}}, $t;
   
}  

#========================================================================
#
#========================================================================
sub compound_key_parser($$$@)
{
  my ($data, $target, $key_fields, @tags) = @_; 

  my $key = join ($SUBSEP, @$data[@$key_fields]);
  #print STDERR "Key is $key\nData is " . join(";", @$data) . "\n";
  
  my $t = $target->{$key} ||= {};
  my $pos = 0;
  foreach my $tag (@tags) {
    $t->{$tag} = $data->[$pos] if ($tag);
    ++$pos;

  }
  
}  


#========================================================================
# 
#========================================================================
sub custom_key_parser($$&@)
{
  my ($data, $target, $key_sub, @tags) = @_; 

  die "key_sub need to be a coderef" unless (ref($key_sub) eq "CODE");

  #print STDERR "Input data " . join(";", @$data) . "\n";
  my $key = &$key_sub($data);
  
  return unless ($key); # Skip the line if the key is blank
  #print STDERR "custome_key_parser: Key is $key\nData is " . join(";", @$data) . "\n";
  
  my $t = $target->{$key} ||= {};
  my $pos = 0;
  foreach my $tag (@tags) {
    $t->{$tag} = $data->[$pos] if ($tag);
    ++$pos;

  }
   
}