#!/usr/bin/perl -w
use strict;
use Data::Dumper;

#use File::Glob;
sub add_file($);
sub gen_report($);

my $debug = 0;
my @errors;
my @bad_customers;

my %xl_files;

my %required_files =
(
   hc  =>  "^([^~].+)_HC_Excel_",
   srr =>  "^([^~].+)_SRR_Report_"
);

my %templates =
(
   '7mode' => "7mode.docx",
   'Cmode' => "cmode.docx",
);

my @files = <*.xlsx>;

#print "Files are " . join (", ", @files) . "\n";

my $word_obj;
my $excel_obj;
END {
#   $word_obj->Quit();  
}

for my $file (@files) {
  add_file ($file);
}


# Verify that we have complete set of files for each customer   
foreach my $customer (keys %xl_files) {
  for my $f (keys %required_files) {
    if (!$xl_files{$customer}->{$f}) {
      push @errors, "No $f file for customer $customer";
      push @bad_customers, $customer
    }
  } 
}
    
#print Dumper (\%xl_files, \@errors, \@bad_customers);    

for my $c (@bad_customers) {
  delete $xl_files{$c};
}

print Dumper (\%xl_files) if $debug;

for my $c (keys %xl_files) {
   print STDERR "Processing $c...\n";
   gen_report ($c);
}

   
$word_obj->Quit() if ($word_obj);
exit(0);

#========================================================================
# 
#========================================================================
sub add_file($)
{
   my $file = shift;
   
   for my $f (keys %required_files) {
      my $regex = $required_files{$f};
      if ($file =~ /$regex/) {
         my $customer = $1;
         $customer =~ s/_+$//;
         my $t = $xl_files{$customer} ||= {};
         if ($t->{$f}) {
            push @errors, "Duplicate $f files $file, $t->{$f} for customer $customer";
            push @bad_customers, $customer;   
         }
         $t->{$f} = $file;
      }
   }
}        

#/*--------------------------------------------------------------------------*/

use Number::Format qw (format_number);
use Spreadsheet::ParseXLSX;
use FindBin '$Bin';
use Win32::OLE;
use Win32::OLE::Const 'Microsoft.Word';    # wd  constants
#use Win32::OLE::Const 'Microsoft Office';  # mso constants
use Win32::OLE::Const 'Microsoft Office 14.0 Object Library';

use English;

BEGIN {
  Win32::OLE->Option(Warn=>0) unless ($debug);
}

sub get_values($$$$);
sub search($$$$);
sub get_data_from_row($$$&@);

sub get_section_data($$$&@);
sub hash_parser($$@);

# Patch the latest version data to leave only the latest recommendation
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

# Add host reference to disk/shelf/aggr data (to facilitate filtering 7mode/cmode systems)
sub add_host_ref($$$;$)
{
   my ($t, $hosts, $host, $sn) = @_;

   my $h = $hosts->{$host};
   if ($h) {
     if ($sn && $h->{"H:S"} ne $sn) {
       die "Host $host serial number $sn does not match data in host table"
     }
     $t->{hostref} = $h;
   } else { die "Host $host not found in host table"; }
}

# Parse disk data
sub disk_parser()
{
  my ($data, $target, $hosts) = @_; 

  my ($host, undef, $m, $c, $v, $l, $uptd) = @$data;
  my $sn;
  ($host, $sn) = split ("/",$host);
#get_section_data ($firmware, "Disk Firmware Review", 6, \&custom_key_parser, \%disks, \&disk_key_parser, "D:H", undef, "D:M", "D:C", "D:V", "D:L", "D:UPTD");
  my $key = "$host$SUBSEP$m";
  
  my $t = $target->{$key} ||= {};
  
  $t->{'D:H'} = $host;
  $t->{'D:C'} = $c;
  $t->{'D:M'} = $m;
  $t->{'D:V'} = $v;
  $t->{'D:L'} = $l;
  $t->{'D:UPTD'} = $uptd;
  $t->{'D:HOSTSN'} = $sn;
  add_host_ref ($t, $hosts, $host, $sn); 
}


# Parse shelf data
sub shelf_parser()
{
  my ($data, $target, $hosts) = @_; 

  my ($host, $c, $type, $m, $mc, $v, $l, $uptd) = @$data;
  my $sn;
  ($host, $sn) = split ("/",$host);
  #print STDERR "Processing shelf data: " . join (";", @$data) . "\n";
  if ($type eq "0") {
    $type = "Internal";
  }
  my $tm = "$type/$m";
;
  my $key = "$host$SUBSEP$tm";
  
  my $t = $target->{$key} ||= {};
  
  $t->{'S:H'} = $host;
  $t->{'S:T'} = $type;
  $t->{'S:C'} = $c;
  $t->{'S:M'} = $m;
  $t->{'S:MC'} = $mc;
  $t->{'S:V'} = $v;
  $t->{'S:L'} = $l;
  $t->{'S:UPTD'} = $uptd;
  $t->{'S:TM'} = $tm;
  $t->{'S:HOSTSN'} = $sn;

  add_host_ref ($t, $hosts, $host, $sn); 
}

# Format number data for presentation
sub format_num($)
{
   my $num = shift;
   if ($num > 1000) {
      return format_number ($num, 0)
   }
   return format_number ($num, 2);
}

# Parse aggregate data 
sub aggregate_parser($$)
{
  my ($data, $target, $hosts) = @_; 
  
  return if ($data->[0] ne 'aggregate');     

  #print STDERR "Processing aggregate data: " . join (";", @$data) . "\n";
  # Remove some old formatting characters
  for (my $i=0; $i<scalar(@$data); ++$i) {
    $data->[$i] =~ s/^\[.*\]//; 
  }

  my (undef, $inst, $used, $capacity, $growthperday, $cap, $date) = @$data;

  my $key = $inst;
  my ($host, $aggr) = split (":", $inst);
  
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

  $t->{'A:HOST'} = $host;
  $t->{'A:AGGR'} = $aggr;
     
  add_host_ref ($t, $hosts, $host); 
   
}

# Prepares array from an Excel row data. Tries to take into account merged cells
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
      # There is some strange formatted return for numeric data. Use non-formatted data for numerics
      if ($value =~ /^\[.*\]/) {
        $value = $cell->unformatted(); 
      }
      
      if ($cell->is_merged() && $value eq "") {
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

# Search for a text in the workbook $w, starting from row $r, looking at column $c. Return the row number
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
  

# Assume there is a table at row $r, parse the data (blank first column is the end of the table)
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
# gen_report: Main report generation sub
#========================================================================
sub gen_report($)
{
  my $customer = shift;

  my $srr_xl = $xl_files{$customer}->{srr};
  my $hc_xl = $xl_files{$customer}->{hc};
   
  my (%hosts, %shelves, %disks, %aggregates);
  my %tables = (
    H => \%hosts,
    D => \%disks,
    S => \%shelves,
    A => \%aggregates,
  );

  $excel_obj ||= Spreadsheet::ParseXLSX->new;
  print STDERR "Processing $srr_xl...\n";
  my $workbook = $excel_obj->parse($srr_xl);


  my $sysinv = $workbook->worksheet("System Inventory");

  my ($min, $max) = $sysinv->row_range();

  #print "Min is $min, max is $max\n";
  for (my $i=$min; $i<=$max; ++$i ) {
    my $cell = $sysinv->get_cell($i, 0);
    if ($cell && $cell->value() eq "Filer") {
      my ($model, $host, $serial, $os) = get_values ($sysinv, $i, 1, 4);
      $hosts{$host} = {"H:H"=>$host, "H:M"=>$model, "H:S"=>$serial, "H:V"=>$os, "H:MODE"=> ($os =~ /7-Mode/i)? "7mode" : "Cmode"};
      $hosts{$host}->{hostref} = $hosts{$host};
      #print "Data: $model $host $serial $os\n";
    }

  }  

  my $firmware = $workbook->worksheet("OS and Firmware") or die;
  
  get_section_data ($firmware, "Operating System Review", 4, \&hash_parser, \%hosts, undef, "H:OS", "H:L"); 
  get_section_data ($firmware, "Motherboard Firmware Review", 5, \&hash_parser, \%hosts, undef, "H:MBV", "H:MBL", "MBUPTD");
  patch_latest (\%hosts, "H:MBL");

  
  get_section_data ($firmware, "Shelf Firmware Review", 9, \&shelf_parser, \%shelves, \%hosts);
  patch_latest (\%shelves, "S:L");

  get_section_data ($firmware, "Disk Firmware Review", 6, \&disk_parser, \%disks, \%hosts);


  print STDERR "Processing $hc_xl...\n";
  
  my $hc_workbook;
  no warnings; 
  $hc_workbook = $excel_obj->parse($hc_xl);
    
  die "Unable to open/parse file $hc_xl" unless ($hc_workbook);
    
  my $capacity = $hc_workbook->worksheet("Predictive Capacity") or die;

  get_data_from_row ($capacity, 6, 7, \&aggregate_parser, \%aggregates, \%hosts);

  #print Dumper (\%hosts, \%shelves, \%disks, \%aggregates);

  # Generate list of system for each mode (7mode, cmode)
  my %host_mode;
  for my $hostname (keys %hosts) {
    my $host = $hosts{$hostname};
    my $mode = $host->{'H:MODE'};
    $host_mode{$mode} ||= [];
    push @{$host_mode{$mode}}, $hostname;     
  }
  print Dumper (\%tables) if ($debug);
  #exit;
  # Generate a separate report for each mode (7mode, cmode)
  for my $mode (keys %host_mode) {
    my $template = $templates{$mode};
    die "Please specify template file for mode $mode" unless ($template);
    gen_word_output ("${customer}_$mode.docx",  $template, \%tables, $mode);
  }    
}

#========================================================================
# 
#========================================================================
sub grep_mode($$$)
{
   my ($key, $table, $mode) = @_;
   my $entry = $table->{$key};
      
   die unless ($entry);
   my $hostref = $entry->{hostref};
   die "No hostref record" unless ($hostref);
   return ($hostref->{'H:MODE'} eq $mode);
}



#========================================================================
# 
#========================================================================
sub gen_word_output($$$$)
{
  my ($outfile, $template, $data_tables, $mode) =@_;

  print STDERR "Generating $outfile from $template...\n"; 

  $word_obj ||= CreateObject Win32::OLE 'Word.Application' or die $!;
  $word_obj->{'Visible'} = 1;

  my $BaseDir=$Bin;
  my $document = $word_obj->Documents->Open("$BaseDir/$template");
  $document->SaveAs("$BaseDir/$outfile");

  my $tables = $word_obj->ActiveDocument->Tables;

  TABLE:
  for my $table (Win32::OLE::in($tables))
  {
    #print "Processing table...\n";

    my $rows = $table->Rows->{Count};
    #print STDERR "Count of rows is $rows\n";
    my $cols = $table->Columns->{Count};

    for (my $r=1; $r<=$rows; ++$r) {
      my $text = $table->Cell($r,1)->Range->{Text};
      $text =~ s/[[:cntrl:]]+//g;
      print "$r: [$text]\n" if ($debug);
      if ($text =~ /^<(\w+):(\w+)>/) {
        my $tbl = $1;
        my $tag = "$1:$2";
        
        #print "Processing this line with tag $tag...\n";
        my $data_table = $data_tables->{$tbl};
        die "There is no table $tbl" unless ($data_table);
        # Replace the first line, add additional
        my @keys = sort grep {grep_mode($_, $data_table, $mode)} keys %$data_table;
        #print STDERR "Keys for table $tbl: " . join (";", @keys) . "\n";

        my $num_keys = scalar(@keys);
        #print "Number of rows in $tbl table: $num_keys\n";
        #$table->Rows->Item($r)->Select;
        #$word->Selection->Copy;
        #$table->Cell($r,1)->Range->{Text} = "Line 1";
        for (2..$num_keys) {
          #print "Adding row in $tbl table...\n";
          $table->Rows->Item($r)->Select;
          $word_obj->Selection->Copy;
          $word_obj->Selection->PasteAppendTable;
          #$table->Rows->Add($table->Rows($r));
        }
        # Process the added rows
        #print "Processing new rows - keys are " . join ("; ", @keys) . "\n\n";
        for (my $i=0; $i<$num_keys; ++$i) {
          my $key = shift @keys;
          my $data_row = $data_table->{$key};
          for (my $col=1; $col<=$cols; ++$col) {
             my $cell = eval {$table->Cell($r+$i,$col)};
             my $txt = $cell ? $cell->Range->{Text} : "";
             #print "Cell ($i, $col): $txt\n";
             if ($txt =~ /^<CMP>/ && $col > 2) {
                # Special tag to compare two previous columns
                my $t1 = $table->Cell($r+$i,$col-2)->Range->{Text};
                my $t2 = $table->Cell($r+$i,$col-1)->Range->{Text};
                my $t = ($t1 eq $t2) ? "No Action Required" : "See Below";
                #print "Comparing previous cells: $t1 and $t2 - $t\n";
                $table->Cell($r+$i,$col)->Range->{Text} = $t;

             }
                 
             if ($txt =~ /^<(\w+):(\w+)>/) {
               my $tag = "$1:$2";
               #print "Searching for $tag in " . Dumper ($data_row) . "\n";
               my $out = $data_row->{$tag};
               if (defined($out) && $out ne "") {
                  #print "Replacing $txt with $out\n" if $debug;
                  $table->Cell($r+$i,$col)->Range->{Text} = $out;
               }
             }
          }
        #$table->Cell($r,1)->Range->{Text} = "Line 2";
        #last; # Stop processing the currentrows, start processing the next table
        }
        $r += $num_keys;
        $rows += $num_keys;  
        #print "New starting row in the table is $r (of total $rows rows)...\n";
      #next TABLE; 
      }
    }
  }
  
  $document->Save();
  $document->Close();
}
