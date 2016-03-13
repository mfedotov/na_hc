#!/usr/bin/perl -w
use strict;

use warnings;
use strict;
use FindBin '$Bin';
use Win32::OLE;
use Win32::OLE::Const 'Microsoft.Word';    # wd  constants
#use Win32::OLE::Const 'Microsoft Office';  # mso constants
use Win32::OLE::Const 'Microsoft Office 14.0 Object Library';
use Data::Dumper;


my $hosts = {
          'wcusprod2' => {
                         'H:MBV' => '5.2.1',
                         'H:L' => '8.3.2RC2',
                         'H:S' => '700000963674',
                         'H:V' => '8.2.1P3 7-Mode',
                         'H:H' => 'wcusprod2',
                         'MBUPTD' => 'Yes',
                         'H:OS' => "8.2.1P3\x{a0}7-Mode",
                         'H:MBL' => '5.3.0',
                         'H:M' => 'FAS3240'
                       },
          'wcusprod1' => {
                         'H:MBV' => '5.2.1',
                         'H:L' => '8.3.2RC2',
                         'H:S' => '700000963662',
                         'H:V' => '8.2.1P3 7-Mode',
                         'H:H' => 'wcusprod1',
                         'MBUPTD' => 'Yes',
                         'H:OS' => "8.2.1P3\x{a0}7-Mode",
                         'H:MBL' => '5.3.0',
                         'H:M' => 'FAS3240'
                       }
        };
my $shelves = 
{
          'wcusprod1DS4243/IOM3' => {
                                     'S:T' => 'DS4243',
                                     'S:H' => 'wcusprod1',
                                     'S:V' => '0181',
                                     'S:M' => 'IOM3',
                                     'S:MC' => '2',
                                     'S:C' => '1',
                                     'S:L' => '0200',
                                     'S:TM' => 'DS4243/IOM3',
                                     'UPTD' => 'No'
                                   },
          'wcusprod1DS14-Mk2-AT/AT-FCX' => {
                                            'S:T' => 'DS14-Mk2-AT',
                                            'S:H' => 'wcusprod1',
                                            'S:V' => '38',
                                            'S:M' => 'AT-FCX',
                                            'S:MC' => '2',
                                            'S:C' => '1',
                                            'S:L' => '38',
                                            'S:TM' => 'DS14-Mk2-AT/AT-FCX',
                                            'UPTD' => 'Yes'
                                          },
          'wcusprod2DS14-Mk2-AT/AT-FCX' => {
                                            'S:T' => 'DS14-Mk2-AT',
                                            'S:H' => 'wcusprod2',
                                            'S:V' => '38',
                                            'S:M' => 'AT-FCX',
                                            'S:MC' => '2',
                                            'S:C' => '1',
                                            'S:L' => '38',
                                            'S:TM' => 'DS14-Mk2-AT/AT-FCX',
                                            'UPTD' => 'Yes'
                                          },
          'wcusprod2DS4243/IOM3' => {
                                     'S:T' => 'DS4243',
                                     'S:H' => 'wcusprod2',
                                     'S:V' => '0181',
                                     'S:M' => 'IOM3',
                                     'S:MC' => '2',
                                     'S:C' => '1',
                                     'S:L' => '0200',
                                     'S:TM' => 'DS4243/IOM3',
                                     'UPTD' => 'No'
                                   }
        };   
      
         
my $disks = {
          'wcusprod1X411_S15K7420A15' => {
                                          'D:C' => '22',
                                          'D:UPTD' => 'No',
                                          'D:L' => 'NA08',
                                          'D:M' => 'X411_S15K7420A15',
                                          'D:V' => 'NA06',
                                          'D:H' => 'wcusprod1'
                                        },
          'wcusprod1X411_HVIPC420A15' => {
                                          'D:C' => '2',
                                          'D:UPTD' => 'No',
                                          'D:L' => 'NA03',
                                          'D:M' => 'X411_HVIPC420A15',
                                          'D:V' => 'NA02',
                                          'D:H' => 'wcusprod1'
                                        },
          'wcusprod2X269_HJUPI01TSSX' => {
                                          'D:C' => '14',
                                          'D:UPTD' => 'Yes',
                                          'D:L' => 'NA01',
                                          'D:M' => 'X269_HJUPI01TSSX',
                                          'D:V' => 'NA01',
                                          'D:H' => 'wcusprod2'
                                        }
        };
my $aggregates = 
   {
          'wcusprod1:aggr0_64' => {
                                  'A:C' => '7,329',
                                  'A:PU' => '59',
                                  'A:RC' => '3,016',
                                  'A:G' => '2.11',
                                  'A:USED' => '4,313',
                                  'A:EST' => 'More than one year',
                                  'A:NAME' => 'wcusprod1:aggr0_64'
                                },
          'wcusprod2:aggr0_64' => {
                                  'A:C' => '7,784',
                                  'A:PU' => '21',
                                  'A:RC' => '6,135',
                                  'A:G' => '1.73',
                                  'A:USED' => '1,649',
                                  'A:EST' => 'More than one year',
                                  'A:NAME' => 'wcusprod2:aggr0_64'
                                }
        };





        
        
my %tables = (
  H => $hosts,
  D => $disks,
  S => $shelves,
  A => $aggregates,
);
        

my $word = CreateObject Win32::OLE 'Word.Application' or die $!;
$word->{'Visible'} = 1;

my $BaseDir=$Bin;
my $document = $word->Documents->Open("$BaseDir/7mode.docx");
save_doc_as($document, "$BaseDir/output.docx");

my $tables = $word->ActiveDocument->Tables;

TABLE:
for my $table (Win32::OLE::in($tables))
{
   print "Processing table...\n";

   my $rows = $table->Rows->{Count};
   #print STDERR "Count of rows is $rows\n";
   my $cols = $table->Columns->{Count};

   for (my $r=1; $r<=$rows; ++$r) {
      my $text = $table->Cell($r,1)->Range->{Text};
      $text =~ s/[[:cntrl:]]+//g;
      print "$r: [$text]\n";
      if ($text =~ /^<(\w+):(\w+)>/) {
        my $tbl = $1;
        my $tag = "$1:$2";
        
        print "Processing this line with tag $tag...\n";
        my $data_table = $tables{$tbl};
        die "There is no table $tbl" unless ($data_table);
        # Replace the first line, add additional
        my $num_keys = scalar(keys %$data_table);
        print "Number of rows in $tbl table: $num_keys\n";
        $table->Rows->Item($r)->Select;
        $word->Selection->Copy;
        #$table->Cell($r,1)->Range->{Text} = "Line 1";
        for (2..$num_keys) {
          print "Adding row in $tbl table...\n";
          $word->Selection->PasteAppendTable;
          #$table->Rows->Add($table->Rows($r));
        }
        # Process the added rows
        my @keys = sort keys %$data_table;
        print "Processing new rows - keys are " . join ("; ", @keys) . "\n\n";
        for (my $i=0; $i<$num_keys; ++$i) {
          my $key = shift @keys;
          my $data_row = $data_table->{$key};
          for (my $col=1; $col<=$cols; ++$col) {
             my $cell = $table->Cell($r+$i,$col);
             my $txt = $cell ? $cell->Range->{Text} : "";
             print "Cell ($i, $col): $txt\n";
             if ($txt =~ /^<CMP>/ && $col > 2) {
                # Special tag to compare two previous columns
                my $t1 = $table->Cell($r+$i,$col-2)->Range->{Text};
                my $t2 = $table->Cell($r+$i,$col-1)->Range->{Text};
                my $t = ($t1 eq $t2) ? "No Action Required" : "See Below";
                print "Comparing previous cells: $t1 and $t2 - $t\n";
                $table->Cell($r+$i,$col)->Range->{Text} = $t;

             }
                 
             if ($txt =~ /^<(\w+):(\w+)>/) {
               my $tag = "$1:$2";
               print "Searching for $tag in " . Dumper ($data_row) . "\n";
               if ($data_row->{$tag}) {
                  print STDERR "Replacing $txt with $data_row->{$tag}\n";
                  $table->Cell($r+$i,$col)->Range->{Text} = $data_row->{$tag};
               }
             }
          }
        #$table->Cell($r,1)->Range->{Text} = "Line 2";
        #last; # Stop processing the currentrows, start processing the next table
        }
        $r += $num_keys;
        $rows += $num_keys;  
        print "New starting row in the table is $r (of total $rows rows)...\n";
      #next TABLE; 
      }
   }
}
#$table->Cell(3,1)->Range->{Text} = "Test string";

$document->Save();   
#save_doc_as($document, 'c:\temp\winword\ttt.doc');

## uncomment the following two if word should shut down
# close_doc($document);
# $word->Quit;

sub text {
  my $document = shift;
  my $text      = shift;

  $document->ActiveWindow->Selection -> TypeText($text);
}

# aka new line, newline or NL
sub enter {
  my $document = shift;

  $document->ActiveWindow->Selection -> TypeParagraph;
}

# use switch_view to change to header, footer, main document and so on...
# possible constants for view are: wdSeekCurrentPageFooter 
#
#   o  wdSeekCurrentPageHeader 
#   o  wdSeekEndnotes 
#   o  wdSeekEvenPagesFooter 
#   o  wdSeekEvenPagesHeader 
#   o  wdSeekFirstPageFooter 
#   o  wdSeekFirstPageHeader 
#   o  wdSeekFootnotes 
#   o  wdSeekMainDocument 
#   o  wdSeekPrimaryFooter 
#   o  wdSeekPrimaryHeader 
#
sub switch_view {
  my $document = shift;
  my $view     = shift;
  $document -> ActiveWindow -> ActivePane -> View -> {SeekView} = $view;
}

sub insert_picture {
  my $document = shift;
  my $file     = shift;
  my $left     = shift;
  my $top      = shift;
  my $width    = shift;
  my $height   = shift;

  my $picture = 
    $document-> Shapes -> AddPicture (
      $file, 
      msoFalse, # link to file
      msoTrue,  # save with document
      $left, $top, $width, $height, 
      $document->ActiveWindow->Selection->{Range}
    );

  return $picture;
}

sub bold {
  my $document = shift;
  my $bold     = shift;

  $document->ActiveWindow->Selection->{Font}->{Bold} = $bold ? msoTrue : msoFalse;
}

sub style_indents {
  my $document          = shift;
  my $style_arg         = shift;
  my $first_line_indent = shift;
  my $other_line_indent = shift;

  my $style = $document->Styles($style_arg->{name});

  $style->ParagraphFormat->{LeftIndent     } =  $other_line_indent;
  $style->ParagraphFormat->{FirstLineIndent} = -$other_line_indent + $first_line_indent;
}

sub items {
  my $document = shift;
  my $title    = shift;
  my $style    = shift;
  my @array    = @_;

  enter($document);
  set_style($document, $style);

  bold($document, 1);
  text($document, $title);
  bold($document, 0);
  text($document, "\x09");

  foreach my $a (@array) {
    text($document, $a);
    text($document, "\x0b");
  }
}

sub insert_box {
  my $document = shift;
  my $left     = shift;
  my $top      = shift;
  my $width    = shift;
  my $height   = shift;

  my $shape = $document->Shapes->AddTextbox(msoTextOrientationHorizontal, $left, $top, $width, $height);
  $shape -> Select;
  my $selection = $word->Selection;
  $selection -> ShapeRange -> Line -> {DashStyle} = msoLineRoundDot;

  return $shape
}

sub close_doc {
  my $document = shift;
  $document -> Close;
}

sub save_doc_as {
  my $document = shift;
  my $filename = shift;

  $document->SaveAs($filename);
}

sub style_keep_with_next {
  my $document     = shift;
  my $style_arg    = shift;

  my $style = $document->Styles($style_arg->{name});

  $style->{ParagraphFormat}->{KeepWithNext} = msoTrue;
}

sub style_keep_together {
  my $document     = shift;
  my $style_arg    = shift;

  my $style = $document->Styles($style_arg->{name});

  $style->{ParagraphFormat}->{KeepTogether} = msoTrue;
}

sub style_border {
  my $document     = shift;
  my $style_arg    = shift;
  my $border       = shift; 
  my $border_style = shift;
  my $border_width = shift;
  my $border_color = shift;

  my $style = $document->Styles($style_arg->{name});

  $style->Borders($border) -> {LineStyle} = $border_style;
  $style->Borders($border) -> {LineWidth} = $border_width;
  $style->Borders($border) -> {Color    } = $border_color;
}

sub style_tab_at_position {
  my $document     = shift;
  my $style_arg    = shift;
  my $position     = shift;
  my $left_or_right= shift;

  my $style = $document->Styles($style_arg->{name});

  $style->ParagraphFormat->{TabStops}->Add($word->InchesToPoints($position), $left_or_right);
}

sub style_space_before {
  my $document     = shift;
  my $style_arg    = shift;
  my $space        = shift; 

  my $style = $document->Styles($style_arg->{name});

  $style->ParagraphFormat->{SpaceBefore} = $space;
}

sub style_space_after {
  my $document     = shift;
  my $style_arg    = shift;
  my $space        = shift; 

  my $style = $document->Styles($style_arg->{name});

  $style->ParagraphFormat->{SpaceAfter} = $space;
}

sub style_alignment {
  my $document     = shift;
  my $style_arg    = shift;
  my $alignment    = shift;

  my $style = $document->Styles($style_arg->{name});

  $style->ParagraphFormat->{Alignment} = $alignment;
}

sub goto_end_of_document {
  my $document  = shift;

  $document->ActiveWindow->Selection->{Range} -> EndKey(wdStory);

  #my $selection = $word->Selection;

  #$selection -> EndKey (wdStory);
}

sub insert_page_break {
  my $document  = shift;

  #my $selection = $word->Selection;

  #$selection -> InsertBreak(wdPageBreak);
  $document->ActiveWindow->Selection->{Range} -> InsertBreak(wdPageBreak);
}

sub landscape {
  my $document = shift;

  $document->PageSetup->{Orientation} = wdOrientLandscape;
}