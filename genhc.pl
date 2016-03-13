#!/usr/bin/perl -w
use strict;
use Data::Dumper;

#use File::Glob;
sub add_file($);
my @errors;

my %xl_files;

my %required_files =
(
   hc  =>  "^([^~].+)_HC_Excel_",
   srr =>  "^([^~].+)_SRR_Report_"
);

my @files = <*.xlsx>;

print "Files are " . join (", ", @files) . "\n";

for my $file (@files) {
  add_file ($file);
}


# Verify that we have complete set of files for each customer   
foreach my $customer (keys %xl_files) {
  for my $f (keys %required_files) {
    if (!$xl_files{$customer}->{$f}) {
      push @errors, "No $f file for customer $customer";
    }
  } 
}
    
print Dumper (\%xl_files, \@errors);    



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
         }
         $t->{$f} = $file;
      }
   }
}        
