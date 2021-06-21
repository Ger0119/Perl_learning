#!/usr/bin/perl
use strict;
use warnings;

use Spreadsheet::Read;

my $workbook = ReadData ("$ARGV[0]");
print $workbook->[1]{cell}[2][2] . "\n";