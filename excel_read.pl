#!/usr/bin/perl
use Spreadsheet::Read;
use Data::Dumper;
use Smart::Comments;

my $file = $ARGV[0];
my $spreadsheet = ReadData($file) or die "Cannot read file";
my $sheet_count = $spreadsheet->[0]{sheets} or die "No sheets in $file\n";

for my $sheet_index($sheet_count){
	my $sheet = $spreadsheet->[$sheet_index] or next;
	#print("%s - %2d: [%-s] %3d Cols, %5d Rows/n",$file,$sheet_index,$sheet->{label},$sheet->{maxcol},$sheet->{maxrow});
	for my $row(1..$sheet->{maxrow}){
		#print my $data = $sheet->{cell}[2][$row];
		print join "/t" => map {
									my $data = $sheet->{cell}[$_][$row] ;
									defined $data ? $data : "-";#’l‚ª–³‚¢‚È‚ç"-"‚Å•\Ž¦
								}1 .. $sheet->{maxcol};
		print "/n";
	};
}