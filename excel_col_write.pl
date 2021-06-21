#!/usr/bin/perl
require 5.8.0;

use warnings;
use Spreadsheet::WriteExcel; 
use Excel::Writer::XLSX; 

my @data;

$data[0] = "123,321,123,321";
$data[1] = "999,888,777,666";
@row = split(',',$data[0]);
@col = split(',',$data[1]);
@data_after = ([@row],
				[@col]);
#$line = \@row;
#$line1 = \@col;
$line = \@data_after;
&make_data_excel;
sub make_data_excel{

	my $xl = Excel::Writer::XLSX->new("TEST.xlsx");
	my $xlsheet = $xl->add_worksheet("Sheet10");


#************Add format*****************	   
	 glob $standard_head = $xl->add_format();
	   $standard_head->set_bold();
	   $standard_head->set_size('11');
	   $standard_head->set_font('MS Gothic'); 
	   
	glob $standard = $xl->add_format();
	   $standard->set_size('9');
	   $standard->set_font('MS Gothic'); 
       	   $standard->set_left(1);
	   $standard->set_right(1);
	   	excel_write($xlsheet);
}
sub excel_write{ 
	my $xlsheet = shift @_;
	my $now = $xlsheet->store_formula('=SUM(A2:B2)+A2');
	
	
	$xlsheet->write_row(1,0,$line,$standard);
	#$xlsheet->write_formula( 1,2,'=SUM(A2:B2)' );
	for my $row(2..5){
	$xlsheet->repeat_formula($row-1,2,$now,$standard,
		"2".qw/$row-1/,'A'.$row,
		"3".qw/$row-1/,'B'.$row,
		"2".qw/$row-1/,'A'.$row);
	};
}
