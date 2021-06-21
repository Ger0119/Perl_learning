#!/usr/bin/perl

use warnings;

$path = './/*';

@file_list = &scan_path($path);

foreach $file(@file_list){
	yoko2tate($file);
}

sub scan_path(){
	my @files = glob($_[0]);
	my $file = '';
	my $len = 0;
	my @file_list;
	
	foreach (@files){

				if($_ =~ /csv/){
					$len = length($_);
					$file = substr($_,3,$len-3);
					push(@file_list,$file);
				}
			}
		
	return @file_list;
}

sub yoko2tate{
	my $file = shift @_;
	my ($tate,$yoko);
	my @data_after;
	my @line;
	my $cell = '';
	
	open(FILE,$file);
	my @data = <FILE>;
	close FILE;
	
	$tate = @data;
	my @data0 = split(',',$data[0]);
	$yoko = @data0;
	
	for(my $i=0;$i<$yoko;$i++){
		for(my $j=0;$j<$tate;$j++){
			@line = split(',',$data[$j]);
			if($j == 0){
				$cell = $line[$i];
			}
			else{
				$cell = $cell.",".$line[$i];
			}
		}
		$cell .= "\n";
		push @data_after,$cell;
		$cell = undef;
	}
	open Output,">new_$file"or die $!;
	print Output @data_after;
	close Output;


}
