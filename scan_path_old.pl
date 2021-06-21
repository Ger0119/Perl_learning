#!/usr/bin/perl

$path = './/*';
@File = &scan_path($path);

sub scan_path{
	my @files  =  glob($_[0]);
	my $file   =  '';
	my $len    =  0;
	my @file_list;
	foreach(@files){
		if($_ =~ /csv$/){ #Fail name conditions
			$len   =  length($_);
			$file  =  substr($_,3,$len-3);
			push(@file_list,$file);
		}
	}
	return @file_list;
}

