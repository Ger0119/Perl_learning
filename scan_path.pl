#!usr/bin/perl

use warnings;
use Cwd;

#$Dir  = getcwd;
@File = scan_path(getcwd.'/');

sub scan_path{
    my $path  =  shift @_;
	my @file  =  glob($path.'*');
	my @files;
	foreach(@file){
		if($_ =~ /csv$/i ) {
            $_ =~ s/$path//;
            push @files,$_;
		}
	}
	return @files;
}

