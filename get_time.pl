#!/usr/bin/perl
use POSIX qw(strftime);

	my $time = strftime "%Y-%m-%d %H:%M:%S", localtime;
	print $time;
