#!/usr/bin/perl
my ($path) = @ARGV;

sub scan_file{
    my @files = glob(@_[0]);
    foreach (@files){
        if(-d $_){
            my $path = "$_/*";
            scan_file($path);
        }elsif(-f $_){
            print "file $_\n";
        }
    }
}
scan_file($path);