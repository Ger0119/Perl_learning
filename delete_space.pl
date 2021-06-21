#!/usr/bin/perl
my @String = '';
$String[0] = " dfd afe    dv adf df\n";
$String[1] = " dfa 2df dafd ";

delete_space(@String);
sub delete_space(){
	my @string = @_;
	my $len= @string;
	for (my $i = 0;$i<$len;$i++){

		$string[$i] =~s/\s+/\,/g;

	}
	print @string;
}