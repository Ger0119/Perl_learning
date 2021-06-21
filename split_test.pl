#!usr/perl/bin

$string = '123,45,5,3,1';
$string1 = '654,321,654';

@str = split(',',$string);

print join " ",@str;

@str = undef;
print "\n";

#@str = split(',',$string1);
print join " ",@str;

