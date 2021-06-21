#!usr/bin/perl

use warnings;

#$file_name = $ARGV[0];
$Path = './/*';
@FILE_list  = &scan_path($Path);

foreach(@FILE_list){
	&make_shmoo($_);
}

sub make_shmoo{
	my $file = shift @_;
	my $Rate;
	my $flag = 0;
	my @data_cell = "";
	my $result1;
	my $result2;
	my $line_data = "     ";
	my @data_after = "";
	my @data_end;
	my $ExRate = "0.000nS";
	open (FILE,$file);
	my @data = <FILE>;
	close FILE;
	
    #	$data_end[0] = "              +---------+---------+---------+------\n";
    #$data_end[1] = "                                  ^ 720.0mV\n";
    #$data_end[2] = "             1.100V                           400.0mV\n\n";
    #
	$data_end[0] = "              +---------+---------+--\n";
	$data_end[1] = "                              ^ 740.0mV\n";
	$data_end[2] = "             1.000V               600.0mV\n\n";
    my $Wno = 0;
    my $X = 0;
    my $Y = 0;

	foreach my $line(@data){
        if($line =~ /Wafer No\s+\(DEC\) DUT \d: (\d+)/){
           $Wno = $1;next;
        }
        if($line =~ /X Address\s+\(DEC\) DUT \d: (\d+)/){
            $X = $1;next;
        }
        if($line =~ /Y Address\s+\(DEC\) DUT \d: (\d+)/){
            $Y = $1;next;
        }
		if(($line =~ /^[VR]/ or $line =~ /^\s+/)and $flag ==1){
			$result = get_shm_result(@data_cell);
			$line_data .= $result;
			@data_cell = "";
	    }
        if($line =~ /^Pattern/){push @data_end,"\n$line\n";next;}
        if($line =~ /DUT1:/){push @data_after,"\nDUT1\n";next;}
		if($line =~ /^Rate :/){
			$Rate = get_Rate($line);
			if($flag == 1 ){
				push @data_after,$line_data."|\n";
				$line_data = "     ";
			}elsif($flag == 0){
				push @data_after,"ALL : 10G\nDUT1\n";
			}
			if($Rate eq "6.400nS"){
				$line_data .= "$Rate >|";
			}else{
				$line_data .= "$Rate  |";
			}
			if($ExRate lt $Rate and $flag ==1){
				push @data_after,@data_end;
				push @data_after,"ALL : 25G\nDUT1\n";
			}
			$ExRate = $Rate;
			next;
		}
		if($line =~ /^Vmin/ and $flag == -1){
			$flag = 1;
		}
		if($flag == 1){
			push @data_cell,$line;
		}
		if($line =~ /^\s+/ and $flag == 1){
			push @data_after,$line_data."|\n";
			push @data_after,@data_end;	
			$line_data = "\t";
			$flag = 0;
		}
		#if($line =~ /\*\*\*\*\* ID:/ and $flag == 1){
		#	push @data_after,@data_end;	
		#	$flag = 0;
		#}
		if($line =~ /^Test ID/){
			push @data_after,"\n";
			$flag = -1;
			next;
		}
		#if($line =~ /^\s+\|\s/ and $flag == 0)  { next; }
		#if($line =~ /^\s+\+\-\-/ and $flag == 0){ next; }
		#if($line =~ /^\s+Axis / and $flag == 0) { next; }
		if($flag == 0){	 push @data_after,$line;	}
	}
	#print @data_after;
	$file =~ s/shm_//;
	open  Output,">shmoo_$file";
    print Output "W".$Wno."X".$X."Y".$Y."\n";
	print Output @data_after;
	close Output;

}

sub get_shm_result{
	my @data = @_;
	my @data_check;
	my $num;
	my $flag;
	my $result;
	my @Ignore = &get_Ignore;
	if(scalar(@data) == 1){ return "","";} 
	foreach my $line(@data){
		if($line =~ /^(\d+)/){
			if(grep /^$1$/,@Ignore){
				next;
			}else{
				push @data_check,$line;
			}
		}
	}
	$result = check_result(@data_check);
	return $result;
}
sub check_result{
	my @data = @_;
	my $result;
	
	$result = '*';	
	foreach(@data){
		if($_ =~ /FAIL/){
			$result = 'F';
			#return $result;
			last;
		}
	}
	return $result;
}
sub get_Rate{
	my $line = shift @_;
	my $Rate;

	if($line =~ /^Rate :([.0-9]+)nS/){
		if(length($1) == 1){
			$Rate = $1.".000nS";
		}elsif(length($1) == 3){
			$Rate = $1."00nS";
		}elsif(length($1) == 4){
			$Rate = $1."0nS";
		}else{
			$Rate = $1."nS";
		}
	}
	return $Rate;
}

sub get_Ignore{
	open (FILE,"Ignore_test_No");
	my $line = <FILE>;
	close FILE;

	my @data = split(',',$line);
	return @data;
}

sub scan_path{
	my @files  =  glob($_[0]);
	my $file   =  '';
	my $len    =  0;
	my @file_list;
	foreach(@files){
		if(($_ =~ /txt$/i or $_ =~ /log$/i) and $_ !~ /^.\/\/shmoo_/){
			$len   =  length($_);
			$file  =  substr($_,3,$len-3);
			push(@file_list,$file);
		}
	}
	return @file_list;
}
