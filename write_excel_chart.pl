#!usr/bin/perl

use warnings;
use Cwd;
use Excel::Writer::XLSX;
use POSIX qw(strftime);
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';

glob $dir = getcwd;
$Win32::OLE::Warn = 3;
$Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit');  # get already active Excel
$Excel->{DisplayAlerts}=0;

glob %hash;
$n = 1;
foreach("A".."Z"){
    $hash{"$n"} = $_;
    $n++;
}
#@file_csv = scan_path('.//');
$row_start = 3;
$col_start = 2;
#foreach(@file_csv){
    &make_excel_chart($row_start,$col_start);

#}
sub make_excel_chart{
    #my $file = shift @_;
    my $row  = shift @_;
    my $col  = shift @_;

    #open (FILE,$file.'.csv');
    #my @data = <FILE>;
    #close FILE;

    my $count = 100;      # pattan count
    #my $count = @data;

    my $workbook  = Excel::Writer::XLSX->new( 'Excel.xlsx' );
    my $worksheet = $workbook->add_worksheet("FF");
    $worksheet = $workbook->add_worksheet("FS");
    $worksheet = $workbook->add_worksheet("TT");
    $worksheet = $workbook->add_worksheet("SF");
    $worksheet = $workbook->add_worksheet("SS");
    $worksheet = $workbook->add_worksheet("Graph");
    foreach(1..$count){
        &make_chart($workbook,$worksheet,$row,$col);
        $row += 18;
    }
    $workbook->close();
    my $Book   = $Excel->Workbooks->Open($dir.'/Excel.xlsx');
    my $Sheet  = $Book->Worksheets(6);

    $row = 3; 
    foreach my $sumple(1..$count){
        foreach(1..15){
            &write_format($Sheet,$_,$row,$sumple);
        }
        $row += 18;
    }
    $Book->Save;
    $Book->Close;
}
sub make_chart{
    my $workbook  = shift @_;
    my $worksheet = shift @_;
    my $row = shift @_;
    my $col = shift @_;
    
    my $chart = $workbook->add_chart( type => 'column',embedded => 1 );
    $chart->set_size( width => 470, height => 350);   #row 18 column 7
    #$chart->set_legend( none => 1);   凡例

    $name = '=Graph!$'.dec2AZ($col).'$'.$row;
    $chart->set_title(   name => "$name",
                         name_font => {
                                 bold => 0,
                                 size => 15},);
    $chart->set_y_axis(  max => 90 , min => 0,major_tick_mark => 'none'); #max sumple count
    $chart->set_x_axis(  major_tick_mark => 'none');

    $chart->add_series(  name => "FF",                 
                         categories => '=Graph!$'.dec2AZ($col).'$'.($row+1).':$'.dec2AZ($col).'$'.($row+14),
                         values => '=Graph!$'.dec2AZ($col+2).'$'.($row+1).':$'.dec2AZ($col+2).'$'.($row+14),);
    $chart->add_series(  name => "FS",                 
                         categories => '=Graph!$'.dec2AZ($col).'$'.($row+1).':$'.dec2AZ($col).'$'.($row+14),
                         values => '=Graph!$'.dec2AZ($col+2+1).'$'.($row+1).':$'.dec2AZ($col+2+1).'$'.($row+14),);
    $chart->add_series(  name => "TT",                 
                         categories => '=Graph!$'.dec2AZ($col).'$'.($row+1).':$'.dec2AZ($col).'$'.($row+14),
                         values => '=Graph!$'.dec2AZ($col+2+2).'$'.($row+1).':$'.dec2AZ($col+2+2).'$'.($row+14),);
    $chart->add_series(  name => "SF",                 
                         categories => '=Graph!$'.dec2AZ($col).'$'.($row+1).':$'.dec2AZ($col).'$'.($row+14),
                         values => '=Graph!$'.dec2AZ($col+2+3).'$'.($row+1).':$'.dec2AZ($col+2+3).'$'.($row+14),);
    $chart->add_series(  name => "SS",                 
                         categories => '=Graph!$'.dec2AZ($col).'$'.($row+1).':$'.dec2AZ($col).'$'.($row+14),
                         values => '=Graph!$'.dec2AZ($col+2+4).'$'.($row+1).':$'.dec2AZ($col+2+4).'$'.($row+14),);

    $chart->add_series( name => 'SPEC',
                        gap => 500,
                        line => {color => 'red'},
                        fill => {color => 'red'},
                        values => '={0,0,200,0,0,0,0,0,0,0,0,0,200,0}',
                        categories => '=Graph!$'.dec2AZ($col).'$'.($row+1).':$'.dec2AZ($col).'$'.($row+14),);

    $worksheet->insert_chart( dec2AZ($col+8).$row, $chart);
}
sub write_format{
    my $Sheet = shift @_;
    my $flag  = shift @_;
    my $row   = shift @_;
    my $sam   = shift @_;
    if($flag == 1){
        $Sheet->Cells($row,2)->{Value} = '=FF!$A$'.$sam;
        $Sheet->Cells($row,4)->{Value} = 'FF';
        $Sheet->Cells($row,5)->{Value} = 'FS';
        $Sheet->Cells($row,6)->{Value} = 'TT';
        $Sheet->Cells($row,7)->{Value} = 'SF';
        $Sheet->Cells($row,8)->{Value} = 'SS';
    }else{
        $Sheet->Cells($row+$flag-1,2)->{Value} = '=FF!$C'.$sam.'+(FF!$D$'.$sam.'-FF!$C$'.$sam.')/10*('.($flag-4).')';
        $Sheet->Cells($row+$flag-1,3)->{Value} = '=FF!$C'.$sam.'+(FF!$D$'.$sam.'-FF!$C$'.$sam.')/10*('.($flag-3).')';
        if($flag == 2){
            $Sheet->Cells($row+$flag-1,4)->{Value} = '=COUNTIFS(FF!$E'.$sam.':$IO'.$sam.',"<"&C'.($flag+$row-1).')';
            $Sheet->Cells($row+$flag-1,5)->{Value} = '=COUNTIFS(FS!$E'.$sam.':$IO'.$sam.',"<"&C'.($flag+$row-1).')';
            $Sheet->Cells($row+$flag-1,6)->{Value} = '=COUNTIFS(TT!$E'.$sam.':$IO'.$sam.',"<"&C'.($flag+$row-1).')';
            $Sheet->Cells($row+$flag-1,7)->{Value} = '=COUNTIFS(SF!$E'.$sam.':$IO'.$sam.',"<"&C'.($flag+$row-1).')';
            $Sheet->Cells($row+$flag-1,8)->{Value} = '=COUNTIFS(SS!$E'.$sam.':$IO'.$sam.',"<"&C'.($flag+$row-1).')';
        }elsif($flag == 15){
            $Sheet->Cells($row+$flag-1,4)->{Value} = '=COUNTIFS(FF!$E'.$sam.':$IO'.$sam.',">="&B'.($flag+$row-1).')';
            $Sheet->Cells($row+$flag-1,5)->{Value} = '=COUNTIFS(FS!$E'.$sam.':$IO'.$sam.',">="&B'.($flag+$row-1).')';
            $Sheet->Cells($row+$flag-1,6)->{Value} = '=COUNTIFS(TT!$E'.$sam.':$IO'.$sam.',">="&B'.($flag+$row-1).')';
            $Sheet->Cells($row+$flag-1,7)->{Value} = '=COUNTIFS(SF!$E'.$sam.':$IO'.$sam.',">="&B'.($flag+$row-1).')';
            $Sheet->Cells($row+$flag-1,8)->{Value} = '=COUNTIFS(SS!$E'.$sam.':$IO'.$sam.',">="&B'.($flag+$row-1).')';
        }else{
            $Sheet->Cells($row+$flag-1,4)->{Value} = '=COUNTIFS(FF!$E'.$sam.':$IO'.$sam.',">="&B'.($flag+$row-1).',FF!$E'.$sam.':$IO'.$sam.',"<"&C'.($flag+$row-1).')';
            $Sheet->Cells($row+$flag-1,5)->{Value} = '=COUNTIFS(FS!$E'.$sam.':$IO'.$sam.',">="&B'.($flag+$row-1).',FS!$E'.$sam.':$IO'.$sam.',"<"&C'.($flag+$row-1).')';
            $Sheet->Cells($row+$flag-1,6)->{Value} = '=COUNTIFS(TT!$E'.$sam.':$IO'.$sam.',">="&B'.($flag+$row-1).',TT!$E'.$sam.':$IO'.$sam.',"<"&C'.($flag+$row-1).')';
            $Sheet->Cells($row+$flag-1,7)->{Value} = '=COUNTIFS(SF!$E'.$sam.':$IO'.$sam.',">="&B'.($flag+$row-1).',SF!$E'.$sam.':$IO'.$sam.',"<"&C'.($flag+$row-1).')';
            $Sheet->Cells($row+$flag-1,8)->{Value} = '=COUNTIFS(SS!$E'.$sam.':$IO'.$sam.',">="&B'.($flag+$row-1).',SS!$E'.$sam.':$IO'.$sam.',"<"&C'.($flag+$row-1).')';
        }
    }
}

sub scan_path{
	my @files  =  glob($_[0].'*');
	my $len    =  0;
	my @file_list;
    my $file;
	foreach(@files){
		if($_ =~ /csv$/i ) {
			$len  =  length($_);
			$file = substr($_,3,$len-3);
            $file =~ s/\.csv//;
            push @file_list,$file;
		}
	}
	return @file_list;
}

sub dec2AZ{
    my $num = shift;
    my $result;
    my $n1;
    my $n2;
    my $n3;
    if ($num > 26){
        $n1 =  ($num-1) % 26 +1;
        $n2 =  ($num-1) / 26;
        $n2 =  int $n2;
        if ($n2 > 26){
            $n3 =  ($n2-1) / 26;
            $n3 =  int $n3;
            $n2 =  ($n2-1) % 26 +1;
            $result = $hash{$n3}.$hash{$n2}.$hash{$n1};
        }
        else{ $result = $hash{$n2}.$hash{$n1};  }
    }
    else{ $result = $hash{$num}; }

    return $result;    
}
