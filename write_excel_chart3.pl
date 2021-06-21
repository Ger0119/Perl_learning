#!usr/bin/perl

use warnings;
use Excel::Writer::XLSX;

glob %hash;
$n = 1;
foreach("A".."Z"){
    $hash{"$n"} = $_;
    $n++;
}
my $workbook  = Excel::Writer::XLSX->new( "Excel.xlsx" );
my $worksheet = $workbook->add_worksheet("result");
$worksheet = $workbook->add_worksheet("graph");
&make_sheet($worksheet);
$workbook->close();
sub make_sheet{
    my $sheet = shift @_;
    my $tar_r = 2;
    my $tar_c = 2;
    my $pat_r = 3;
    my $pat_c = 2;
    foreach my $sam(1..16){
        $tar_c = 2;
        foreach my $pat(1..25){
            $tar_c = make_chart($sheet,$sam*3-1,$pat+4,$tar_r,$tar_c);
        }
        $tar_r += 6;
    
    }
}

sub make_chart{
    my $sheet  = shift @_;
    my $ti     = shift @_;
    my $pat    = shift @_;
    my $t_row  = shift @_;
    my $t_col  = shift @_;
    my $H_limit = '=result!$'.dec2AZ($ti).'$3:$'.dec2AZ($ti+2).'$3';
    my $target = dec2AZ($t_row).$t_col;
    my $title  = '=result!$'.dec2AZ($ti).'$2';
    my $name   = '=result!$A$'.$pat;
    my $chart = $workbook->add_chart( type => 'line',subtype => 'stacked',embedded => 1 );
    $chart->set_size (   width => 380, height => 270); 
    $chart->set_title(
        name => $title,
        name_font => {size => '12'},
    );
    $chart->set_y_axis(
        name => '=result!$'.dec2AZ($ti+1).'$1',
        # major_unit => 0.05,
        min => 0,
    );
    $chart->add_series(
        name        => '=result!$'.dec2AZ($ti+1).'$1',
        categories  => '=result!$B$4:$D$4',
        values      => '=result!$'.dec2AZ($ti).'$'.$pat.':$'.dec2AZ($ti+2).'$'.$pat,
        data_labels => {value =>1 , position => 'top'},
    );
    $chart->add_series(
        name => 'SPEC',
        categories => '=result!$B$4:$D$4',
        values => $H_limit,
        line => { color => 'red'},
    );

    $sheet->insert_chart( $target , $chart);
    $t_col += 14;
    return $t_col;
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
        }else{ $result = $hash{$n2}.$hash{$n1};  }
    }
    else{ $result = $hash{$num} }

    return $result;    
}
