#!usr/bin/perl

use warnings;
use Win32::OLE::Const 'Microsoft Excel';
use Cwd;

$dir = getcwd;
$Win32::OLE::Warn = 3;
$Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit') or die "NO EXCEL APPLICAPTION!!";  # get already active Excel
$Excel->{DisplayAlerts}=0;  

sub write_excel{
    my $file = shift;
    my $book = $Excel->Workbooks->Open($dir."/".$file);
    my $sheet = $book->Worksheets(1);
    my $data = $sheet->Range('G1:G12000')->{Value}; 
    $book->Close;

    my $Book = $Excel->Workbooks->Open($dir."/".$new_exl);
 
    my $Sheet = $Book->Worksheets(1);
    $Sheet->Range('G1:G12000')->{Value} = $data;

    $Book->Save;
    $Book->Close;
}

