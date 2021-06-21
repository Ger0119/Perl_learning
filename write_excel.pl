#!usr/bin/perl

use warnings;
use Excel::Writer::XLSX;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use Cwd;

$dir = getcwd;
$Win32::OLE::Warn = 3;
$Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit');  # get already active Excel
$Excel->{DisplayAlerts}=0;  
#$file_name = $ARGV[0];
$Path = 'C:/Users/a5122641/Desktop/make_tool/wt_map/ad/*';
@FILE_list  = &scan_path($dir."/*");

foreach(@FILE_list){
	&write_excel($_);
	#print $_;
}

sub write_excel{
	my $Book = $Excel->Workbooks->Open($_); 
	my $Sheet = $Book->Worksheets(1);
	my $num = $Sheet->Cells(1,"A")->{Value};
	print $num;
	&write_cell($Sheet,2,3,"some");
	#$Sheet->Cells(3,"B")->{Value}="good";
	$Book->Save;
	$Book->Close;
	  
}
sub write_cell{
	my $Sheet  = shift;
	my $row    = shift;
	my $colnum = shift;
	my $value  = shift;

	$Sheet->Cells($row,$colnum)->{Value} = $value;

	my @csv = scan_path_csv('.//*');
	open (FILE,$csv[0]);
	my @data = <FILE>;
	close FILE;

	my @array = split(',',$data[0]);

	foreach(@array){
		$Sheet->Cells($row,$colnum)->{Value} = $_;

	}
}

sub scan_path{
	my @files  =  glob($_[0]);
	my $file   =  '';
	my $len    =  0;
	my @file_list;
	foreach(@files){
		if($_ =~ /xlsx$/i ) {
			$len   =  length($_);
			#$file  =  substr($_,3,$len-3);
			push(@file_list,$_);
		}
	}
	return @file_list;
}
sub scan_path_csv{
	my @files  =  glob($_[0]);
	my $file   =  '';
	my $len    =  0;
	my @file_list;
	foreach(@files){
		if($_ =~ /csv$/i ) {
			$len   =  length($_);
			$file  =  substr($_,3,$len-3);
			push(@file_list,$file);
		}
	}
	return @file_list;
}
