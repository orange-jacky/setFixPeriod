#! c:/perl/bin/perl.exe 

use strict;
use Win32::OLE;
use Win32::OLE::Variant;
use Cwd;

my $src = "CallCenter系统监控维护期申请表20130320.docx";
my $dest = "维护期导入模板_1029.xlsx";
my $wc1 = 4;
my $wc2 = 9;
my $wc3 = 10;
my $ec1 = 1;
my $ec2 = 3;
my $ec3 = 4;

my $path = getcwd();

#open word
my $word = Win32::OLE->GetActiveObject('Word.Application') || Win32::OLE->new('Word.Application', 'Quit');
my $document = $word->Documents->Open("${path}/${src}") || die("Unable to open document ", Win32::OLE->LastError());

#open excel
my $excel = Win32::OLE->GetActiveObject('Excel.Application') || Win32::OLE->new('Excel.Application', 'Quit');
my $book = $excel->Workbooks->Open("${path}/${dest}") || die("Unable to open book ", Win32::OLE->LastError());
my $sheet = $book->Worksheets(1);

#get data from word
my $table = $document->Tables(1);
my $totalLine =  $table->{'Rows'}->{'count'}; #get total line of table

my $tmp;
my $var;
my $a;
my $b;

foreach my $row (2 .. $totalLine)
{
	#format ip
	$tmp = $table->Cell($row, $wc1)->Range->{Text};
	$tmp =~ /^(.+)\s(.+)$/;
	$sheet->Cells($row, $ec1)->{Value} = $1;
	
	#formate start
	$tmp = $table->Cell($row, $wc2)->Range->{Text};	
	$tmp =~ /^(.+)\s(.+)\s(.+)$/;
	$a = $1;
	$b = $2;
	
  $var= $a." ".$b;
	$sheet->Cells($row, $ec2)->{Value} = 	$var;
	
	#formate end
	$tmp = $table->Cell($row, $wc3)->Range->{Text};	
	$tmp =~ /^(.+)\s(.+)\s(.+)$/;
  $var= $1." ".$2;
	$sheet->Cells($row, $ec3)->{Value} = $var;
}


#save excel
$excel->Save;
$document->Save;

#close all
$document->Close;
$book->Close;

