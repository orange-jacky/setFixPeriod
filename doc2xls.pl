#! c:/perl/bin/perl.exe -w

use strict;
use Win32::OLE;
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



#open word document
my $word = Win32::OLE->GetActiveObject('Word.Application')
    || Win32::OLE->new('Word.Application', 'Quit'); 
my $document = $word->Documents->Open("${path}/${src}")
    || die("Unable to open document ", Win32::OLE->LastError());
    
#open excel document    
my $excel = Win32::OLE->GetActiveObject('Excel.Application')
    || Win32::OLE->new('Excel.Application', 'Quit'); 
my $book = $excel->Workbooks->Open("${path}/${dest}") 
		|| die("Unable to open document ", Win32::OLE->LastError());
my $sheet = $book->Worksheets(1);

# set range format as strings 
$sheet->Cells->{NumberFormatLocal} = "@";
 

#get data from word  then  write  into excel
my $table = $document->Tables(1);#get table object from word
my $lineCount = $table->{'Rows'}->{'count'};#get line's count of table object

my $tmp;
my $var;

foreach my $row (2 .. $lineCount)
{
	#format ip
	$tmp = $table->Cell($row,$wc1)->Range->{Text}; #get data from table in word
	$tmp =~ /^(.+)\s(.+)$/;
	$sheet->Cells($row, $ec1)->{Value} = $1; #write data into excel
	
	#format start time
	$tmp = $table->Cell($row,$wc2)->Range->{Text};
	$tmp =~ /^(.+)\s(.+)\s(.+)$/;
	$var = $1." ".$2;
	$sheet->Cells($row, $ec2)->{Value} = $var;
	
	#format end time
	$tmp = $table->Cell($row,$wc3)->Range->{Text};
	$tmp =~ /^(.+)\s(.+)\s(.+)$/;
	$var = $1." ".$2;
	$sheet->Cells($row, $ec3)->{Value} = $var;	

} 



# Resize Columns  method 1
$sheet->Cells->{Columns}->AutoFit();

# Resize Columns  method 2
#my @columnheaders = qw(A:Z);
#foreach my $range(@columnheaders){
#    $sheet->Columns($range)->AutoFit();
#}


#close all
$book->Save;
$book->Close;
$document->Close;




