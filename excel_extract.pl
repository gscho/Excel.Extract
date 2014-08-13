#!/usr/bin/perl
use strict;
use warnings;
use Spreadsheet::Read;
use Spreadsheet::WriteExcel;

#The code below is for opening the excel file and for creating a new seperated excel file 
#The names of the files will be given by the user via the terminal at runtime

my $excelFile = "CountryCodes.xls";
my $newExcelFile = "UKCustomers.xls";
my $book = ReadData  ($excelFile);
my $parser = Spreadsheet::ParseExcel->new();
my $newBook = Spreadsheet::WriteExcel->new($newExcelFile);
my $newSheet = $newBook->add_worksheet();
my $format = $newBook->add_format();
$format->set_bold();


 
#Setting up the new excel sheet that will be created. These are the new column titles.
$newSheet->write(0,0,'BSBACCTID',$format);
$newSheet->write(0,1,'COMPANYNAME',$format);
$newSheet->write(0,2,'ADDRESS',$format);
$newSheet->write(0,3,'LOCALE',$format);
$newSheet->write(0,4,'EMAIL',$format);


my $row = 2;
my $col = 4;
my $newrow = 1;
my $newcol = 0;
my $countrycode = "gb";
#Loops through all the rows and columns in the excel file
while ($row <= 16698){

my $cell4 = $book->[1]{cell}[$col][$row];
my $cell1 = $book->[1]{cell}[$col-3][$row];
my $cell2 = $book->[1]{cell}[$col-2][$row];
my $cell3 = $book->[1]{cell}[$col-1][$row];
my $cell5 = $book->[1]{cell}[$col+1][$row];

#The code below starts searching for the country code given by the user
if (index($cell4, $countrycode) != -1){
	my @values4 = split(',', $cell4);
	if(index($values4[3], "gb") != -1){
		
		$newSheet->write($newrow,$newcol,"$cell1",$format);
		$newSheet->write($newrow,$newcol+1,"$cell2",$format);
		$newSheet->write($newrow,$newcol+2,"$cell3",$format);
		$newSheet->write($newrow,$newcol+3,"$cell4",$format);
		$newSheet->write($newrow,$newcol+4,"$cell5",$format);
		#print "$cell1 $cell2 $cell3 $cell4 $cell5\n";
		$newrow++;
		
		
	}
}
	$row++;
}