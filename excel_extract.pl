#!/usr/bin/perl
use strict;
use warnings;

use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::SaveParser;

use Path::Class;
use autodie;# die if problem reading or writing a file

my @testResult;
my @data;
my $linecount = 0;
my $file = file("file.txt");

# Read in the entire contents of a file
my $content = $file->slurp();

# openr() returns an IO::File object to read from
my $file_handle = $file->openr();

# Read in line at a time
while( my $line = $file_handle->getline() ) {
        if (index($line, '***') != -1){	
                @testResult = split(' ',  $line);
                $data[$linecount] = "$testResult[1]";
                $linecount++;
                print "$linecount\n";
        }
       
       
}
        foreach (@data){
       print "$_\n";
}

#################Writing to excel section###################

my $row = 0;
my $col = 0;


my $parser   = new Spreadsheet::ParseExcel::SaveParser;
my $modify = $parser->Parse('test.xls');

    

    # Get the format from the cell
    #my $format   = $template->{Worksheet}[$sheet]
                            #->{Cells}[$row][$col]
                            #->{FormatNo};

    # Write data to some cells
   # $template->AddCell(0, $row,   $col,   1,     $format);
    #$template->AddCell(0, $row+1, $col, "Hello", $format);

    # Add a new worksheet
    #$template->AddWorksheet('New Data');

    # The SaveParser SaveAs() method returns a reference to a
    # Spreadsheet::WriteExcel object. If you wish you can then
    # use this to access any of the methods that aren't
    # available from the SaveParser object. If you don't need
    # to do this just use SaveAs().
    #
    my $workbook;

    {
        # SaveAs generates a lot of harmless warnings about unset
        # Worksheet properties. You can ignore them if you wish.
        local $^W = 0;

        # Rewrite the file or save as a new file
        $workbook = $modify->SaveAs('test.xls');
    }

    # Use Spreadsheet::WriteExcel methods
    my $worksheet  = $workbook->sheets(0);
    my $parser2 = Spreadsheet::ParseExcel->new();
    my $isEmpty = $parser2->parse('test.xls');
    if($isEmpty->get_cell(1,1) eq "")
    {
          
        print "Hello Workld\n";
    }
    $worksheet->write($row+2, $col, "World8");

    $workbook->close();