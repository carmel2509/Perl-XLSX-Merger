#!/usr/bin/perl
use strict;
use warnings;
use Spreadsheet::Read; # Counts from 1...
use Excel::Writer::XLSX; # Counts from 0...
use List::Util qw[min max];

sub MergedXLSX{
	my $filename1 = $_[0];
	my $filename2 = $_[1];

	my $Wb1  = ReadData($filename1); # Reads workbook
	my $No_row1 = $Wb1->[1]{maxrow}; # Number of active rows in sheet 1 (Returned as non-int)
	my $No_col1 = $Wb1->[1]{maxcol}; # Number of active columns in sheet 1


	my $Wb2 = ReadData($filename2); # Reads workbook
	my $No_row2 = $Wb2->[1]{maxrow}; # Number of active rows in sheet 1 (Returned as non-int)
	my $No_col2 = $Wb2->[1]{maxcol}; # Number of active columns in sheet 1
    
	my $workbook  =Excel::Writer::XLSX->new("Merged.xlsx");
	my $worksheet = $workbook->add_worksheet();      
          
	my $Min_row = min($No_row1,$No_row2);
	my $Min_col = min($No_col1,$No_col2); # Number of matching headers

	my $Max_row = max($No_row1,$No_row2);

	my $No_row = $No_row1 + $No_row2; # Used during iteration
	my $No_col = $No_col1; # Same number of columns


	my $Mid_row;
	if ($No_row1 < $No_row2){
	$Mid_row = $Min_row;
	}else{
		$Mid_row = $Max_row;
	}

	for ( my $Row = 1; $Row <= int($No_row); $Row = $Row + 1){  
		for ( my $Col = 1; $Col <= int($No_col); $Col = $Col + 1){ 
			if (defined $Wb1->[1]{cell}[$Col][$Row]){ # If Cell has value
				$worksheet->write( $Row-1,$Col-1,  $Wb1->[1]{cell}[$Col][$Row]);   			
			}elsif ($Row == $Mid_row + 1){ # Excludes header of wb2
				next;
			}else{
				$worksheet->write( $Row-2,$Col-1,  $Wb2->[1]{cell}[$Col][$Row-$Mid_row]); 
			}
		} 
	}

	$workbook->close();   
}