#!/usr/bin/perl
use strict;
use warnings;
use feature 'say';
use Excel::Writer::XLSX;
use Spreadsheet::Read qw(ReadData);

# Writting to Spredsheet

my $workbook  = Excel::Writer::XLSX->new( 'simple.xlsx' );
my $worksheet = $workbook->add_worksheet();

my @data_for_row = (1, 2, 3);
my @table = (
    [4, 5],
	[6, 7],
);
my @data_for_column = (10, 11, 12);


$worksheet->write( "A1", "Hi Excel!" );
$worksheet->write( "A2", "second row" );

$worksheet->write( "A3", \@data_for_row );
$worksheet->write( 4, 0, \@table );
$worksheet->write( 0, 4, [ \@data_for_column ] );

$workbook->close;

# Reading from Spredsheet

my $book = ReadData('simple.xlsx');

# The parameters here is sheet, column
my @row = Spreadsheet::Read::row($book->[1], 1);
for my $i (0 .. $#row) {
    say 'A' . ($i+1) . ' ' . ($row[$i] // '');
}

__END__

$ perl test.pl
A1 Hi Excel!
A2
A3
A4
A5 10
