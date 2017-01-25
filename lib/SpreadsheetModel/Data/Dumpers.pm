package SpreadsheetModel::Data::Dumpers;

=head Copyright licence and disclaimer

Copyright 2008-2014 Reckon LLP and others.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice,
this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation
and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

=cut

use warnings;
use strict;
use utf8;

sub xlsWriter {
    require Spreadsheet::WriteExcel;
    sub {
        my ( $infile, $workbook ) = @_;
        die unless $infile;
        my $outfile = "$infile cleaned.xls";
        $outfile =~ s/\.xlsx? cleaned.xls$/ cleaned.xls/is;
        if ( -e $outfile ) {
            warn "$infile skipped";
            return;
        }
        my $outputBook = new Spreadsheet::WriteExcel($outfile);
        for my $worksheet ( $workbook->worksheets() ) {
            my $outputSheet = $outputBook->add_worksheet( $worksheet->{Name} );
            my ( $row_min, $row_max ) = $worksheet->row_range();
            my ( $col_min, $col_max ) = $worksheet->col_range();
            for my $row ( $row_min .. $row_max ) {
                for my $col ( $col_min .. $col_max ) {
                    my $cell = $worksheet->get_cell( $row, $col );
                    next unless $cell;
                    next unless my $v = $cell->unformatted;
                    $v =~ /=/
                      ? $outputSheet->write_string( $row, $col, $v )
                      : $outputSheet->write( $row, $col, $v );
                }
            }
        }
    };
}

sub xlsFlattener {
    require Spreadsheet::WriteExcel;
    sub {
        my ( $infile, $workbook ) = @_;
        die unless $infile;
        my $outfile = "$infile flattened.xls";
        $outfile =~ s/\.xlsx? flattened.xls$/ flattened.xls/is;
        if ( -e $outfile ) {
            warn "$infile skipped";
            next;
        }
        my $outputBook  = new Spreadsheet::WriteExcel($outfile);
        my $outputSheet = $outputBook->add_worksheet('Flattened');
        my $outputRow   = -1;
        for my $worksheet ( $workbook->worksheets() ) {
            my ( $row_min, $row_max ) = $worksheet->row_range();
            my ( $col_min, $col_max ) = $worksheet->col_range();
            for my $row ( $row_min .. $row_max ) {
                $outputSheet->write_string( ++$outputRow, $col_min,
                    $worksheet->{Name} );
                for my $col ( $col_min .. $col_max ) {
                    my $cell = $worksheet->get_cell( $row, $col );
                    next unless $cell;
                    next unless my $v = $cell->unformatted;
                    $v =~ /=/
                      ? $outputSheet->write_string( $outputRow, $col + 1, $v )
                      : $outputSheet->write( $outputRow, $col + 1, $v );
                }
            }
        }
    };
}

sub xlsSplitter {
    require Spreadsheet::WriteExcel;
    sub {
        my ( $infile, $workbook ) = @_;
        die unless $infile;
        for my $worksheet ( $workbook->worksheets() ) {
            my $outfile = "$infile $worksheet->{Name}.xls";
            if ( -e $outfile ) {
                warn "$infile skipped";
                next;
            }
            my $outputBook  = new Spreadsheet::WriteExcel($outfile);
            my $outputSheet = $outputBook->add_worksheet( $worksheet->{Name} );
            my ( $row_min, $row_max ) = $worksheet->row_range();
            my ( $col_min, $col_max ) = $worksheet->col_range();
            for my $row ( $row_min .. $row_max ) {
                for my $col ( $col_min .. $col_max ) {
                    my $cell = $worksheet->get_cell( $row, $col );
                    next unless $cell;
                    next unless my $v = $cell->unformatted;
                    $v =~ /=/
                      ? $outputSheet->write_string( $row, $col, $v )
                      : $outputSheet->write( $row, $col, $v );
                }
            }
        }
    };
}

sub tsvDumper {
    my ($output) = @_;
    my $joinChar = $output eq 'csv' ? ',' : "\t";
    sub {
        my ( $infile, $workbook ) = @_;
        my $fh;
        if ( ref $output ) {
            $fh = $output;
        }
        else {
            open $fh, '>', "$infile.$output";
        }
        binmode $fh, ':utf8';
        for my $worksheet ( $workbook->worksheets() ) {
            my ( $row_min, $row_max ) = $worksheet->row_range();
            my ( $col_min, $col_max ) = $worksheet->col_range();
            $col_min = 0;    # Use completely blank columns
            print {$fh} join(
                $joinChar,
                $infile,
                $worksheet->{Name},
                0,
                map {
                    my $aa = int( $_ / 26 );
                    ( $aa ? chr( 64 + $aa ) : '' ) . chr( 65 + ( $_ % 26 ) );
                } $col_min .. $col_max
            ) . "\n";
            for my $row ( $row_min .. $row_max ) {
                print {$fh} join(
                    $joinChar,
                    $infile,
                    $worksheet->{Name},
                    1 + $row,
                    map {
                        my $cell = $worksheet->get_cell( $row, $_ );
                        local $_ = !$cell ? '' : $cell->unformatted;
                        s/\n/\\n/gs;
                        s/\r/\\r/gs;
                        $_;
                    } $col_min .. $col_max
                ) . "\n";
            }
        }
    };
}

sub tallDumper {
    my ($output) = @_;
    my $joinChar = $output eq 'csv' ? ',' : "\t";
    sub {
        my ( $infile, $workbook ) = @_;
        my $fh;
        if ( ref $output ) {
            $fh = $output;
        }
        else {
            open $fh, '>', "$infile.$output";
        }
        binmode $fh, ':utf8';
        print {$fh} join( $joinChar, qw(File Sheet Row Column Cell Contents) )
          . "\n";
        for my $worksheet ( $workbook->worksheets() ) {
            my ( $row_min, $row_max ) = $worksheet->row_range();
            my ( $col_min, $col_max ) = $worksheet->col_range();
            for my $row ( $row_min .. $row_max ) {
                foreach ( $col_min .. $col_max ) {
                    my $aa = int( $_ / 26 );
                    $aa =
                      ( $aa ? chr( 64 + $aa ) : '' ) . chr( 65 + ( $_ % 26 ) );
                    my $cell = $worksheet->get_cell( $row, $_ );
                    next unless $cell;
                    my $v = $cell->unformatted;
                    next unless defined $v && $v ne '';
                    $v =~ s/\n/\\n/gs;
                    $v =~ s/\r/\\r/gs;
                    print {$fh} join( $joinChar,
                        $infile, $worksheet->{Name}, 1 + $row, 1 + $_,
                        $aa . ( 1 + $row ), $v, )
                      . "\n";
                }
            }

        }
    };
}

1;
