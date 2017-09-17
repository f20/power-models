package PowerModels::Data::Autocheck;

=head Copyright licence and disclaimer

Copyright 2016-2017 Franck Latrémolière, Reckon LLP and others.

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
use Encode;
use File::Spec;
use Fcntl qw(:flock :seek);

use constant { AC_CSV => 0, };

sub new {
    my ( $class, $homes ) = @_;
    my ($tFolder) =
      grep { -e $_; } ( map { File::Spec->catdir( $_, 't' ); } @$homes ),
      @$homes;
    bless [ File::Spec->catfile( $tFolder, 'Checksums.csv' ), ], $class;
}

sub processChecksum {

    my ( $autocheck, $book, $tableNumber, $checksumType, $checksum ) = @_;

    ( undef, undef, my $file ) = File::Spec->splitpath($book);
    $file =~ s/\.xlsx?$//si;
    my $revision = '';
    $file =~ s/\+$//s;
    $revision = $1 if $file =~ s/[+-](r[0-9]+)$//si;
    my $company = '';
    my $year    = '';
    ( $company, $year ) = ( $1, $2 )
      if $file =~ s/^(.+)-(20[0-9]{2}-[0-9]{2})-//s;

    my @records;
    my $fh;
         open $fh, '+<', $autocheck->[AC_CSV]
      or open $fh, '<',  $autocheck->[AC_CSV]
      or open $fh, '+>', $autocheck->[AC_CSV]
      or warn "Could not open $autocheck->[AC_CSV]";
    flock $fh, LOCK_EX or die 'flock failed';
    local $/ = "\n";
    @records = <$fh>;
    chomp foreach @records;

    foreach (@records) {
        my @a = split /[\t,]/;
        next unless @a > 4;
        if (   $a[0] eq $file
            && $a[1] eq $year
            && $a[2] eq $company
            && $a[3] == $tableNumber )
        {
            return if $a[4] eq $checksum;
            die "\n\n*** Expected $a[4], got $checksum"
              . " for $book table $tableNumber\n\n";
        }
    }
    push @records, join ',', $file, $year, $company, $tableNumber, $checksum,
      $revision ? $revision : ();
    seek $fh, 0, SEEK_SET;
    print {$fh} map { "$_\n"; } sort @records;

}

sub makeWriterAndParserOptions_slow {
    my ($autocheck) = @_;
    sub {
        my ( $book, $workbook ) = @_;
        for my $worksheet ( $workbook->worksheets() ) {
            my ( $row_min, $row_max ) = $worksheet->row_range();
            my ( $col_min, $col_max ) = $worksheet->col_range();
            my $tableNumber = 0;
            for my $row ( $row_min .. $row_max ) {
                my $rowName;
                for my $col ( $col_min .. $col_max ) {
                    my $cell = $worksheet->get_cell( $row, $col );
                    my $v;
                    $v = $cell->unformatted if $cell;
                    next unless defined $v;
                    eval { $v = Encode::decode( 'UTF-16BE', $v ); }
                      if $v =~ m/\x{0}/;
                    if ( !$col && $v =~ /^([0-9]+)\. /s ) {
                        $tableNumber = $1;
                    }
                    elsif ( $v =~ /^Table checksum ([0-9]{1,2})$/si ) {
                        my $checksumType = $1;
                        my $checksumCell =
                          $worksheet->get_cell( $row + 1, $col );
                        $autocheck->processChecksum( $book, $tableNumber,
                            $checksumType, $checksumCell->unformatted )
                          if $checksumCell;
                    }
                }
            }
        }
    };
}

sub makeWriterAndParserOptions {
    my ($autocheck) = @_;
    my @tableNumber;
    my @checksumLocation;
    my $book;
    (
        sub { },
        Setup => sub { $book = $_[0]; },
        NotSetCell  => 1,
        CellHandler => sub {
            my ( $wbook, $sheetIdx, $row, $col, $cell ) = @_;
            my $v;
            $v = $cell->unformatted if $cell;
            return unless defined $v;
            eval { $v = Encode::decode( 'UTF-16BE', $v ); }
              if $v =~ m/\x{0}/;
            if ( !$col && $v =~ /^([0-9]+)\. /s ) {
                local $_ = $1;
                return 1 unless /^(?:15|16|37|45)/;
                $tableNumber[$sheetIdx] = $_;
            }
            elsif ($checksumLocation[$sheetIdx]
                && $checksumLocation[$sheetIdx][0] == $row
                && $checksumLocation[$sheetIdx][1] == $col )
            {
                $autocheck->processChecksum( $book, $tableNumber[$sheetIdx],
                    $checksumLocation[$sheetIdx][2], $v );
            }
            elsif ( $v =~ /^Table checksum ([0-9]{1,2})$/si ) {
                $checksumLocation[$sheetIdx] = [ $row + 1, $col, $1 ];
            }
            0;
        }
    );
}

1;
