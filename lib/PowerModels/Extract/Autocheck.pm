package PowerModels::Extract::Autocheck;

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

use constant {
    AC_CSV_MASTER_AND_LOCK => 0,
    AC_CSV_NEW_CHECKSUMS   => 1,
    AC_REGEX_VALID_SHEETS  => 2,
};

sub new {
    my ( $class, $homes, $validSheetsRegex ) = @_;
    my ($tFolder) =
      grep { -e $_; } ( map { File::Spec->catdir( $_, 't' ); } @$homes ),
      @$homes;
    bless [
        File::Spec->catfile( $tFolder, 'Checksums.csv' ),
        File::Spec->catfile( $tFolder, 'Checksums-new.csv' ),
        $validSheetsRegex,
    ], $class;
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

    local $/ = "\n";
    my @records;
    if ( open my $fh, '<', $autocheck->[AC_CSV_MASTER_AND_LOCK] ) {
        flock $fh, LOCK_EX or die 'flock failed';
        @records = <$fh>;
    }
    else {
        warn "Could not open $autocheck->[AC_CSV_MASTER_AND_LOCK]";
    }
    if ( open my $fh, '<', $autocheck->[AC_CSV_NEW_CHECKSUMS] ) {
        push @records, <$fh>;
    }

    foreach (@records) {
        chomp;
        my @a = split /[\t,]/;
        next unless @a > 4;
        if (   $a[0] eq $file
            && $a[1] eq $year
            && $a[2] eq $company
            && $a[3] == $tableNumber )
        {
            return if $a[4] eq $checksum;
            die "** Checksum error **\nExpected $a[4], got $checksum for"
              . " $book table $tableNumber\n\n";
        }
    }

    my $fh;
    if ( !open $fh, '+>>', $autocheck->[AC_CSV_NEW_CHECKSUMS] and open $fh,
        '<', $autocheck->[AC_CSV_NEW_CHECKSUMS] )
    {
        my @new = <$fh>;
        push @records, @new;
        close $fh;
        unlink $autocheck->[AC_CSV_NEW_CHECKSUMS];
        open $fh, '>>', $autocheck->[AC_CSV_NEW_CHECKSUMS];
        print $fh @new;
    }

    print {$fh} join( ',',
        $file, $year, $company, $tableNumber, $checksum,
        $revision ? $revision : () )
      . "\n";

}

sub writerAndParserOptions {
    my ($autocheck) = @_;
    my @tableNumber;
    my @checksumLocation;
    my $book;
    my $validSheetsRegex = $autocheck->[AC_REGEX_VALID_SHEETS] || qr/./;
    study $validSheetsRegex;

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
                return 1 unless /$validSheetsRegex/;
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
