package SpreadsheetModel::Data::ParseCsv;

# Copyright 2011-2017 Franck Latrémolière, Reckon LLP and others.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice,
# this list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
# this list of conditions and the following disclaimer in the documentation
# and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
# EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
# THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

use warnings;
use strict;
use utf8;

sub parseCsvInputData {

    my ( $dataHandle, $blob, $fileName ) = @_;

    local $/ = "\n";
    my $csvParser;
    eval {
        require Text::CSV;
        $csvParser = new Text::CSV( { binary => 1 } );
    };
    eval {
        require Text::CSV_PP;
        $csvParser = new Text::CSV_PP( { binary => 1 } );
    };
    unless ($csvParser) {
        warn "No CSV parsing module; $fileName ignored";
        return;
    }
    my $headers = $csvParser->getline($blob) or return;

    if ( grep { /t([0-9]+)c([0-9]+)/; } @$headers ) {
        my ( @table, @column );
        for ( my $i = 1 ; $i < @$headers ; ++$i ) {
            if ( $headers->[$i] =~ /t([0-9]+)c([0-9]+)/ ) {
                $table[$i]  = $1;
                $column[$i] = $2;
            }
            else {
                $table[$i]  = $headers->[$i];
                $column[$i] = 0;
            }
        }
        while ( !eof($blob) ) {
            if ( my $row = $csvParser->getline($blob) ) {
                my $book = $row->[0];
                my $line = 'Value';
                $line = $row->[1] if $row->[1] && $row->[1] =~ /^[0-9]+$/s;
                $dataHandle->{$book}{ $table[$_] }[ $column[$_] ]{$line} =
                  $row->[$_]
                  foreach grep { $table[$_] } 1 .. $#table;
            }
        }
        return;
    }

    my @rank = map {
            /area|dno/i           ? 0.0
          : /period|year/i        ? 1.0
          : /option/i             ? 2.0
          : /tab/i                ? 3.0
          : /col.*(?:no|number)/i ? 4.0
          : /col/i                ? 4.1
          : /normalis.*row/i      ? 5.0
          : /row.*label/i         ? 5.1
          : /row.*name/i          ? 5.2
          : /row/i                ? 5.3
          : /tariff/i             ? 5.4
          : /value/i              ? 6.0
          : /^v/i                 ? 6.1
          :                         7.0;
    } @$headers;
    my ( $area, $period, $options, $table, $column, $rowName, $value );
    my @selectedColumns;
    foreach ( sort { $rank[$a] <=> $rank[$b]; } 0 .. $#rank ) {
        my $cat = int( $rank[$_] );
        last if $cat == 7;
        $selectedColumns[$cat] = $_
          unless defined $selectedColumns[$cat];
    }
    my @mapped = map {
        defined $selectedColumns[$_]
          ? $headers->[ $selectedColumns[$_] ]
          : 'Unmapped';
    } 0 .. 6;
    my $mapping = <<EOM;
    Area: $mapped[0]
    Period: $mapped[1]
    Options: $mapped[2]
    Table number: $mapped[3]
    Column number: $mapped[4]
    Row name: $mapped[5]
    Value: $mapped[6]
EOM
    if ( grep { !defined $selectedColumns[$_]; } 0, 3 .. 6 ) {
        die "Cannot import $fileName:\n$mapping";
    }
    warn "Importing $fileName using:\n$mapping";
    while ( !eof($blob) ) {
        if ( my $row = $csvParser->getline($blob) ) {
            my $book = join ' ', grep { defined $_ && length $_; }
              map  { "@{$row}[$_]"; }
              grep { defined $_; } @selectedColumns[ 0 .. 2 ];
            $book =~ tr/ /-/;
            $dataHandle->{$book}{ $row->[ $selectedColumns[3] ] }
              [ $row->[ $selectedColumns[4] ] ]
              { _normalisedRowName( $row->[ $selectedColumns[5] ] ) } =
              $row->[ $selectedColumns[6] ];
        }
    }

}

sub _normalisedRowName {
    ( local $_ ) = @_;
    s/[^A-Za-z0-9-]/ /g;
    s/- / /g;
    s/ +/ /g;
    s/^ //;
    s/ $//;
    $_;
}

1;
