package CDCM;

# Copyright 2018 Franck Latrémolière, Reckon LLP and others.
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
use SpreadsheetModel::Shortcuts ':all';

sub tcdbClosure {
    my ( $model, $wbook ) = @_;
    sub {
        my ($wsheet) = @_;
        $wsheet->set_column( 0, 0, 96 );
        $wsheet->set_column( 1, 1, 48 );
        $wsheet->set_column( 2, 2, 32 );
        $wsheet->set_column( 3, 3, 16 );
        my $wsRow = $wsheet->{nextFree} || 0;
        my @cols = map { $_->objectShortName; } @{ $model->{allTariffColumns} };
        my @tableLocs = map {
            my ( $s, $r, $c ) = $_->wsWrite( $wbook, $wsheet );
            [ q%='% . $s->get_name . q%'!%, $r, $c ];
        } @{ $model->{allTariffColumns} };
        my $formatth = $wbook->getFormat('th');
        my @formats =
          map { $wbook->getFormat( $_->{defaultFormat} || '0.000copy' ); }
          @{ $model->{allTariffColumns} };
        my $rowsar = $model->{allTariffColumns}[0]{rows}{list};
        for ( my $y = 0 ; $y < @$rowsar ; ++$y ) {
            my $row = $rowsar->[$y];
            $row =~ s/^.*\n//s;
            next if $row =~ /LDNO|QNO/;
            for ( my $x = 0 ; $x < @cols ; ++$x ) {
                next
                  unless $model->{componentMapForTcdb}{ $rowsar->[$y] }
                  { $cols[$x] };
                $wsheet->write( $wsRow, 0, $model->{nickNames}{$wbook},
                    $formatth );
                $wsheet->write( $wsRow, 1, $row,      $formatth );
                $wsheet->write( $wsRow, 2, $cols[$x], $formatth );
                $wsheet->write(
                    $wsRow, 3,
                    $tableLocs[$x][0]
                      . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $tableLocs[$x][1] + $y,
                        $tableLocs[$x][2]
                      ),
                    $formats[$x]
                );
                ++$wsRow;
            }
        }
        $wsheet->{nextFree} = $wsRow;
    };
}

1;
