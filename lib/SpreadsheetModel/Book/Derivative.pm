package SpreadsheetModel::Book::Derivative;

=head Copyright licence and disclaimer

Copyright 2014 Franck Latrémolière, Reckon LLP and others.

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
use Spreadsheet::WriteExcel::Utility;

sub registerSourceModel {

    my ( $model, $sourceModel ) = @_;

    my @backupDatasets;
    {
        my $m = $sourceModel;
        while ( my $d = $m->{dataset} ) {
            push @backupDatasets, $d;
            $m = $m->{sourceModel};
        }
    }

    my $sourceTableHashref;
    sub {

        my ( $theTable, $formulaMaker ) = @_;
        my $hardData = $model->{dataset} ? $model->{dataset}{$theTable} : undef;
        return if ref $hardData eq 'CODE';

        $model->{dataset}{$theTable} = sub {

            my ( $table, $wb, $ws ) = @_;
            $sourceTableHashref ||=
              { map { $_->{number} => $_ } @{ $sourceModel->{inputTables} } };
            my @columns;

            if ( my $d = $sourceTableHashref->{$table} ) {
                my ( $s, $r, $c ) =
                  ( $d->{columns} ? $d->{columns}[0] : $d )
                  ->wsWrite( $wb, $ws );
                $s = $s->get_name;
                my $width = 0;
                if ( $d->{columns} ) {
                    $width += $_->lastCol + 1 foreach @{ $d->{columns} };
                }
                else {
                    $width = 1 + $d->lastCol;
                }
                my @rows =
                  $d->{rows}
                  ? @{ $d->{rows}{list} }
                  : 'MAGICAL SINGLE ROW NAME';

                for ( my $i = 0 ; $i < @rows ; ++$i ) {

                    local $_ = $rows[$i];
                    s/.*\n//s;
                    s/[^A-Za-z0-9 -]/ /g;
                    s/- / /g;
                    s/ +/ /g;
                    s/^ //;
                    s/ $//;
                    my $row = $_;

                    foreach my $col ( 1 .. $width ) {
                        $columns[$col]{$row} = $formulaMaker->(
                            "'$s'!"
                              . xl_rowcol_to_cell( $r + $i, $c + $col - 1 ),
                            $rows[$i], $col, $wb, $ws, $row
                        );
                    }
                }
            }

            map {
                for ( my $icolumn = 1 ; $icolumn < @$_ ; ++$icolumn ) {
                    foreach my $irow ( keys %{ $_->[$icolumn] } ) {
                        $columns[$icolumn]{$irow} =
                          $_->[$icolumn]{$irow};
                    }
                }
              } grep { $_ } $hardData,
              map    { $_->{$theTable} }
              ref $model->{dataOverride} eq 'ARRAY'
              ? @{ $model->{dataOverride} }
              : $model->{dataOverride};

            map {
                for ( my $icolumn = 1 ; $icolumn < @$_ ; ++$icolumn ) {
                    foreach my $irow ( keys %{ $_->[$icolumn] } ) {
                        $columns[$icolumn]{$irow} = $_->[$icolumn]{$irow}
                          unless
                          exists $columns[$icolumn]{'MAGICAL SINGLE ROW NAME'}
                          || exists $columns[$icolumn]{$irow};
                    }
                }
              }
              grep { ref $_ eq 'ARRAY' }
              map  { $_->{$theTable} } @backupDatasets;

            \@columns;

        };

    };

}

1;
