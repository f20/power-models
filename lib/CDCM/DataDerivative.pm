package CDCM;

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

sub derivativeDataset {

    my ( $model, $sourceModel ) = @_;

    my $sourceTableHashref;
    my $setDatasetTable = sub {
        my ( $theTable, $formulaMaker ) = @_;
        my $hardData = $model->{dataset}{$theTable};
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
                my @rows = $d->{rows} ? @{ $d->{rows}{list} } : 'z';
                for ( my $i = 0 ; $i < @rows ; ++$i ) {
                    local $_ = $rows[$i];
                    s/.*\n//s;
                    s/[^A-Za-z0-9 -]/ /g;
                    s/- / /g;
                    s/ +/ /g;
                    s/^ //;
                    s/ $//;
                    my $row = $_;
                    $columns[$_]{$row} = $formulaMaker->(
                        "'$s'!" . xl_rowcol_to_cell( $r + $i, $c + $_ - 1 ),
                        $rows[$i], $_, $wb, $ws,
                    ) foreach 1 .. $width;
                }
            }
            map {
                for ( my $icolumn = 1 ; $icolumn < @$_ ; ++$icolumn ) {
                    foreach my $irow ( keys %{ $_->[$icolumn] } ) {
                        $columns[$icolumn]{$irow} =
                          $_->[$icolumn]{$irow};
                    }
                }
            } grep { $_ } $hardData, $model->{dataOverride}{$theTable};
            \@columns;
        };
    };

    $setDatasetTable->( defaultClosure => sub { my ($cell) = @_; "=$cell"; } );

    if (
        $model->{arpSharedData}
        && ( my $getAssumptionCell =
            $model->{arpSharedData}->assumptionsLocator( $model, $sourceModel )
        )
      )
    {

        $setDatasetTable->(
            1001 => sub {
                my ( $cell, $row, $col, $wb, $ws ) = @_;
                return "=$cell"
                  unless $col == 4 && $row =~ /RPI Indexation Factor/i;
                my $ac = $getAssumptionCell->( $wb, $ws, 'RPI' );
                "=(1+$ac)*$cell";
            }
        );

        $setDatasetTable->(
            1020 => sub {
                ( my $cell, local $_, my $col, my $wb, my $ws ) = @_;
                my $ac = $getAssumptionCell->(
                    $wb, $ws,
                    /132kV\/EHV/     ? 2
                    : /132kV\/HV/    ? 5
                    : /132kV/        ? 1
                    : /EHV\/HV/      ? 4
                    : /EHV/          ? 3
                    : /HV\/LV/       ? 8
                    : /HV/           ? 6
                    : /LV circuits/i ? 9
                    :                  10
                );
                "=(1+$ac)*$cell";
            }
        );

        {
            my $ac;
            $setDatasetTable->(
                1022 => sub {
                    ( my $cell, local $_, my $col, my $wb, my $ws ) = @_;
                    $ac ||= $getAssumptionCell->( $wb, $ws, 10 );
                    "=(1+$ac)*$cell";
                }
            );
        }

        {
            my $ac;
            $setDatasetTable->(
                1023 => sub {
                    ( my $cell, local $_, my $col, my $wb, my $ws ) = @_;
                    $ac ||= $getAssumptionCell->( $wb, $ws, 7 );
                    "=(1+$ac)*$cell";
                }
            );
        }

        $setDatasetTable->(
            1053 => sub {
                ( my $cell, local $_, my $col, my $wb, my $ws ) = @_;
                my $ac = $getAssumptionCell->(
                    $wb, $ws,
                    $col < 4
                    ? ( /unmet/i ? 21 : /gener/i ? 22 : /half[ -]hourly/i
                          && !/aggreg/i ? 17 : 15 )
                    : $col == 4 ? ( /gener/i ? 23 : /half[ -]hourly/i
                          && !/aggreg/i ? 18 : 16 )
                    : $col == 5 ? 19
                    : ( /gener/i ? 24 : 20 ),
                );
                "=(1+$ac)*$cell";
            }
        );

        $setDatasetTable->(
            1055 => sub {
                ( my $cell, local $_, my $col, my $wb, my $ws ) = @_;
                my $ac = $getAssumptionCell->( $wb, $ws, 14 );
                "=(1+$ac)*$cell";
            }
        );

        $setDatasetTable->(
            1059 => sub {
                ( my $cell, local $_, my $col, my $wb, my $ws ) = @_;
                my $no;
                $no = 11 if $col == 1;
                $no = 12 if $col == 2;
                $no = 13 if $col == 4;
                join '', '=',
                  $no
                  ? ( '(1+', $getAssumptionCell->( $wb, $ws, $no ), ')*' )
                  : (), $cell;
            }
        );

    }

    $setDatasetTable->( $_ => sub { my ($cell) = @_; "=$cell"; } )
      foreach grep { /^[0-9]+$/s } keys %{ $model->{dataset} };

}

1;
