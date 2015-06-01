package ModelM;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others.
Copyright 2012-2015 Franck Latrémolière, Reckon LLP and others.

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

use SpreadsheetModel::Shortcuts ':all';

sub units {

    my ( $model, $allocLevelset ) = @_;
    my $key = join '&', 'units?allocLevelset=' . ( 0 + $allocLevelset ),
      map { $model->{$_} ? "$_=1" : (); } qw(calcUnits);

    return $model->{objects}{$key} if $model->{objects}{$key};

    return $model->{objects}{$key} = Dataset(
        name  => 'Units flowing',
        lines => 'In a legacy Method M workbook, these data are on'
          . ' sheet Calc-Units, cells C23, C23, D23, E23.',
        data          => [ map { 100 } @{ $allocLevelset->{list} } ],
        defaultFormat => '0hard',
        number        => 1320,
        cols          => $allocLevelset,
        dataset    => $model->{dataset},
        appendTo   => $model->{objects}{inputTables},
        validation => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    ) unless $model->{calcUnits};

    my $distributed = Dataset(
        name  => 'Units distributed (GWh)',
        lines => 'These data are taken from the'
          . ' 2007/2008 regulatory reporting pack (table 5.1), cells G34 to G36.',
        rows => Labelset( list => [ 'EHV (Includes 132kV)', 'HV', 'LV' ] ),
        data       => [qw(2000 5000 25000)],
        number     => 1321,
        dataset    => $model->{dataset},
        appendTo   => $model->{objects}{inputTables},
        validation => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    );

    my $losses = Dataset(
        name  => 'Losses (GWh)',
        lines => 'This data item is taken from the'
          . ' 2007/2008 regulatory reporting pack (table 5.1), cell G40.',
        data       => [2500],
        number     => 1322,
        dataset    => $model->{dataset},
        appendTo   => $model->{objects}{inputTables},
        validation => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    );

    require Spreadsheet::WriteExcel::Utility;
    my $factors = SpreadsheetModel::Custom->new(
        name          => 'Adjustment factors to LV (kWh/GWh)',
        cols          => $allocLevelset,
        rows          => $distributed->{rows},
        defaultFormat => '0soft',
        arithmetic =>
'=1000000*(1+IV6/(IV22+IV21/2+IV20/4)/[1, 2 or 4])/(1+IV7/(IV32+IV31/2+IV30/4))',
        custom => [
'=1000000*(1+IV6/(IV22+IV21/2+IV20/4)/4)/(1+IV7/(IV32+IV31/2+IV30/4))',
'=1000000*(1+IV6/(IV22+IV21/2+IV20/4)/2)/(1+IV7/(IV32+IV31/2+IV30/4))',
            '=1000000',
        ],
        arguments => {
            IV20 => $distributed,
            IV21 => $distributed,
            IV22 => $distributed,
            IV30 => $distributed,
            IV31 => $distributed,
            IV32 => $distributed,
            IV6  => $losses,
            IV7  => $losses,
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                return '', $wb->getFormat('0con'), '0'
                  if $y == 1 && $x < 2 || $y == 0 && $x < 3;
                '', $format, $formula->[$y], map {
                    $_ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + ( /IV[23]([012])/ ? $1 : 0 ),
                        $colh->{$_}, 1, 0, )
                } @$pha;
            };
        },
    );

    $model->{objects}{$key} = SumProduct(
        name          => 'Units flowing, loss adjusted to LV (kWh)',
        defaultFormat => '0soft',
        matrix        => $factors,
        vector        => $distributed,
    );

}

1;
