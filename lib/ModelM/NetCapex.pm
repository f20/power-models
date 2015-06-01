package ModelM;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others.
Copyright 2012-2014 Franck Latrémolière, Reckon LLP and others.

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

sub netCapexRawData {
    my ($model) = @_;
    return unless $model->{netCapex};
    $model->{objects}{netCapexRawData} ||= Dataset(
        name          => 'Net capex analysis pre-DCP 118 (£)',
        defaultFormat => '0hard',
        lines => 'In a pre-DCP 118 legacy Method M workbook, these data are on'
          . ' sheet Calc-Net capex, possibly cells G6 to G10.',
        data       => [qw(100 100 100 100 100)],
        number     => 1369,
        rows       => Labelset( list => [qw(LV LV/HV HV EHV 132kV)] ),
        dataset    => $model->{dataset},
        appendTo   => $model->{objects}{inputTables},
        validation => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    );
}

sub netCapexPercentages {
    my ( $model, $allocLevelset ) = @_;
    my $netCapex = $model->netCapexRawData;
    return $model->{objects}{netCapexPercentages}{ 0 + $allocLevelset } ||=
      Stack(    # for Numbers for iPad which cannot do SUMPRODUCT across sheets
        sources => [
            Dataset(
                name => 'Net capex percentages',
                lines =>
                  'In a pre-DCP 118 legacy Method M workbook, these data are on'
                  . ' sheet Calc-Net capex, possibly cells H6 to H10.',
                data          => [ map { 0 } @{ $allocLevelset->{list} } ],
                defaultFormat => '%hard',
                number        => 1370,
                cols       => $allocLevelset,
                dataset    => $model->{dataset},
                appendTo   => $model->{objects}{inputTables},
                validation => {
                    validate => 'decimal',
                    criteria => '>=',
                    value    => 0,
                },
            )
        ]
      ) unless $netCapex;
    $model->{objects}{netCapexPercentages}{ 0 + $netCapex }
      { 0 + $allocLevelset } ||= SpreadsheetModel::Custom->new(
        name          => 'Net capex percentages',
        defaultFormat => '%soft',
        cols          => $allocLevelset,
        arithmetic    => '=(IV5 or IV6+IV7)/SUM(IV1:IV2)',
        custom =>
          [ '=IV6/(IV1+IV2+IV3+IV4+IV5)', '=(IV6+IV7)/(IV1+IV2+IV3+IV4+IV5)', ],
        arguments => {
            IV1 => $netCapex,
            IV2 => $netCapex,
            IV3 => $netCapex,
            IV4 => $netCapex,
            IV5 => $netCapex,
            IV6 => $netCapex,
            IV7 => $netCapex,
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[ $x == 3 ? 1 : 0 ], map {
                    $_ => Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + (
                              /IV([12345])/ ? $1 - 1
                            : /IV([67])/    ? $x + $1 - 6
                            : 0
                        ),
                        $colh->{$_},
                        1, 1,
                      )
                } @$pha;
            };
        },
      );
}

sub netCapexPercentageServiceLV {
    my ( $model, $lvOnly, $lvServiceOnly ) = @_;
    $model->{objects}{netCapexPercentageServiceLV}{ 0 + $lvOnly }
      { 0 + $lvServiceOnly } ||= Dataset(
        name => 'Net capex: ratio of LV services to LV total',
        lines =>
          q%Calculated as SUM('FBPQ NL1'!D10:M13)/SUM('FBPQ NL1'!D10:M16).%,
        data          => [.5],
        defaultFormat => '%hard',
        number        => 1380,
        cols          => $lvOnly,
        rows          => $lvServiceOnly,
        dataset       => $model->{dataset},
        appendTo      => $model->{objects}{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
      );
}

1;
