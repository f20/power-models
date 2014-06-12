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
use utf8;
use SpreadsheetModel::Shortcuts ':all';

sub makeStatisticsAssumptions {

    my ($model) = @_;

    return @{ $model->{arpSharedData}{statisticsAssumptionsStructure} }
      if $model->{arpSharedData}
      && $model->{arpSharedData}{statisticsAssumptionsStructure};

    my $kw6524 = 65 * 365.25 * 24;
    my $kw6511 = 65 * 365.25 * 11;

    my @rows =
      split /\n/, <<EOL;
Domestic Unrestricted | Low usage,1900
Domestic Unrestricted | Medium usage,3800
Domestic Unrestricted | High usage,7600
Small Non Domestic Unrestricted | Medium usage,$kw6511
Small Non Domestic Unrestricted | High usage,$kw6524
LV Network Non-Domestic Non-CT | Off peak load,,,0,0,715
LV Network Non-Domestic Non-CT | Continuous load,,,65
LV HH Metered | Small off peak load,,69,0,0,715
LV HH Metered | Small continuous load,,69,65
HV HH Metered | Continuous load,,5000,4500
HV HH Metered | Off peak load,,5000,0,0,49500
HV HH Metered | Peaky load,,5000,2000,2500,0
EOL

    my @defaultDataColumns = map { [] } 0 .. 4;

    for ( my $i = 0 ; $i < @rows ; ++$i ) {
        my @a = split /,\s*/, $rows[$i];
        $rows[$i] = shift @a;
        for ( my $j = 0 ; $j < @a ; ++$j ) {
            $defaultDataColumns[$j][$i] = $a[$j]
              if defined $a[$j] && length $a[$j];
        }
    }

    my $rowset = Labelset( list => \@rows );

    my @columns;

    push @columns,
      Dataset(
        name          => 'Annual consumption (kWh)',
        defaultFormat => '0hardnz',
        rows          => $rowset,
        data          => [ $defaultDataColumns[0] ],
      );

    push @columns,
      Dataset(
        name          => 'Capacity (kVA)',
        defaultFormat => '0hardnz',
        rows          => $rowset,
        data          => [ $defaultDataColumns[1] ],
      );

    push @columns,
      Dataset(
        name          => 'Continuous usage (kW)',
        defaultFormat => '0hardnz',
        rows          => $rowset,
        data          => [ $defaultDataColumns[2] ],
      );

    push @columns,
      Dataset(
        name          => 'Additional red usage (kW)',
        defaultFormat => '0hardnz',
        rows          => $rowset,
        data          => [ $defaultDataColumns[3] ],
      );

    push @columns,
      Dataset(
        name          => 'Additional green usage (kWh/day)',
        defaultFormat => '0hardnz',
        rows          => $rowset,
        data          => [ $defaultDataColumns[4] ],
      );

    Columnset(
        name     => 'Assumed usage for illustrative customers',
        number   => 1202,
        appendTo => $model->{arpSharedData}
        ? $model->{arpSharedData}{statsAssumptions}
        : $model->{inputTables},
        dataset => $model->{dataset},
        columns => \@columns,
    );

    $model->{arpSharedData}{statisticsAssumptionsStructure} =
      [ \@rows, @columns ]
      if $model->{arpSharedData};

    \@rows, @columns;

}

sub makeStatisticsTables {

    my ( $model, $tariffTable, $daysInYear, $nonExcludedComponents,
        $componentMap, $allTariffs, $unitsInYear, )
      = @_;

    my ( $rows, $annualUnits, $capacity, $generalkW, $redkW, $green, ) =
      $model->makeStatisticsAssumptions;

    my @tariffTableRows = map {
        my $s = $_;
        $s =~ s/ *\|.*//;
        $s = qr($s);
        $s = (
            (
                grep {
                    local $_ = $allTariffs->{list}[$_];
                    !/^>/ && /$s/;
                } 0 .. $#{ $allTariffs->{list} }
            )[0]
        );
    } @$rows;

    my $stats = SpreadsheetModel::Custom->new(
        name          => 'Annual charges for illustrative customers (£/year)',
        defaultFormat => '0softnz',
        rows          => $annualUnits->{rows},
        custom        => [
            '=0.01*(IV1*IV91+IV71*IV94)',
            '=0.01*('
              . '(IV31+IV4)*IV61*IV91+IV32*IV62*IV92+(IV33*IV63+IV5*IV72)*IV93'
              . '+IV71*(IV94+IV2*IV95))'
        ],
        arguments => {
            IV1  => $annualUnits,
            IV2  => $capacity,
            IV31 => $generalkW,
            IV32 => $generalkW,
            IV33 => $generalkW,
            IV4  => $redkW,
            IV5  => $green,
            IV61 => $model->{hoursByRedAmberGreen},
            IV62 => $model->{hoursByRedAmberGreen},
            IV63 => $model->{hoursByRedAmberGreen},
            IV71 => $daysInYear,
            IV72 => $daysInYear,
            IV91 => $tariffTable->{'Unit rate 1 p/kWh'},
            IV92 => $tariffTable->{'Unit rate 2 p/kWh'},
            IV93 => $tariffTable->{'Unit rate 3 p/kWh'},
            IV94 => $tariffTable->{'Fixed charge p/MPAN/day'},
            IV95 => $tariffTable->{'Capacity charge p/kVA/day'},
        },
        rowFormats => [ map { $_ ? undef : 'unavailable'; } @tariffTableRows ],
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                my $cellFormat =
                    $self->{rowFormats}[$y]
                  ? $wb->getFormat( $self->{rowFormats}[$y] )
                  : $format;
                my $tariffY = $tariffTableRows[$y];
                return '', $cellFormat unless $tariffY;
                my $tariff = $allTariffs->{list}[$tariffY];
                my $style  = $tariff !~ /gener/i
                  && !$componentMap->{$tariff}{'Unit rate 2 p/kWh'} ? 0 : 1;
                '', $cellFormat, $formula->[$style], map {
                    $_ => Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + (
                              /^IV9/         ? $tariffY
                            : /^IV(?:[1-5])/ ? $y
                            : 0
                        ),
                        $colh->{$_} +
                          ( $_ eq 'IV62' ? 1 : $_ eq 'IV63' ? 2 : 0 ),
                        1, 1,
                      )
                } @$pha;
            };
        },
    );

    $model->{arpSharedData}
      ->addStats( 'Annual charges for illustrative customers (£/year)',
        $model, $stats )
      if $model->{arpSharedData};

    push @{ $model->{statisticsTables} }, $stats;

}

1;
