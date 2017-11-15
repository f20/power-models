package Impact;

=head Copyright licence and disclaimer

Copyright 2017 Franck Latrémolière, Reckon LLP and others.

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

sub processTables6001and6002 {

    my ( $model, $baselineData, $scenarioData, $areaName, $sheetTitle, ) = @_;

    $model->{endUserLevelset} ||= Labelset(
        list => [
            "LV demand\t",
            "LV Sub demand\t(LV gen EDCM)",
            "HV demand\t(LVS gen EDCM)",
            "\t(HV gen EDCM)",
        ]
    );

    my $bd_edcm = $baselineData->{6001} or die;
    my $bd_cdcm = $baselineData->{6002} or die;
    my $sd_edcm = $scenarioData->{6001} or die;
    my $sd_cdcm = $scenarioData->{6002} or die;

    my $boundaryLevelset = Labelset(
        list => [
            'Boundary LV', 'Boundary HV',
            map { $sd_edcm->[0][ 5 - $_ ]; } 0 .. 4
        ]
    );

    my $baseline = Constant(
        name          => 'Baseline discount',
        defaultFormat => '%copy',
        cols          => $model->{endUserLevelset},
        rows          => $boundaryLevelset,
        data          => [
            [
                $bd_cdcm->[2][1], $bd_cdcm->[3][1],
                map { $bd_edcm->[1][ 5 - $_ ]; } 0 .. 4
            ],
            [
                undef, $bd_cdcm->[4][1],
                map { $bd_edcm->[2][ 5 - $_ ]; } 0 .. 4
            ],
            [
                undef, $bd_cdcm->[5][1],
                map { $bd_edcm->[3][ 5 - $_ ]; } 0 .. 4
            ],
            [ undef, undef, undef, map { $bd_edcm->[4][ 5 - $_ ]; } 1 .. 4 ],
        ],
    );

    my $scenario = Constant(
        name          => 'Scenario discount',
        defaultFormat => '%copy',
        cols          => $model->{endUserLevelset},
        rows          => $boundaryLevelset,
        data          => [
            [
                $sd_cdcm->[2][1], $sd_cdcm->[3][1],
                map { $sd_edcm->[1][ 5 - $_ ]; } 0 .. 4
            ],
            [
                undef, $sd_cdcm->[4][1],
                map { $sd_edcm->[2][ 5 - $_ ]; } 0 .. 4
            ],
            [
                undef, $sd_cdcm->[5][1],
                map { $sd_edcm->[3][ 5 - $_ ]; } 0 .. 4
            ],
            [ undef, undef, undef, map { $sd_edcm->[4][ 5 - $_ ]; } 1 .. 4 ],
        ],
    );

    my $change = Arithmetic(
        name          => 'Change in LDNO discounts',
        defaultFormat => '%softpm',
        arithmetic    => '=A2-A1',
        arguments     => { A1 => $baseline, A2 => $scenario, },
    );

    SpreadsheetModel::MatrixSheet->new(
        noLines       => 1,
        noDoubleNames => 1,
        noNumbers     => 1,
      )->addDatasetGroup(
        name    => 'Baseline LDNO discount percentages',
        columns => [$baseline],
      )->addDatasetGroup(
        name    => 'Scenario LDNO discount percentages',
        columns => [$scenario],
      )->addDatasetGroup(
        name    => 'Change in LDNO discount percentages',
        columns => [$change],
      );

    if ($sheetTitle) {
        push @{ $model->{worksheetsAndClosures} }, $areaName => sub {
            my ( $wsheet, $wbook ) = @_;
            $wsheet->set_column( 0, 0,   48 );
            $wsheet->set_column( 1, 254, 16 );
            $wsheet->freeze_panes( 5, 1 );
            $_->wsWrite( $wbook, $wsheet ),
              foreach Notes( name => $sheetTitle ), $baseline, $scenario,
              $change;
        };
    }
    else {
        push @{ $model->{columnsetFilterFood} },
          [ $areaName, $baseline, $scenario, $change, ];
    }

}

1;
