package Impact;

=head Copyright licence and disclaimer

Copyright 2014-2017 Franck Latrémolière, Reckon LLP and others.

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
use SpreadsheetModel::MatrixSheet;

sub processTables6001and6002 {

    my ( $model, $dno, $title, $baselineData, $scenarioData, ) = @_;
    my $bd_edcm = $baselineData->{6001} or die;
    my $bd_cdcm = $baselineData->{6002} or die;
    my $sd_edcm = $scenarioData->{6001} or die;
    my $sd_cdcm = $scenarioData->{6002} or die;

    my $discountsetNotUsed = Labelset(
        list => [
            ( map { $_->[0]; } @$sd_cdcm[ 2 .. 5 ] ),
            (
                map {
                    my $row = 5 - $_;
                    map {
                        local $_ = "$sd_edcm->[0][$row]: $sd_edcm->[$_][0]";
                        s/^Boundary/LDNO/;
                        $_;
                    } 1 .. ( $_ ? 4 : 3 );
                } 0 .. 4
            )
        ]
    );

    my $endUserLevelset = Labelset(
        list => [
            'LV demand',
            'LV Sub demand and LV generation EDCM',
            'HV demand and LV Sub generation EDCM',
            'HV generation EDCM',
        ]
    );

    my $boundaryLevelset = Labelset(
        list => [
            'Boundary LV', 'Boundary HV',
            map { $sd_edcm->[0][ 5 - $_ ]; } 0 .. 4
        ]
    );

    my $baseline = Constant(
        name          => '',
        defaultFormat => '%copy',
        cols          => $endUserLevelset,
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
        name          => '',
        defaultFormat => '%copy',
        cols          => $endUserLevelset,
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

    SpreadsheetModel::MatrixSheet->new( verticalSpace => 2 )->addDatasetGroup(
        name    => 'Baseline LDNO discount percentages',
        columns => [$baseline],
      )->addDatasetGroup(
        name    => 'Scenario LDNO discount percentages',
        columns => [$scenario],
      )->addDatasetGroup(
        name    => 'Change in LDNO discount percentages',
        columns => [$change],
      );

    push @{ $model->{sheetNames} }, $dno;
    push @{ $model->{sheetTables} },
      [ Notes( name => $title ), $baseline, $scenario, $change, ];

}

1;
