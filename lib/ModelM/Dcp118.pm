package ModelM;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others.
Copyright 2012-2013 Franck Latrémolière, Reckon LLP and others.

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

sub adjust118 {

    my ( $model, $netCapexPercentages, $meavPercentages, ) = @_;

    my $assetLevelset = Labelset(
        name => 'All asset levels',
        list => [ split /\n/, <<EOT ] );
GSP
132kV circuits
132kV/EHV
EHV circuits
EHV/HV
132kV/HV
HV circuits
HV/LV
LV circuits
LV customer
HV customer
EOT

    my $cdcmAssets = $model->{objects}{cdcmAssets} ||= Dataset(
        name  => 'Assets in CDCM model (£)',
        lines => [
            'These data are taken from the CDCM tariff model (Otex sheet).',
            'They are also used as input data in the EDCM tariff model.',
        ],
        defaultFormat => '0hard',
        cols          => $assetLevelset,
        data          => [ map { 5e8 } @{ $assetLevelset->{list} } ],
        number        => 1331,
        dataset       => $model->{dataset},
        appendTo      => $model->{objects}{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        }
    );

    my $edcmAssets = $model->{objects}{edcmAssets} ||= Dataset(
        name          => 'All notional assets in EDCM (£)',
        lines         => 'These data are taken from the EDCM tariff model.',
        defaultFormat => '0hard',
        data          => [5e7],
        number        => 1332,
        dataset       => $model->{dataset},
        appendTo      => $model->{objects}{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        }
    );

    my $theCols = $netCapexPercentages->{cols};

# Duplicating $cdcmAssets for Numbers for iPad which cannot do SUMPRODUCT across sheets
    my $cdcmProp = Arithmetic(
        name => 'Proportion of EHV notional assets which are in the CDCM',
        defaultFormat => '%soft',
        cols =>
          Labelset( list => [ $theCols->{list}[ $#{ $theCols->{list} } ] ] ),
        arithmetic => '=1/(1+A1/SUMPRODUCT(A2_A3,A4_A5))',
        arguments  => {
            A1    => $edcmAssets,
            A2_A3 => Constant(
                name  => 'EHV asset levels',
                cols  => $assetLevelset,
                byrow => 1,
                data  => [ [qw(0 1 1 1 1 1 0 0 0 0 0)] ]
            ),
            A4_A5 => $cdcmAssets,
        },
    );

    my $propKept = Stack(
        name          => 'Proportion to be kept',
        cols          => $theCols,
        defaultFormat => '%copy',
        sources       => [
            $cdcmProp,
            Constant(
                name          => 'Default to 1',
                cols          => $theCols,
                defaultFormat => '%con',
                data          => [ map { [1] } @{ $theCols->{list} } ],
            )
        ],
    );

    map {
        Arithmetic(
            name       => "$_->{name} after DCP 118 exclusions",
            arithmetic => '=A1*A2/SUMPRODUCT(A3_A4,A5_A6)',
            arguments  => {
                A1    => $_,
                A2    => $propKept,
                A3_A4 => $_,
                A5_A6 => $propKept,
            },
            defaultFormat => '%soft'
        );
    } $netCapexPercentages, $meavPercentages;

}

1;
