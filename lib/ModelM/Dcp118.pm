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

    my $cdcmAssets = Dataset(
        name => 'Assets in CDCM model (£) (from CDCM table 2705 or 2706)',
        defaultFormat => '0hard',
        cols          => $assetLevelset,
        data          => [ map { 5e8 } @{ $assetLevelset->{list} } ],
        number        => 1331,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        }
    );

    my $edcmAssets = Dataset(
        name => 'All notional assets in EDCM (£) '
          . '(from EDCM table 4167, 4168, 4169 or 4170)',
        defaultFormat => '0hard',
        data          => [5e7],
        number        => 1332,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables},
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
        arithmetic => '=1/(1+IV1/SUMPRODUCT(IV2_IV3,IV4_IV5))',
        arguments  => {
            IV1     => $edcmAssets,
            IV2_IV3 => Constant(
                name  => 'EHV asset levels',
                cols  => $assetLevelset,
                byrow => 1,
                data  => [ [qw(0 1 1 1 1 1 0 0 0 0 0)] ]
            ),
            IV4_IV5 => $cdcmAssets,
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
        # for Numbers for iPad which cannot do SUMPRODUCT across sheets
        my $dup = Stack( sources => [$_] );
        Arithmetic(
            name       => "$_->{name} after exclusions",
            arithmetic => '=IV1*IV2/SUMPRODUCT(IV3_IV4,IV5_IV6)',
            arguments  => {
                IV1     => $dup,
                IV2     => $propKept,
                IV3_IV4 => $dup,
                IV5_IV6 => $propKept,
            },
            defaultFormat => '%soft'
        );
    } $netCapexPercentages, $meavPercentages;

}

1;
