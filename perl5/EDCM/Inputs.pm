package EDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and contributors. All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted
provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of
conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of
conditions and the following disclaimer in the documentation and/or other materials provided
with the distribution.

THIS SOFTWARE IS PROVIDED BY ENERGY NETWORKS ASSOCIATION LIMITED AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL ENERGY
NETWORKS ASSOCIATION LIMITED OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF
THE POSSIBILITY OF SUCH DAMAGE.

=cut

=head Table numbers used in this file

911
912
913
953
1101
1111
1112
1114

=cut

use warnings;
use strict;
use utf8;

use SpreadsheetModel::Shortcuts ':all';

sub generalInputs {

    my ($model) = @_;

    my $days = Dataset(
        name          => 'Days in year',
        defaultFormat => '0hard',
        data          => [365],
        dataset       => $model->{dataset},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        }
    );

    $model->{daysInYear} = $days;

    my $tExit = Dataset(
        name          => 'Transmission exit charges (£/year)',
        defaultFormat => '0hard',
        data          => [10e6],
        dataset       => $model->{dataset},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        }
    );

    my $peakLoad = Constant(
        name          => 'Not used',
        defaultFormat => '0hardnz',
        data          => [],
        dataset       => $model->{dataset},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        }
    );

    my $allowedRev = Dataset(
        name => 'The amount of money that the DNO wants to raise from use of system charges, less transmission exit (£/year)',
        defaultFormat => '0hard',
        data          => [300e6],
        dataset       => $model->{dataset},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        }
    );

    my $direct = Dataset(
        name          => 'Direct cost (£/year)',
        defaultFormat => '0hard',
        data          => [40_000_000],
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        }
    );

    my $indirect = Dataset(
        name          => 'Indirect cost (£/year)',
        defaultFormat => '0hard',
        data          => [60_000_000],
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        }
    );

    my $indirectProp = Constant(
        name => 'Indirect cost proportion',

        #        defaultFormat => '%hard',
        data       => [1.0],
        validation => {
            validate => 'decimal',
            criteria => 'between',
            minimum  => 0,
            maximum  => 1,
        }
    );

    my $rates = Dataset(
        name          => 'Network rates (£/year)',
        defaultFormat => '0hard',
        data          => [20_000_000],
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        }
    );

    my $ehvIntensity = Constant(
        name => 'EHV operating expenditure intensity',

        #        defaultFormat => '%hard',
        data       => [0.68],
        validation => {
            validate => 'decimal',
            criteria => 'between',
            minimum  => 0,
            maximum  => 1,
        }
    );

    my $scope = Dataset(
        name          => 'Applicability of EDCM',
        lines         => 'Comma-delimited list of customer classes',
        defaultFormat => 'texthard',
        data          => ['A,B'],
        number        => 1101,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables}
    );

    my $powerFactorInModel = Constant(
        name       => 'Power factor in 500 MW model',
        validation => {
            validate => 'decimal',
            criteria => 'between',
            minimum  => 0.001,
            maximum  => 1,
        },
        data => [ $model->{powerFactor} || 0.95 ]
    );

    Columnset(
        name    => 'Miscellaneous parameters',
        columns => [
            $days,         $indirectProp,
            $ehvIntensity, $powerFactorInModel ? $powerFactorInModel : ()
        ],
        number   => 1111,
        dataset  => $model->{dataset},
        appendTo => $model->{inputTables}
    );

    Columnset(
        name    => 'DNO revenue, expenditure and load data',
        columns => [
            $allowedRev, $peakLoad,
            $tExit,      $direct,
            $indirect ? $indirect : (), $rates ? $rates : ()
        ],
        number   => 1112,
        dataset  => $model->{dataset},
        appendTo => $model->{inputTables}
    );

    $scope, $days, $direct, $indirect, $indirectProp, $rates, $ehvIntensity,
      $tExit, $peakLoad, $allowedRev, $powerFactorInModel,
      Labelset(
        name => 'EHV asset levels',
        list => [ split /\n/, <<EOT ] );
GSP
132kV circuits
132kV/EHV
EHV circuits
EHV/HV
132kV/HV
EOT

}

sub loadFlowInputs {

    my ($model) = @_;

    return if $model->{method} eq 'none';

    $model->{numLocations} ||= 13;

    $model->{locationSet} = Labelset(
        name          => 'Locations',
        list          => [ 1 .. $model->{numLocations} ],
        defaultFormat => 'thloc'
    );

    $model->{locations} = Dataset(
        name          => 'Location name/ID',
        rows          => $model->{locationSet},
        data          => [ map { 'Not used' } 1 .. $model->{numLocations} ],
        defaultFormat => 'texthard',
        dataset       => $model->{dataset}
    );

    $model->{locDorG} = Dataset(
        name          => 'Demand or Generation',
        defaultFormat => 'texthard',
        data          => [ map { '' } 1 .. $model->{numLocations} ],
        rows          => $model->{locationSet},
        dataset       => $model->{dataset}
      )
      if $model->{method} =~ /LRIC/i;

    $model->{level} = Dataset(
        name          => Label('Level'),
        rows          => $model->{locationSet},
        data          => [ map { 3 } 1 .. $model->{numLocations} ],
        defaultFormat => '0hardnz',
        dataset       => $model->{dataset}
      )
      unless $model->{method} =~ /LRIC/i;

    $model->{parent} = Dataset(
        name => Label(
            $model->{method} =~ /LRIC/i
            ? 'Linked location (if any)'
            : 'Parent location (if any)'
        ),
        rows          => $model->{locationSet},
        data          => [ map { '' } 1 .. $model->{numLocations} ],
        defaultFormat => 'texthard',
        dataset       => $model->{dataset}
    );

    $model->{MCpeak} =
         $model->{method} =~ /LRIC/i
      && $model->{method} =~ /split/i
      ? [
        Dataset(
            name          => Label('Local charge 1 £/kVA/year'),
            rows          => $model->{locationSet},
            data          => [ map { 0 } 1 .. $model->{numLocations} ],
            defaultFormat => '0.000hard',
            dataset       => $model->{dataset}
        ),
        Dataset(
            name          => Label('Remote charge 1 £/kVA/year'),
            rows          => $model->{locationSet},
            data          => [ map { 0 } 1 .. $model->{numLocations} ],
            defaultFormat => '0.000hard',
            dataset       => $model->{dataset}
        )
      ]
      : Dataset(
        name          => Label('Charge 1 £/kVA/year'),
        rows          => $model->{locationSet},
        data          => [ map { 0 } 1 .. $model->{numLocations} ],
        defaultFormat => '0.000hard',
        dataset       => $model->{dataset}
      );

    $model->{MCoffp} =
         $model->{method} =~ /LRIC/i
      && $model->{method} =~ /split/i
      ? [
        Dataset(
            name          => Label('Local charge 2 £/kVA/year'),
            rows          => $model->{locationSet},
            data          => [ map { 0 } 1 .. $model->{numLocations} ],
            defaultFormat => '0.000hard',
            dataset       => $model->{dataset}
        ),
        Dataset(
            name          => Label('Remote charge 2 £/kVA/year'),
            rows          => $model->{locationSet},
            data          => [ map { 0 } 1 .. $model->{numLocations} ],
            defaultFormat => '0.000hard',
            dataset       => $model->{dataset}
        )
      ]
      : Dataset(
        name          => Label('Charge 2 £/kVA/year'),
        rows          => $model->{locationSet},
        data          => [ map { 0 } 1 .. $model->{numLocations} ],
        defaultFormat => '0.000hard',
        dataset       => $model->{dataset}
      );

    $model->{kWpeakD} = Dataset(
        name => Label(
            $model->{method} =~ /LRIC/i
            ? 'Maximum demand run: kW'
            : 'Maximum demand run: load kW'
        ),
        rows          => $model->{locationSet},
        data          => [ map { 0 } 1 .. $model->{numLocations} ],
        defaultFormat => '0hard',
        dataset       => $model->{dataset}
    );

    $model->{kVArpeakD} = Dataset(
        name => Label(
            $model->{method} =~ /LRIC/i
            ? 'Maximum demand run: kVAr'
            : 'Maximum demand run: load kVAr'
        ),
        rows          => $model->{locationSet},
        data          => [ map { 0 } 1 .. $model->{numLocations} ],
        defaultFormat => '0hard',
        dataset       => $model->{dataset}
    );

    $model->{kWpeakG} = Dataset(
        name          => Label('Maximum demand run: generation kW'),
        rows          => $model->{locationSet},
        data          => [ map { 0 } 1 .. $model->{numLocations} ],
        defaultFormat => '0hard',
        dataset       => $model->{dataset}
    );

    $model->{kVArpeakG} = Dataset(
        name          => Label('Maximum demand run: generation kVAr'),
        rows          => $model->{locationSet},
        data          => [ map { 0 } 1 .. $model->{numLocations} ],
        defaultFormat => '0hard',
        dataset       => $model->{dataset}
    );

    $model->{kWoffpD} = Dataset(
        name => Label(
            $model->{method} =~ /LRIC/i
            ? 'Minimum demand run: kW'
            : 'Minimum demand run: load kW'
        ),
        rows          => $model->{locationSet},
        data          => [ map { 0 } 1 .. $model->{numLocations} ],
        defaultFormat => '0hard',
        dataset       => $model->{dataset}
    );

    $model->{kVAroffpD} = Dataset(
        name => Label(
            $model->{method} =~ /LRIC/i
            ? 'Minimum demand run: kVAr'
            : 'Minimum demand run: load kVAr'
        ),
        rows          => $model->{locationSet},
        data          => [ map { 0 } 1 .. $model->{numLocations} ],
        defaultFormat => '0hard',
        dataset       => $model->{dataset}
    );

    $model->{kWoffpG} = Dataset(
        name          => Label('Minimum demand run: generation kW'),
        rows          => $model->{locationSet},
        data          => [ map { 0 } 1 .. $model->{numLocations} ],
        defaultFormat => '0hard',
        dataset       => $model->{dataset}
    );

    $model->{kVAroffpG} = Dataset(
        name          => Label('Minimum demand run: generation kVAr'),
        rows          => $model->{locationSet},
        data          => [ map { 0 } 1 .. $model->{numLocations} ],
        defaultFormat => '0hard',
        dataset       => $model->{dataset}
    );

    $model->{GSPA} = Dataset(
        name          => Label('Maximum demand run: flow through GSPs'),
        rows          => $model->{locationSet},
        data          => [ map { 0 } 1 .. $model->{numLocations} ],
        defaultFormat => '0hard',
        dataset       => $model->{dataset}
    );

    $model->{GSPB} = Dataset(
        name          => Label('Minimum demand run: flow through GSPs'),
        rows          => $model->{locationSet},
        data          => [ map { 0 } 1 .. $model->{numLocations} ],
        defaultFormat => '0hard',
        dataset       => $model->{dataset}
    );

    $model->{table911} = Columnset(
        columns => [
            $model->{locations},
            $model->{method} =~ /LRIC/i ? $model->{locDorG} : $model->{level},
            $model->{parent},
            $model->{method} =~ /LRIC/i && $model->{method} =~ /split/i
            ? (
                $model->{MCpeak}[0], $model->{MCoffp}[0],
                $model->{MCpeak}[1], $model->{MCoffp}[1],
              )
            : ( $model->{MCpeak}, $model->{MCoffp}, ),
            $model->{kWpeakD},
            $model->{kVArpeakD},
            $model->{method} =~ /LRIC/i
            ? ()
            : ( $model->{kWpeakG}, $model->{kVArpeakG} ),
            $model->{kWoffpD},
            $model->{kVAroffpD},
            $model->{method} =~ /LRIC/i
            ? ()
            : ( $model->{kWoffpG}, $model->{kVAroffpG} ),
            $model->{texit}
              && $model->{texit} =~ /2/ ? @{$model}{qw(GSPA GSPB)} : ()
        ],
        $model->{method} =~ /LRIC/i
        ? (
            name => 'LRIC power flow modelling data',
            $model->{method} =~ /split/i
            ? (
                number   => 913,
                location => 913
              )
            : (
                number   => 912,
                location => 912
            )
          )
        : (
            name     => 'FCP power flow modelling data',
            number   => 911,
            location => 911
        ),
        doNotCopyInputColumns => 1,
    );

    @{$model}{
        qw(locations level locDorG parent MCpeak MCoffp kWpeakD kVArpeakD kWpeakG kVArpeakG kWoffpD kVAroffpD kWoffpG kVAroffpG GSPA GSPB)
      };

}

sub tariffInputs {

    my ( $model, $ehvAssetLevelset, ) = @_;

    $model->{numTariffs} ||= 1;

    $model->{tariffSet} = Labelset(
        name          => 'Tariffs',
        list          => [ 1 .. $model->{numTariffs} ],
        defaultFormat => 'thtar',
    );

    my @columns = (
        Dataset(
            name          => 'LLFC',
            defaultFormat => 'texthard',
            data          => [ map { ' ' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name          => 'Tariff name',
            defaultFormat => 'texthard',
            data          => [ map { 'Not used' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name          => 'Demand or Generation',
            defaultFormat => 'texthard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name          => 'Customer class',
            defaultFormat => 'texthard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name          => 'Days for which not a customer',
            defaultFormat => '0hard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name          => 'Hours in super-red for which not a customer',
            defaultFormat => '0hard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name => 'Maximum import capacity or maximum export capacity (kVA)',
            defaultFormat => '0hard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name          => 'Capacity subject to DSM/GSM constraints (kVA)',
            defaultFormat => '0hard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name    => 'Peak-time kW divided by kVA capacity',
            data    => [ map { '' } 1 .. $model->{numTariffs} ],
            rows    => $model->{tariffSet},
            dataset => $model->{dataset}
        ),
        Dataset(
            name    => 'Peak-time kVAr divided by kVA capacity',
            data    => [ map { '' } 1 .. $model->{numTariffs} ],
            rows    => $model->{tariffSet},
            dataset => $model->{dataset}
        ),
        Dataset(
            name          => 'Units eligible for generation credits (kWh)',
            defaultFormat => '0hard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name => 'Capacity eligible for GSP generation credits (kVA)',
            defaultFormat => '0hard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name          => 'Sole use asset MEAV (£)',
            defaultFormat => '0hardnz',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name =>
              'Proportion of sole use asset MEAV not chargeable to this tariff',
            defaultFormat => '%hardnz',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name          => 'Proportion of site which is pre-2005',
            defaultFormat => '%hard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset},
            validation    => {
                validate => 'decimal',
                criteria => 'between',
                minimum  => 0,
                maximum  => 1,
            },
        ),
        Dataset(
            name          => 'LRIC location',
            defaultFormat => 'texthard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name          => 'FCP network group',
            defaultFormat => 'texthard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name    => 'Network support factor',
            data    => [ map { 1 } 1 .. $model->{numTariffs} ],
            rows    => $model->{tariffSet},
            dataset => $model->{dataset}
        ),
        Dataset(
            name          => 'Customer category for demand scaling',
            rows          => $model->{tariffSet},
            defaultFormat => '0000hard',
            data          => [ map { 1111 } 1 .. $model->{numTariffs} ],
            validation    => {
                validate => 'list',
                value    => [ split /\n/, <<EOL ],
0000
0001
0002
0010
0011
0100
0101
0110
0111
1000
1001
1100
1101
1110
1111
EOL
            },
            dataset => $model->{dataset}
        ),
        Dataset(
            name => 'Network use factor',
            data => [
                map {
                    [ map { '' } 1 .. $model->{numTariffs} ]
                  } 1 .. $#{ $ehvAssetLevelset->{list} }
            ],
            rows => $model->{tariffSet},
            cols => Labelset(
                list => [
                    @{ $ehvAssetLevelset->{list} }
                      [ 1 .. $#{ $ehvAssetLevelset->{list} } ]
                ]
            ),
            dataset    => $model->{dataset},
            validation => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
        ),
        Dataset(
            name          => 'Import tariff for generator?',
            defaultFormat => 'boolhard',
            rows          => $model->{tariffSet},
            data          => [ map { '#VALUE!' } 1 .. $model->{numTariffs} ],
            dataset       => $model->{dataset},
            defaultFormat => 'boolhard',
            validation    => {
                validate => 'list',
                value    => [ 'TRUE', 'FALSE' ],
            },
        ),
        Dataset(
            name          => 'Import tariff for LDNO?',
            defaultFormat => 'boolhard',
            rows          => $model->{tariffSet},
            data          => [ map { '#VALUE!' } 1 .. $model->{numTariffs} ],
            dataset       => $model->{dataset},
            defaultFormat => 'boolhard',
            validation    => {
                validate => 'list',
                value    => [ 'TRUE', 'FALSE' ],
            },
        ),
        Dataset(
            name          => 'Exceeded capacity (kVA)',
            defaultFormat => '0hard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name          => 'Income from previous charging year (£/year)',
            defaultFormat => '0hard',
            rows          => $model->{tariffSet},
            data          => [ map { 0 } 1 .. $model->{numTariffs} ],
            dataset       => $model->{dataset}
        ),
    );

    $model->{table953} = Columnset(
        name    => 'Tariff data',
        columns => [
            map {
                $_ ? $_ : Constant(
                    rows => $model->{tariffSet},
                    name => 'Not used',
                    data => [ map { '' } 1 .. $model->{numTariffs} ],
                  )
              } @columns
        ],
        number                => 953,
        location              => 953,
        doNotCopyInputColumns => 1,
    );

    @columns;

}

1;
