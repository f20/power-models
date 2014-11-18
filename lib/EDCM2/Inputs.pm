package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2012 Energy Networks Association Limited and others.
Copyright 2013-2014 Franck Latrémolière, Reckon LLP and others.

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

=head Table numbers used in this file

911
912
913
935
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
            criteria => 'between',
            minimum  => 365,
            maximum  => 366,
        },
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

    my $allowedRev = Dataset(
        name => 'The amount of money that the DNO wants to raise'
          . ' from use of system charges, less transmission exit (£/year)',
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
        name       => 'EHV operating expenditure intensity',
        data       => [0.68],
        validation => {
            validate => 'decimal',
            criteria => 'between',
            minimum  => 0,
            maximum  => 1,
        }
    );

    my $powerFactorInModel = Constant(
        name       => 'Power factor in 500 MW model',
        validation => {
            validate => 'decimal',
            criteria => 'between',
            minimum  => 0.001,
            maximum  => 1,
        },
        data => [0.95]
    );

    my $genPot20p = Dataset(
        name => 'O&M charging rate based on FBPQ data (£/kW/year)',
        data => [0],
    );

    my $genPotGP = Dataset(
        name          => 'Average adjusted GP (£/year)',
        defaultFormat => '0hard',
        data          => [0],
    );

    my $genPotGL = Dataset(
        name => 'GL term from the DG incentive revenue calculation (£/year)',
        defaultFormat => '0hard',
        data          => [0],
    );

    my $genPotCdcmCap20052010 = Dataset(
        name          => 'Total CDCM generation capacity 2005-2010 (kVA)',
        defaultFormat => '0hard',
        data          => [0],
    );

    my $genPotCdcmCapPost2010 = Dataset(
        name          => 'Total CDCM generation capacity post-2010 (kVA)',
        defaultFormat => '0hard',
        data          => [0],
    );

    my $hoursInRed = Dataset(
        name          => 'Annual hours in super-red',
        defaultFormat => '0.0hardnz',
        data          => [300],
    );

    Columnset(
        name    => 'General inputs',
        columns => [
            $days,                  $genPot20p,
            $hoursInRed,            $allowedRev,
            $tExit,                 $direct,
            $indirect,              $rates,
            $genPotGP,              $genPotGL,
            $genPotCdcmCap20052010, $genPotCdcmCapPost2010,
        ],
        number   => 1113,
        dataset  => $model->{dataset},
        appendTo => $model->{inputTables}
    );

    $days, $direct, $indirect, $rates, $tExit, $ehvIntensity, $allowedRev,
      $powerFactorInModel, $genPot20p, $genPotGP, $genPotGL,
      $genPotCdcmCap20052010, $genPotCdcmCapPost2010,
      $hoursInRed;

}

sub loadFlowInputs {

    my ($model) = @_;

    return if $model->{method} eq 'none';
    $model->{numLocations} ||= 16;
    $model->{locationSet}  ||= Labelset(
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
    ) if $model->{method} =~ /LRIC/i;

    $model->{level} = Dataset(
        name          => Label('Level'),
        rows          => $model->{locationSet},
        data          => [ map { 3 } 1 .. $model->{numLocations} ],
        defaultFormat => '0hardnz',
        dataset       => $model->{dataset}
    ) unless $model->{method} =~ /LRIC/i;

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
    ) unless $model->{method} =~ /LRIC/i;

    $model->{kVArpeakG} = Dataset(
        name          => Label('Maximum demand run: generation kVAr'),
        rows          => $model->{locationSet},
        data          => [ map { 0 } 1 .. $model->{numLocations} ],
        defaultFormat => '0hard',
        dataset       => $model->{dataset}
    ) unless $model->{method} =~ /LRIC/i;

    $model->{table911} = Columnset(
        columns => [
            $model->{locations},
            $model->{method} =~ /LRIC/i ? $model->{locDorG} : $model->{level},
            $model->{parent},
            $model->{method} =~ /LRIC/i
            ? (
                $model->{MCpeak}[0],
                $model->{MCpeak}[1],
                Constant(
                    name          => 'Not used',
                    defaultFormat => 'unused',
                    rows          => $model->{locationSet},
                    data          => [ map { 0 } 1 .. $model->{numLocations} ],
                ),
                Constant(
                    name          => 'Not used',
                    defaultFormat => 'unused',
                    rows          => $model->{locationSet},
                    data          => [ map { 0 } 1 .. $model->{numLocations} ],
                ),
              )
            : (
                $model->{MCpeak},
                Constant(
                    name          => 'Not used',
                    defaultFormat => 'unused',
                    rows          => $model->{locationSet},
                    data          => [ map { 0 } 1 .. $model->{numLocations} ],
                ),
            ),
            $model->{kWpeakD},
            $model->{kVArpeakD},
            $model->{method} =~ /LRIC/i
            ? (
                Constant(
                    name          => 'Not used',
                    defaultFormat => 'unused',
                    rows          => $model->{locationSet},
                    data          => [ map { 0 } 1 .. $model->{numLocations} ],
                ),
                Constant(
                    name          => 'Not used',
                    defaultFormat => 'unused',
                    rows          => $model->{locationSet},
                    data          => [ map { 0 } 1 .. $model->{numLocations} ],
                ),
              )
            : (
                $model->{kWpeakG},
                $model->{kVArpeakG},
                Constant(
                    name          => 'Not used',
                    defaultFormat => 'unused',
                    rows          => $model->{locationSet},
                    data          => [ map { 0 } 1 .. $model->{numLocations} ],
                ),
                Constant(
                    name          => 'Not used',
                    defaultFormat => 'unused',
                    rows          => $model->{locationSet},
                    data          => [ map { 0 } 1 .. $model->{numLocations} ],
                ),
                Constant(
                    name          => 'Not used',
                    defaultFormat => 'unused',
                    rows          => $model->{locationSet},
                    data          => [ map { 0 } 1 .. $model->{numLocations} ],
                ),
                Constant(
                    name          => 'Not used',
                    defaultFormat => 'unused',
                    rows          => $model->{locationSet},
                    data          => [ map { 0 } 1 .. $model->{numLocations} ],
                ),
            ),

        ],
        $model->{method} =~ /LRIC/i
        ? (
            name     => 'LRIC power flow modelling data',
            number   => 913,
            location => 913
          )
        : (
            name     => 'FCP power flow modelling data',
            number   => 911,
            location => 911
        ),
        doNotCopyInputColumns => 1,
    );

    @{$model}{qw(locations parent MCpeak kWpeakD kVArpeakD kWpeakG kVArpeakG)};

}

sub tariffInputs {

    my ( $model, $ehvAssetLevelset, ) = @_;

    $model->{numTariffs} ||= $model->{transparency}
      && $model->{transparency} =~ /impact/i ? 0 : 16;
    $model->{tariffSet} ||=
      $model->{useTariffNicknames}
      ? Labelset(
        name     => 'Tariffs',
        editable => Dataset(
            rows => Labelset(
                list =>
                  [ map { "Nickname of tariff $_" } 1 .. $model->{numTariffs} ]
            ),
            name          => 'Tariff nicknames',
            defaultFormat => 'texthard',
            appendTo      => $model->{inputTables},
            number        => 1101,
            data          => [ 1 .. $model->{numTariffs} ],
            lines         => 'These nicknames do not affect any calculations.',
        )
      )
      : Labelset(
        name          => 'Tariffs',
        list          => [ 1 .. $model->{numTariffs} ],
        defaultFormat => 'thtar',
      );

    my @columns = (
        Dataset(
            name          => 'Name',
            defaultFormat => 'texthard',
            data          => [ map { 'Not used' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name          => 'Maximum import capacity (kVA)',
            defaultFormat => '0hard',
            data          => [ map { 'VOID' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name          => 'Exempt export capacity (kVA)',
            defaultFormat => '0hardnz',
            data          => [ map { 'VOID' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name          => 'Non-exempt pre-2005 export capacity (kVA)',
            defaultFormat => '0hardnz',
            data          => [ map { 'VOID' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset},
        ),
        Dataset(
            name          => 'Non-exempt 2005-2010 export capacity (kVA)',
            defaultFormat => '0hardnz',
            data          => [ map { 'VOID' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset},
        ),
        Dataset(
            name          => 'Non-exempt post-2010 export capacity (kVA)',
            defaultFormat => '0hardnz',
            data          => [ map { 'VOID' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset},
        ),
        Dataset(
            name          => 'Sole use asset MEAV (£)',
            defaultFormat => '0hard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        !$model->{dcp189} ? () : $model->{dcp189} =~ /proportion/i ? Dataset(
            name => 'Percentage of sole use assets where '
              . 'Customer is entitled to reduction for capitalised O&M',
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
          ) : Dataset(
            name => 'Customer entitled to reduction for capitalised O&M',
            defaultFormat => '0hard',
            data          => [ map { 'N' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset},
            validation    => {
                validate => 'list',
                value    => [qw(N Y)],
            },
          ),
        $model->{method} =~ /LRIC/i ? Dataset(
            name          => 'LRIC location',
            defaultFormat => 'texthard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
          )
        : $model->{method} =~ /FCP/i ? Dataset(
            name          => 'FCP network group',
            defaultFormat => 'texthard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
          )
        : undef,
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
            name    => 'Super-red kW import divided by kVA capacity',
            data    => [ map { '' } 1 .. $model->{numTariffs} ],
            rows    => $model->{tariffSet},
            dataset => $model->{dataset}
        ),
        Dataset(
            name    => 'Super-red kVAr import divided by kVA capacity',
            data    => [ map { '' } 1 .. $model->{numTariffs} ],
            rows    => $model->{tariffSet},
            dataset => $model->{dataset}
        ),
        Dataset(
            name => 'Proportion exposed to indirect cost allocation'
              . ( $model->{dcp185} ? ' and fixed adder' : '' ),
            data       => [ map { 1 } 1 .. $model->{numTariffs} ],
            rows       => $model->{tariffSet},
            dataset    => $model->{dataset},
            validation => {
                validate => 'decimal',
                criteria => 'between',
                minimum  => 0,
                maximum  => 1,
            },
        ),
        Dataset(
            name          => 'Capacity subject to DSM (kVA)',
            defaultFormat => '0hard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),

        Dataset(
            name          => 'Super-red units exported (kWh)',
            defaultFormat => '0hard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name => 'Capacity eligible for GSP generation credits (kW)',
            defaultFormat => '0hard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name    => 'Proportion eligible for charge 1 credits',
            data    => [ map { 1 } 1 .. $model->{numTariffs} ],
            rows    => $model->{tariffSet},
            dataset => $model->{dataset}
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
            defaultFormat => '0.0hard',
            data          => [ map { '' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),

        Dataset(
            name          => 'Import charge in previous charging year (£/year)',
            defaultFormat => '0hard',
            rows          => $model->{tariffSet},
            data          => [ map { 0 } 1 .. $model->{numTariffs} ],
            dataset       => $model->{dataset}
        ),
        Dataset(
            name          => 'Export charge in previous charging year (£/year)',
            defaultFormat => '0hard',
            rows          => $model->{tariffSet},
            data          => [ map { 0 } 1 .. $model->{numTariffs} ],
            dataset       => $model->{dataset}
        ),

        Dataset(
            name          => 'LLFC import',
            defaultFormat => 'texthard',
            data          => [ map { ' ' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),
        Dataset(
            name          => 'LLFC export',
            defaultFormat => 'texthard',
            data          => [ map { ' ' } 1 .. $model->{numTariffs} ],
            rows          => $model->{tariffSet},
            dataset       => $model->{dataset}
        ),

    );

    $model->{table935} = Columnset(
        name    => 'Tariff data',
        columns => [
            map {
                $_ ? $_ : Constant(
                    rows          => $model->{tariffSet},
                    name          => 'Not used',
                    defaultFormat => 'unused',
                    data          => [ map { '' } 1 .. $model->{numTariffs} ],
                  )
            } @columns
        ],
        number                => 935,
        location              => 935,
        doNotCopyInputColumns => 1,
    );

    @columns;

}

1;
