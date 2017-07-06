package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others.
Copyright 2016-2017 Franck Latrémolière, Reckon LLP and others.

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

sub customerContributions {
    my ( $model, $assetDrmLevels, $drmExitLevels, $operatingDrmExitLevels,
        $chargingDrmExitLevels, $allTariffsByEndUser, )
      = @_;

    my $customerTypesForContributions = Labelset(
        name => 'Network level of supply (for connection contributions)',
        list => [
            ( split /\n/, <<'EOL' ),
LV network
LV substation
HV network
HV substation
EOL
            $model->{ehv}
              || $model->{portfolio} && $model->{portfolio} =~ /ehv/i
            ? ( '33kV network', '33kV substation', '132kV network', 'GSP' )
            : (),
        ]
    );

    my $customerContributions = Dataset(
        number   => 1060,
        appendTo => $model->{inputTables},
        dataset  => $model->{dataset},
        name     => 'Customer contributions'
          . ' under current connection charging policy',
        validation => {
            validate      => 'decimal',
            criteria      => '>=',
            value         => 0,
            input_title   => 'Customer contribution:',
            input_message => 'Percentage',
            error_title   => 'Invalid customer contribution',
            error_message => 'Invalid customer contribution'
              . ' (negative number or unused cell).'
        },
        lines => [
'Source: analysis of expenditure data and/or survey of capital expenditure schemes.',
'Customer contribution percentages by network level of supply and by asset network level.',
'These proportions should reflect the current connection charging method, '
              . 'not necessarily the method that was in place when the connection was built.'
        ],
        cols          => $assetDrmLevels,
        rows          => $customerTypesForContributions,
        defaultFormat => '%hard',
        byrow         => 1,
        data          => $model->{extraLevels}
        ? [ map { [ split /\s+/ ] } split /\n/, <<'EOT' ]
0   0   0   0.1  .1  0.35   0.99   0.99
0   0   0   0.1  .1  0.35   0.99
0   0   0.05   0.2 .2  0.7
0   0   0.05   0.4
0   0.05   0.2
0.05   0.1
0.1
EOT
        : [ map { [ split /\s+/ ] } split /\n/, <<'EOT' ]
0   0   0   0.1   0.35   0.99   0.99
0   0   0   0.1   0.35   0.99
0   0   0.05   0.2   0.7
0   0   0.05   0.4
0   0.05   0.2
0.05   0.1
0.1
EOT
    );

    # second line if using large/small: 0	0	0	0	0.3	1	1

    my $proportionChargeable =
      $model->{noReplacement} && $model->{noReplacement} =~ /blanket/i
      ? Constant(
        name => 'Annuity proportion for customer-contributed assets',
        data => [ [0] ]
      )
      : Dataset(
        name       => 'Annuity proportion for customer-contributed assets',
        lines      => 'Source: financial assumption.',
        validation => {
            validate      => 'decimal',
            criteria      => 'between',
            minimum       => 0,
            maximum       => 1,
            input_title   => 'Proportion of annuity:',
            input_message => 'Between 0% and 100%',
            error_message => 'The proportion chargeable'
              . ' must be between 0% and 100%.'
        },
        defaultFormat => '%hardnz',
        data          => [.453]
      );

    push @{ $model->{summaryColumns} },
      Stack( sources => [$proportionChargeable] );

=Old

    lines   => 'By type of connection and network level.',
    rows    => $assetDrmLevels,
    cols    => $customerTypesForContributions,
    data    => [
            map { [ split /\s+/ ] } split /\n/, <<'EOT'
0   0	0	0	0.4 0.4	0.4
0	0	0	0	0.4	0.4	0.4
0   0   0   0   0.4 0.4
0	0	0.4	0.4 0.4
0   0   0.4 0.4
0   0.4 0.4
0.4
EOT
        ]

=cut

    my $customerTypeMatrixForContributions = Constant(
        rows => $allTariffsByEndUser,
        cols => $customerTypesForContributions,

        # hard-coded customer types

        # LV network
        # LV substation
        # HV network
        # HV substation
        # 33kV network
        # 33kV substation
        # 132kV network
        # GSP

        byrow => 1,
        data  => [
            map {
                    /^((?:LD|Q)NO )?LV sub/i   ? [qw(0 1 0 0 0 0 0 0)]
                  : /^((?:LD|Q)NO )?LV/i       ? [qw(1 0 0 0 0 0 0 0)]
                  : /^((?:LD|Q)NO )?HV sub/i   ? [qw(0 0 0 1 0 0 0 0)]
                  : /^((?:LD|Q)NO )?HV/i       ? [qw(0 0 1 0 0 0 0 0)]
                  : /^((?:LD|Q)NO )?33kV sub/i ? [qw(0 0 0 0 0 1 0 0)]
                  : /^((?:LD|Q)NO )?33/i       ? [qw(0 0 0 0 1 0 0 0)]
                  : /^((?:LD|Q)NO )?132/i      ? [qw(0 0 0 0 0 0 1 0)]
                  : /^GSP/i                    ? [qw(0 0 0 0 0 0 0 1)]
                  : []
            } @{ $allTariffsByEndUser->{list} }
        ],
        defaultFormat => '%connz',
        name =>
          'Network level of supply (for customer contributions) by tariff',
        1
        ? ()
        : ( lines =>
'Fractional figures indicate a mixture of user types (for customer contributions) within a tariff.'
        )
    );

    my $lvTariffset = Labelset(
        name => 'LV tariffs',
        list => [
            grep  { /^((?:LD|Q)NO )?LV/i }
              map { $allTariffsByEndUser->{list}[$_] }
              $allTariffsByEndUser->indices
        ]
    );

    0 and $customerTypeMatrixForContributions = Stack
      name => 'Mapping of network level of supply '
      . '(for customer contributions) to each LV tariff',
      rows          => $allTariffsByEndUser,
      cols          => $customerTypesForContributions,
      defaultFormat => '%copynz',
      sources       => [
        Dataset(
            defaultFormat => '%hardnz',
            rows          => $lvTariffset,
            cols          => Labelset(
                name => 'LV types for contributions',
                list => [ @{ $customerTypesForContributions->{list} }[ 0, 1 ] ]
            ),
            byrow => 1,
            data  => [ map { [ 0.5, 0.5 ] } @{ $lvTariffset->{list} } ],
            name  => 'Mapping of tariffs to network level of'
              . ' supply (for customer contributions)',
            validation => {
                validate      => 'decimal',
                criteria      => 'between',
                minimum       => 0,
                maximum       => 1,
                input_title   => 'Proportion:',
                input_message => 'Between 0% and 100%',
                error_message => 'This data point'
                  . ' must be between 0% and 100%.'
            },
        ),
        $customerTypeMatrixForContributions
      ];

    0 and my $proportionAssetsCoveredByContributions = SumProduct(
        name => Label(
            'Contribution proportion',
            'Contribution proportion by tariff and network level'
              . ' (proportion of assets deemed covered '
              . 'by customer contributions when first built)'
        ),
        matrix        => $customerTypeMatrixForContributions,
        vector        => $customerContributions,
        defaultFormat => '%softnz'
    );

    my $proportionCoveredByContributions = SumProduct(
        name => Label(
            'Contribution proportion',
            'Proportion of asset annuities'
              . ' deemed to be covered by customer contributions'
        ),
        matrix => $customerTypeMatrixForContributions,
        vector => Arithmetic(
            rows => $assetDrmLevels,
            cols => $customerTypesForContributions,
            name =>
              'Contribution proportion of asset annuities, by customer type'
              . ' and network level of assets',
            arithmetic => '=A1*(1-A2)',
            arguments =>
              { A1 => $customerContributions, A2 => $proportionChargeable },
            defaultFormat => '%softnz'
        ),
        defaultFormat => '%softnz'
    );

    push @{ $model->{contributions} }, $customerContributions,
      $proportionChargeable, $proportionCoveredByContributions;

    push @{ $model->{contributions} },
      my $allLevelsProportionCoveredByContributions = Stack(
        name => 'Proportion of annual charge covered by contributions'
          . ' (for all charging levels)',
        defaultFormat => '%copynz',
        rows          => $allTariffsByEndUser,
        cols          => $chargingDrmExitLevels,
        sources       => [
            Constant(
                name          => 'Zero for operating expenditure',
                defaultFormat => '%connz',
                rows          => $allTariffsByEndUser,
                cols          => $operatingDrmExitLevels,
                data          => [
                    map {
                        [ map { 0 } @{ $allTariffsByEndUser->{list} } ]
                    } @{ $operatingDrmExitLevels->{list} }
                ],
            ),
            Constant(
                name          => 'Zero for GSPs level',
                defaultFormat => '%connz',
                rows          => $allTariffsByEndUser,
                cols =>
                  Labelset( list => [ $chargingDrmExitLevels->{list}[0] ] ),
                data => [ [ map { 0 } @{ $allTariffsByEndUser->{list} } ] ],
            ),
            $model->{generationCreditsContrib}
            ? (
                $model->{generationCreditsContrib} =~ /100/
                ? Arithmetic(
                    name          => 'Full contribution for generation',
                    defaultFormat => '%soft',
                    rows          => Labelset(
                        name => 'Generation tariffs',
                        list => [
                            grep { /gener/i; } @{ $allTariffsByEndUser->{list} }
                        ]
                    ),
                    cols       => $chargingDrmExitLevels,
                    arithmetic => '=1-A2',
                    arguments  => { A2 => $proportionChargeable, },
                  )
                : Constant(
                    name          => 'Zero for generation',
                    defaultFormat => '%con',
                    rows          => Labelset(
                        name => 'Generation tariffs',
                        list => [
                            grep { /gener/i; } @{ $allTariffsByEndUser->{list} }
                        ]
                    ),
                    cols => $chargingDrmExitLevels,
                    data => [
                        map {
                            [ map { 0 } @{ $allTariffsByEndUser->{list} } ]
                          } grep { /gener/i; }
                          @{ $chargingDrmExitLevels->{list} }
                    ],
                )
              )
            : (),
            $proportionCoveredByContributions
        ]
      );

    my $replacementShare;

    $replacementShare = Arithmetic(
        name => 'Share of network model annuity that relates'
          . ' to replacement of customer contributed assets',
        defaultFormat => '%softnz',
        arithmetic    => '=A1*A2/(1-A3)/(1-A4)',
        arguments     => {
            A1 => $proportionCoveredByContributions,
            A2 => $proportionChargeable,
            A3 => $proportionCoveredByContributions,
            A4 => $proportionChargeable,
        }
    ) if $model->{noReplacement} && $model->{noReplacement} =~ /hybrid/i;

    $proportionCoveredByContributions,            $proportionChargeable,
      $allLevelsProportionCoveredByContributions, $replacementShare;

}

1;
