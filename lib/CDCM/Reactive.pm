package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others.
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

sub reactive {
    my (
        $model,                  $drmExitLevels,
        $chargingDrmExitLevels,  $chargingLevels,
        $componentMap,           $allTariffsByEndUser,
        $unitYardstick,          $costToSml,
        $loadCoefficients,       $lineLossFactorsToGsp,
        $lineLossFactorsNetwork, $proportionCoveredByContributions,
        $daysInYear,             $powerFactorInModel,
        $tariffsExMatching,      $componentLabelset,
        $sourceMap
    ) = @_;

    my $banded = $model->{reactive} && $model->{reactive} =~ /band/i;

    my $step7 = $model->{reactive} && $model->{reactive} =~ /step7/i;

    my $idno = $model->{reactive} && $model->{reactive} =~ /idno/i;

    my $sensibleReactiveMethod = !$step7 && !$banded;

    push @{ $model->{optionLines} }, 'Reactive power unit charges: step 7'
      if $step7;

    my $nineteen = $sensibleReactiveMethod ? 20 : 19;

    my @reactiveBands =
      map {
        join '', 'Power factor ', .05 * $_ - .05,
          ( $_ == 1 ? '.00' : $_ % 2 ? '0' : () ), '-', .05 * $_,
          ( $_ == 20 ? '.00' : $_ % 2 ? () : '0' )
      } reverse 1 .. $nineteen;

    my $reactiveBandset = Labelset(
        name => 'Power factor bands',
        list => \@reactiveBands
    );

    my $powerFactorMidBand = Constant(
        cols => $reactiveBandset,
        name => 'Power factor in the middle of the power factor band',
        data => [ map { 0.05 * $_ - 0.025 } reverse 1 .. $nineteen ]
    );

    my @relevantGroups =
      grep { $componentMap->{$_}{'Reactive power charge p/kVArh'} }
      @{ $allTariffsByEndUser->{groups} || $allTariffsByEndUser->{list} };

    @relevantGroups =
      map { Labelset( name => $_->{name}, list => [ $_->{list}[0] ] ) }
      @relevantGroups
      unless $idno;

    my $tariffsetForReactiveByEndUser = Labelset(
        name   => 'Tariffs with reactive power unit charges',
        groups => \@relevantGroups
    );

    my $tariffsetForReactiveByEndUserStandard = Labelset(
        name   => 'Tariffs with standard reactive power unit charges',
        groups => [
            grep {
                $componentMap->{$_}{'Reactive power charge p/kVArh'} eq
                  'Standard kVArh'
            } @relevantGroups
        ]
    );

    my $tariffsetForReactiveByEndUserPayg = Labelset(
        name   => 'Tariffs with pay-as-you-go ' . 'reactive power unit charges',
        groups => [
            grep {
                $componentMap->{$_}{'Reactive power charge p/kVArh'} eq
                  'PAYG kVArh'
            } @relevantGroups
        ]
    );

    my $tariffsetForReactive = Labelset
      name => 'Tariffs with reactive power unit charges',
      list => [
        grep  { $componentMap->{$_}{'Reactive power charge p/kVArh'}; }
          map { $allTariffsByEndUser->{list}[$_] }
          $allTariffsByEndUser->indices
      ];

    push @{ $model->{reactiveResults} },
      my $routeingFactorsReactiveUnits = Constant(
        rows  => $tariffsetForReactiveByEndUserPayg,
        cols  => $drmExitLevels,
        byrow => 1,
        data  => [
            map {
                /^((?:LD|Q)NO )?LV sub/i
                  ? [ 1, 1, 1, 1, $model->{extraLevels} ? 1 : (), 1, 1, 1, 0 ]
                  : /^((?:LD|Q)NO )?LV/i ? [ map { 1 } 0 .. 8 ]
                  : /^((?:LD|Q)NO )?HV sub/i
                  ? [ 1, 1, 1, 1, $model->{extraLevels} ? 0 : (), 1, 0, 0, 0 ]
                  : /^((?:LD|Q)NO )?HV/i
                  ? [ 1, 1, 1, 1, $model->{extraLevels} ? 1 : (), 1, 1, 0, 0 ]
                  : /^((?:LD|Q)NO )?33kV sub/i
                  ? [ 1, 1, 1, 0, $model->{extraLevels} ? 0 : (), 0, 0, 0, 0 ]
                  : /^((?:LD|Q)NO )?33/i
                  ? [ 1, 1, 1, 1, $model->{extraLevels} ? 0 : (), 0, 0, 0, 0 ]
                  : /^GSP/i
                  ? [ 1, 0, 0, 0, $model->{extraLevels} ? 0 : (), 0, 0, 0, 0 ]
                  : /^((?:LD|Q)NO )?132/i
                  ? [ 1, 1, 0, 0, $model->{extraLevels} ? 0 : (), 0, 0, 0, 0 ]
                  : [ 0, 0, 0, 0, $model->{extraLevels} ? 0 : (), 0, 0, 0, 0 ]
            } @{ $tariffsetForReactiveByEndUserPayg->{list} }
        ],
        defaultFormat => '0connz',
        name          => Label(
            'Network use factors for generator reactive unit charges',
            'Network use factor (reactive)'
        ),
        lines => <<'EOT'
These factors differ from the network use factors for active power charges/credits in the case of generators, who do not qualify
for active power credits at the voltage of connection but are charged reactive unit charges for costs caused at that voltage.
EOT
      );

    my $paygUnitForReactive = Arithmetic(
        name =>
          'Pay-as-you-go components p/kWh for reactive power (absolute value)',
        defaultFormat => '0.000softnz',
        rows          => $tariffsetForReactiveByEndUserPayg,
        cols          => $chargingDrmExitLevels,
        arithmetic    => '=A1*A2*A90/A91*(1-A92)*A93/(24*A5)*100',
        arguments     => {
            A1 => $costToSml,
            A2 => Arithmetic(
                name => Label(
                    'Absolute load coefficient',
                    'Absolute value of load coefficient (kW peak / average kW)'
                ),
                rows => Labelset(
                    name =>
                      'End users with pay-as-you-go reactive power charges',
                    list => $tariffsetForReactiveByEndUserPayg->{groups}
                      || $tariffsetForReactiveByEndUserPayg->{list}
                ),
                arithmetic => '=ABS(A1)',
                arguments  => { A1 => $loadCoefficients }
            ),
            A90 => $lineLossFactorsToGsp,
            A91 => $lineLossFactorsNetwork,
            A92 => $proportionCoveredByContributions,
            A93 => $routeingFactorsReactiveUnits,
            A5  => $daysInYear
        }
    );

    my $standardUnitForReactive = Arithmetic(
        name => 'Standard components '
          . 'p/kWh for reactive power (absolute value)',
        rows       => $tariffsetForReactiveByEndUserStandard,
        cols       => $chargingDrmExitLevels,
        arithmetic => '=ABS(A1)',
        arguments  => { A1 => $unitYardstick->{source} }
    );

    my ( $paygReactive, $standardReactive );

    if ($sensibleReactiveMethod) {

        my $averageKvarByKva = Dataset(
            name       => 'Average kVAr by kVA, by network level',
            validation => {
                validate => 'decimal',
                criteria => 'between',
                minimum  => 0,
                maximum  => 1,
            },
            number   => 1092,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            cols     => $drmExitLevels,
            lines    => [
                'Source: analysis of operational data.',
                'This is the average of MVAr/MVA or SQRT(1-PF^2)'
                  . ' across relevant network elements.',
            ],
            data => [ map { 0.3 } @{ $drmExitLevels->{list} } ]
        );

        $model->{sharedData}
          ->addStats( 'Input data (15 months notice of network model etc)',
            $model, $averageKvarByKva )
          if $model->{sharedData};

        $paygReactive = Arithmetic(
            name       => 'Pay-as-you-go reactive p/kVArh',
            cols       => $chargingDrmExitLevels,
            arithmetic => '=A1*A2*A3',
            arguments  => {
                A1 => $paygUnitForReactive,
                A2 => $averageKvarByKva,
                A3 => $powerFactorInModel
            }
        );

        $standardReactive = Arithmetic(
            name       => 'Standard reactive p/kVArh',
            cols       => $chargingDrmExitLevels,
            arithmetic => '=A1*A2*A3',
            arguments  => {
                A1 => $standardUnitForReactive,
                A2 => $averageKvarByKva,
                A3 => $powerFactorInModel
            }
        );

        $sourceMap->{'Reactive power charge p/kVArh'}{'Standard kVArh'} =
          [$standardReactive];

        $sourceMap->{'Reactive power charge p/kVArh'}{'PAYG kVArh'} =
          [$paygReactive];

        if ( $model->{reactive} && $model->{reactive} =~ /notScaled/i ) {

            my $components = Stack(
                name    => "Reactive power charge p/kVArh by network level",
                rows    => $tariffsetForReactive,
                cols    => $chargingDrmExitLevels,
                sources => [ $paygReactive, $standardReactive ]
            );

            my $aggregate = GroupBy(
                name   => 'Reactive power charge p/kVArh',
                rows   => $tariffsetForReactive,
                source => $components
            );

            push @{ $model->{reactiveResults} }, $standardReactive,
              $paygReactive,
              Columnset(
                name    => 'Reactive power charges',
                columns => [ $components, $aggregate ]
              );

            $tariffsExMatching->{'Reactive power charge p/kVArh'} = Stack(
                rows    => $allTariffsByEndUser,
                sources => [$aggregate]
            );

        }

        else {

            my $components = Stack(
                name    => "Reactive power charge p/kVArh (elements)",
                rows    => $allTariffsByEndUser,
                cols    => $chargingLevels,
                sources => [ $paygReactive, $standardReactive ]
            );

            my $aggregate = GroupBy(
                name   => 'Reactive power charge p/kVArh',
                rows   => $allTariffsByEndUser,
                source => $components
            );

            push @{ $model->{reactiveResults} }, $standardReactive,
              $paygReactive;

            $tariffsExMatching->{'Reactive power charge p/kVArh'} = $aggregate;

        }

    }

    else {

        $paygReactive = Arithmetic(
            name       => 'Pay-as-you-go reactive yardstick',
            cols       => $reactiveBandset,
            arithmetic => '=A1*SQRT(1-A2^2)*A3',
            arguments  => {
                A1 => GroupBy(
                    name   => 'p/kVAh',
                    rows   => $tariffsetForReactiveByEndUserPayg,
                    source => $paygUnitForReactive
                ),
                A2 => $powerFactorMidBand,
                A3 => $powerFactorInModel
            }
        );

        $standardReactive = Arithmetic(
            name       => 'Standard reactive yardstick',
            cols       => $reactiveBandset,
            arithmetic => '=A1*SQRT(1-A2^2)*A3',
            arguments  => {
                A1 => GroupBy(
                    name   => 'p/kVAh',
                    rows   => $tariffsetForReactiveByEndUserStandard,
                    source => $standardUnitForReactive
                ),
                A2 => $powerFactorMidBand,
                A3 => $powerFactorInModel
            }
        );

        if ($step7) {

            my $weights = Dataset(
                name =>
                  'Percentage of reactive units in each power factor band',
                validation => {
                    validate => 'decimal',
                    criteria => 'between',
                    minimum  => 0,
                    maximum  => 1,
                },
                cols          => $reactiveBandset,
                defaultFormat => '%hard',
                data => [ map { 1.0 / $nineteen } reverse 1 .. $nineteen ]
            );

            $standardReactive = GroupBy(
                name   => 'Standard reactive unit charge after step 7',
                rows   => $standardReactive->{rows},
                cols   => 0,
                source => Arithmetic(
                    name =>
'Contributions to standard reactive unit charges under step 7',
                    arithmetic => '=A1*A2',
                    arguments  => { A1 => $standardReactive, A2 => $weights }
                )
            );

            $paygReactive = GroupBy(
                name   => "Pay-as-you-go reactive unit charge after step 7",
                rows   => $paygReactive->{rows},
                cols   => 0,
                source => Arithmetic(
                    name => 'Contributions to pay-as-you-go'
                      . ' reactive unit charges under step 7',
                    arithmetic => '=A1*A2',
                    arguments  => { A1 => $paygReactive, A2 => $weights }
                )
            );

        }
        else {
            $componentLabelset->{'Reactive power charge p/kVArh'} =
              $reactiveBandset;
        }

        push @{ $model->{reactiveResults} }, $standardReactive, $paygReactive;

        $tariffsExMatching->{'Reactive power charge p/kVArh'} = Stack(
            name    => 'Reactive power charge p/kVArh',
            rows    => $allTariffsByEndUser,
            cols    => $componentLabelset->{'Reactive power charge p/kVArh'},
            sources => [ $standardReactive, $paygReactive, ]
        );

    }

}

1;
