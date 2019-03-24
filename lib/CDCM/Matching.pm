package CDCM;

# Copyright 2009-2011 Energy Networks Association Limited and others.
# Copyright 2011-2017 Franck Latrémolière, Reckon LLP and others.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice,
# this list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
# this list of conditions and the following disclaimer in the documentation
# and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
# EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
# THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Shortcuts ':all';
use SpreadsheetModel::SegmentRoot;

sub matching {

    my (
        $model,                        $revenueShortfall,
        $revenuesSoFar,                $totalRevenuesSoFar,
        $totalSiteSpecificReplacement, $componentMap,
        $allTariffsByEndUser,          $demandTariffsByEndUser,
        $allEndUsers,                  $chargingLevels,
        $nonExcludedComponents,        $allComponents,
        $daysInYear,                   $volumeData,
        $revenueBefore,                $loadCoefficients,
        $lineLossFactorsToGsp,         $tariffsExMatching,
        $unitsInYear,                  $generationCapacityTariffsByEndUser,
        $fFactors,                     $annuityRate,
        $modelLife,                    $costToSml,
        $modelCostToSml,               $operatingCostToSml,
        $routeingFactors,              $replacementShare,
        $siteSpecificOperatingCost,    $siteSpecificReplacement,
        $simultaneousMaximumLoadUnits, $simultaneousMaximumLoadCapacity,
        @unitRateSystemLoadCoefficients,
    ) = @_;

    my $doNotScaleReactive =
      $model->{reactive} && $model->{reactive} =~ /notScaled/i;

    push @{ $model->{optionLines} },
      'Reactive power unit charges not subject to scaler'
      if $doNotScaleReactive;

    my $scaledComponents =
      $doNotScaleReactive ? $nonExcludedComponents : $allComponents;

    my (
        $scalerTable,    $totalRevenuesFromScaler,
        $assetScalerPot, $siteSpecificCharges
    );

    if ( $model->{scaler} && $model->{scaler} =~ /adder/i ) {
        if ($totalSiteSpecificReplacement) {
            $siteSpecificCharges = Arithmetic(
                name => 'Total site specific sole use asset charges (£/year)',
                defaultFormat => '0softnz',
                cols          => Labelset(
                    name => 'Site specific levels',
                    list => [
                        map { "Site-specific $_" }
                          @{ $siteSpecificReplacement->{cols}{list} }
                    ],
                    accepts => [
                        $siteSpecificReplacement->{cols},
                        $siteSpecificOperatingCost->{cols}
                    ]
                ),
                arithmetic => '=A2+A1',
                arguments  => {
                    A1 => $siteSpecificOperatingCost,
                    A2 => $siteSpecificReplacement,
                }
            );
        }
    }

    else {    # cost scaler

        my $dontScaleGeneration =
          $model->{scaler} && $model->{scaler} =~ /nogen/i;
        my ( $scalerWeightsConsumption, $scalerWeightsStanding );
        if ( $model->{scaler} && $model->{scaler} =~ /gen([0-9.]+)/ ) {
            my $generationWeight   = $1;
            my @scalerWeightsItems = (
                rows => $allTariffsByEndUser,
                data => [
                    map { /gener/i ? $generationWeight : 1; }
                      @{ $allTariffsByEndUser->{list} }
                ],
                validation => {
                    validate      => 'decimal',
                    criteria      => '>=',
                    value         => 0,
                    input_title   => 'Tariff weighting:',
                    input_message => 'Non-negative number',
                    error_message =>
                      'This weighting cannot sensibly be negative.'
                },
            );
            if ( $model->{scaler} =~ /editable/i ) {
                $scalerWeightsConsumption = Dataset(
                    name => 'Scaler weighting on consumption charges',
                    @scalerWeightsItems, usePlaceholderData => 1,
                );
                $scalerWeightsStanding = Dataset(
                    name => 'Scaler weighting on standing charges',
                    @scalerWeightsItems, usePlaceholderData => 1,
                );
                Columnset(
                    name =>
                      'Tariff-specific weightings for revenue matching scaler',
                    appendTo => $model->{inputTables},
                    dataset  => $model->{dataset},
                    number   => 1079,
                    columns =>
                      [ $scalerWeightsConsumption, $scalerWeightsStanding, ],
                );
            }
            else {
                $scalerWeightsConsumption = $scalerWeightsStanding = Constant(
                    name =>
                      'Tariff-specific weighting for revenue matching scaler',
                    @scalerWeightsItems,
                );
            }
        }

        my $levelled = $model->{scaler} && $model->{scaler} =~ /levelled/i;

        my $capped = $model->{scaler} && $model->{scaler} =~ /capped/i;

        push @{ $model->{optionLines} },
          $levelled
          ? 'Revenue matching by '
          . (
            $model->{scaler} =~ /pick/i
            ? (
                $model->{scaler} =~ /ehv/i
                ? 'same £/kW/year at all EHV'
                : $model->{scaler} =~ /exit/i
                ? '£/kW/year at transmission exit'
                : '£/kW/year at selected'
              )
            : 'same £/kW/year at all'
          )
          . (
              $model->{scaler} =~ /exit/i  ? ' level'
            : $model->{scaler} =~ /opass/i ? ' operating and asset levels'
            : $model->{scaler} =~ /op/i    ? ' operating levels'
            : $model->{scaler} =~ /pick/i  ? ' levels'
            :                                ' asset levels'
          )
          : 'Revenue matching by scaler'
          . (
            $model->{scaler} && $model->{scaler} =~ /pick/i
            ? ' at '
              . ( $model->{scaler} =~ /ehv/i ? 'EHV' : 'selected' )
              . ' levels'
            : ''
          );

        push @{ $model->{optionLines} },
          'Scaler is capped'
          . (
            $model->{scaler} && $model->{scaler} =~ /cappedwithadder/i
            ? ' with adder for surplus'
            : ''
          ) if $capped;

        my $assetFlag;

        if ( !$levelled ) {
            $assetFlag =
              $model->{scaler} && $model->{scaler} =~ /pick/i
              ? (
                $model->{scaler} =~ /exit/i
                ? Constant(
                    name => 'Which charging elements are subject to the scaler',
                    cols => $chargingLevels,
                    data =>
                      [ map { /exit/i ? 1 : 0 } @{ $chargingLevels->{list} } ],
                    defaultFormat => '0connz'
                  )
                : $model->{scaler} =~ /all/i ? Constant(
                    name => 'Which charging elements are subject to the scaler',
                    cols => $chargingLevels,
                    data => [ map { 1 } @{ $chargingLevels->{list} } ],
                    defaultFormat => '0connz'
                  )
                : $model->{scaler} =~ /ehv/i ? Constant(
                    name => 'Which charging elements are subject to the scaler',
                    cols => $chargingLevels,
                    data => [
                        map { /exit|operating/i && /132|ehv/i ? 0 : 1 }
                          @{ $chargingLevels->{list} }
                    ],
                    defaultFormat => '0connz'
                  )
                : Dataset(
                    name       => 'Which levels are subject to the scaler',
                    validation => {
                        validate => 'decimal',
                        criteria => 'between',
                        minimum  => -999_999.999,
                        maximum  => 999_999.999,
                    },
                    cols          => $chargingLevels,
                    data          => [ map { 1 } @{ $chargingLevels->{list} } ],
                    defaultFormat => '0hardnz'
                )
              )
              : Constant(
                name => 'Which charging elements are subject to the scaler',
                cols => $chargingLevels,
                data => [
                    map { /exit|operating/i ? 0 : 1 }
                      @{ $chargingLevels->{list} }
                ],
                defaultFormat => '0connz'
              );
        }
        else {
            if (    $model->{scaler} =~ /pick/i
                and $model->{scaler} !~ /ehv/i || $model->{scaler} =~ /opass/i )
            {

                $assetFlag = Stack(
                    name    => 'Applicability factor for £1/kW/year scaler',
                    cols    => $chargingLevels,
                    sources => [
                        $model->{scaler} =~ /exit|ehv/i
                        ? Arithmetic(
                            name => 'Factor to scale to £1/kW/year '
                              . (
                                $model->{scaler} =~ /exit/i
                                ? 'at transmission exit level'
                                : 'at each level'
                              ),
                            arithmetic => '=IF(A1,1/A2,0)',
                            rows       => 0,
                            cols       => $model->{scaler} =~ /exit/i
                            ? Labelset(
                                name => 'Transmission exit level',
                                list => [
                                    grep { /exit/i }
                                      @{ $costToSml->{cols}{list} }
                                ]
                              )
                            : $model->{scaler} =~ /ehv/i ? Labelset(
                                name => 'EHV levels',
                                list => [
                                    grep { m#132|EHV# }
                                      @{ $costToSml->{cols}{list} }
                                ]
                              )
                            : $costToSml->{cols},
                            arguments => {
                                A1 => $costToSml,
                                A2 => $costToSml,
                            }
                          )

                        : Arithmetic(
                            name => 'Factor to scale to £1/kW/year '
                              . ' at each level',
                            arithmetic => '=IF(A1,A3/A2,0)',
                            rows       => 0,
                            cols       => $costToSml->{cols},
                            arguments  => {
                                A1 => $costToSml,
                                A2 => $costToSml,
                                A3 => Dataset(
                                    name =>
                                      'Which levels are subject to the scaler',
                                    rows       => $costToSml->{cols},
                                    validation => {
                                        validate => 'decimal',
                                        criteria => 'between',
                                        minimum  => -999_999.999,
                                        maximum  => 999_999.999,
                                    },
                                    data => [
                                        map { 1 } @{ $costToSml->{cols}{list} }
                                    ]
                                )
                            }
                        ),
                        Constant(
                            name => 'Zero for other levels',
                            cols => $chargingLevels,
                            data =>
                              [ map { [0] } @{ $chargingLevels->{list} } ],
                        )
                      ]

                );
            }

            else {
                my @factors;
                if ( $model->{scaler} =~ /op/i ) {
                    push @factors,
                      Arithmetic(
                        name => 'Factor to scale to £1/kW/year (operating)'
                          . ' at each level',
                        arithmetic => '=IF(A1,1/A2,0)',
                        rows       => 0,
                        $model->{scaler} =~ /ehv/i
                        ? (
                            cols => Labelset(
                                name => 'EHV levels',
                                list => [
                                    grep { m#132|EHV# }
                                      @{ $operatingCostToSml->{cols}{list} }
                                ]
                            )
                          )
                        : (),
                        arguments => {
                            A1 => $operatingCostToSml,
                            A2 => $operatingCostToSml,
                        }
                      );
                }
                else {
                    push @factors,
                      $model->{scaler} =~ /ehv/
                      ? Arithmetic(
                        name => 'Factor to scale to £1/kW/year (assets)'
                          . ' at each level',
                        arithmetic => '=IF(A1,A3/A2,0)',
                        cols       => $modelCostToSml->{rows},
                        rows       => 0,
                        arguments  => {
                            A1 => $modelCostToSml,
                            A2 => $modelCostToSml,
                            A3 => Constant(
                                name => 'Which network levels get the scaler',
                                cols => $modelCostToSml->{rows},
                                defaultFormat => '0connz',
                                data          => [
                                    map { /132|ehv/i ? 1 : 0; }
                                      @{ $modelCostToSml->{rows}{list} }
                                ]
                            )
                        }
                      )
                      : Arithmetic(
                        name => 'Factor to scale to £1/kW/year (assets)'
                          . ' at each level',
                        arithmetic => '=IF(A1,1/A2,0)',
                        cols       => $modelCostToSml->{rows},
                        rows       => 0,
                        arguments =>
                          { A1 => $modelCostToSml, A2 => $modelCostToSml, }
                      );
                }
                $assetFlag = Stack(
                    name =>
                      'Applicability factor for £1/kW/year (assets) scaler',
                    cols    => $chargingLevels,
                    sources => [
                        @factors,
                        Constant(
                            name => 'Zero for other levels',
                            cols => $chargingLevels,
                            data =>
                              [ map { [0] } @{ $chargingLevels->{list} } ],
                        )
                    ]
                );
            }
        }

        my $assetElements = {
            map {
                $_ => SumProduct(
                    name   => "$_ scalable part",
                    matrix => $tariffsExMatching->{$_}{source},
                    vector => $assetFlag
                );
            } @$scaledComponents
        };

        push @{ $model->{assetScaler} },
          Columnset(
            name    => 'Scalable elements of tariff components',
            columns => [ @{$assetElements}{@$scaledComponents} ]
          );

        $assetScalerPot =
          Labelset( list => [ $model->{scaler} ? 'Scaler' : 'Asset scaler' ] );

        my $totalRevenuesFromScalable;

        if ( $model->{scaler} && $model->{scaler} =~ /minzero/ ) {

            # madcapping of scaler (ignoring any site-specific elements)

            if ($capped) {  # before madcapping: if we are capping total scaling

                my $max = Arithmetic(
                    name => 'Maximum revenue that can be recovered by scaler',
                    defaultFormat => '0softnz',
                    arithmetic    => '=A1*A4',
                    arguments     => {
                        A1 => Dataset(
                            name  => 'Maximum scaling',
                            lines => 'Set this to zero to prohibit scaling up.',
                            defaultFormat => '%hard',
                            number        => 1012,
                            appendTo      => $model->{inputTables},
                            dataset       => $model->{dataset},
                            validation    => {
                                validate => 'decimal',
                                criteria => '>=',
                                minimum  => 0,
                                error_message =>
                                  'Must be a non-negative value.',
                            },
                            data => [
                                [ $model->{scaler} =~ /([0-9.]+)/ ? $1 : 0.0 ]
                            ],
                        ),
                        A4 => $totalRevenuesSoFar,
                    }
                );

                $revenueShortfall = Arithmetic(
                    name          => 'Revenue to be recovered by scaler',
                    defaultFormat => '0softnz',
                    arithmetic    => '=MIN(A1,A2)',
                    arguments     => {
                        A1 => $revenueShortfall,
                        A2 => $max,
                    }
                );

                Columnset(
                    name    => 'Application of scaler cap',
                    columns => [ $max, $revenueShortfall ]
                );

            }

            # now get on with madcapping

            my @slope = map {
                my $weighting =
                  /day/ ? $scalerWeightsStanding : $scalerWeightsConsumption;
                Arithmetic(
                    name          => "Effect through $_",
                    defaultFormat => '0softnz',
                    arithmetic    => (
                          /day/                ? '=A4*A2*A1/100'
                        : $dontScaleGeneration ? '=IF(A3<0,0,A4*A1*10)'
                        :                        '=A4*A1*10'
                      )
                      . ( $weighting ? '*A5' : '' ),
                    arguments => {
                        /day/ || !$dontScaleGeneration ? ()
                        : ( A3 => $loadCoefficients ),
                        A2 => $daysInYear,
                        A1 => $volumeData->{$_},
                        A4 => $assetElements->{$_},
                        $weighting ? ( A5 => $weighting ) : (),
                    },
                );
            } @$scaledComponents;

            push @{ $model->{assetScaler} },
              my $slopeSet = Columnset(
                name    => 'Marginal revenue effect of scaler',
                columns => \@slope,
              );

            my %minScaler = map {
                my $weighting =
                  /day/ ? $scalerWeightsStanding : $scalerWeightsConsumption;
                my $tariffComponent = $_;
                $_ => Arithmetic(
                    name       => "Scaler threshold for $_",
                    arithmetic => $weighting ? '=IF(A3*A31,0-A1/(A2*A21),0)'
                    : '=IF(A3,0-A1/A2,0)',
                    arguments => {
                        A1 => $tariffsExMatching->{$_},
                        A2 => $assetElements->{$_},
                        A3 => $assetElements->{$_},
                        $weighting ? ( A21 => $weighting, A31 => $weighting, )
                        : (),
                    },
                    rowFormats => [
                        map {
                            $componentMap->{$_}{$tariffComponent} ? undef
                              : 'unavailable';
                        } @{ $allTariffsByEndUser->{list} }
                    ]
                );
            } @$scaledComponents;

            my $minScalerSet = Columnset(
                name    => 'Scaler value at which the minimum is breached',
                columns => [ @minScaler{@$scaledComponents} ]
            );

            push @{ $model->{optionLines} },
              'Scaler subject to capping of each tariff component to zero';

            my $scalerRate = new SpreadsheetModel::SegmentRoot(
                name   => 'General scaler rate',
                slopes => $slopeSet,
                target => $revenueShortfall,
                min    => $minScalerSet,
            );

            $scalerTable = {
                map {
                    my $weighting =
                      /day/
                      ? $scalerWeightsStanding
                      : $scalerWeightsConsumption;
                    my $factor1 = $weighting ? '*A71' : '';
                    my $factor2 = $weighting ? '*A72' : '';
                    my $if      = /kWh/
                      ? "IF(IF(A41<0,-1,1)*(A1*A3$factor1+A9)>0,A11*A31$factor2,0-A91)"
                      : "IF(A1*A3$factor1+A9>0,A11*A31$factor2,0-A91)";
                    $_ => Arithmetic(
                        name       => "$_ scaler",
                        cols       => $assetScalerPot,
                        arithmetic => /day/
                          || !$dontScaleGeneration ? "=$if" : "=IF(A4<0,0,$if)",
                        arguments => {
                            A1  => $assetElements->{$_},
                            A11 => $assetElements->{$_},
                            A3  => $scalerRate,
                            A31 => $scalerRate,
                            !/day/ && $dontScaleGeneration
                            ? ( A4 => $loadCoefficients )
                            : (),
                            /kWh/ ? ( A41 => $loadCoefficients ) : (),
                            $weighting
                            ? ( A71 => $weighting, A72 => $weighting, )
                            : (),
                            A9  => $tariffsExMatching->{$_},
                            A91 => $tariffsExMatching->{$_},
                        }
                    );
                } @$scaledComponents
            };

            if ($totalSiteSpecificReplacement) {
                $siteSpecificCharges = Arithmetic(
                    name =>
                      'Total site specific sole use asset charges (£/year)',
                    defaultFormat => '0softnz',
                    cols          => $siteSpecificReplacement->{cols}
                    ? Labelset(
                        name => 'Site specific levels',
                        list => [
                            map { "Site-specific $_" }
                              @{ $siteSpecificReplacement->{cols}{list} }
                        ],
                        accepts => [
                            $siteSpecificReplacement->{cols},
                            $siteSpecificOperatingCost->{cols}
                        ]
                      )
                    : 0,
                    arithmetic => '=A2*(1+A3)+A1',
                    arguments  => {
                        A1 => $siteSpecificOperatingCost,
                        A2 => $siteSpecificReplacement,
                        A3 => $scalerRate,
                    }
                );
            }

        }

        else {

            my $revenuesFromAsset;

            {
                my @termsNoDays;
                my @termsWithDays;
                my %args = ( A400 => $daysInYear );
                my $i = 1;
                foreach (@$nonExcludedComponents) {
                    ++$i;
                    my $pad = "$i";
                    $pad = "0$pad" while length $pad < 3;
                    if (m#/day#) {
                        push @termsWithDays, "A2$pad*A3$pad";
                    }
                    else {
                        push @termsNoDays, "A2$pad*A3$pad";
                    }
                    $args{"A2$pad"} = $assetElements->{$_};
                    $args{"A3$pad"} = $volumeData->{$_};
                }
                $revenuesFromAsset = Arithmetic(
                    name => 'Net revenues to which the scaler applies',
                    rows => $dontScaleGeneration ? $demandTariffsByEndUser
                    : $allTariffsByEndUser,
                    arithmetic => '='
                      . join( '+',
                        @termsWithDays
                        ? ( '0.01*A400*(' . join( '+', @termsWithDays ) . ')' )
                        : ('0'),
                        @termsNoDays
                        ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
                        : ('0'),
                      ),
                    arguments     => \%args,
                    defaultFormat => '0soft'
                  )
            };

            $totalRevenuesFromScalable = GroupBy(
                name   => 'Total net revenues from scalable elements (£)',
                rows   => 0,
                cols   => 0,
                source => $revenuesFromAsset,
                defaultFormat => '0soft'
            );

            $totalRevenuesFromScalable = Arithmetic(
                name => 'Total net revenues from scalable elements'
                  . ' including site specific sole use assets' . ' (£)',
                arithmetic => '=A1+A2',
                arguments  => {
                    A1 => $totalRevenuesFromScalable,
                    A2 => $totalSiteSpecificReplacement
                },
                defaultFormat => '0soft'
            ) if $totalSiteSpecificReplacement;

            push @{ $model->{assetScaler} }, $totalRevenuesFromScalable;

            if ($capped) {

                my $max = Arithmetic(
                    name => 'Maximum revenue that can be recovered by scaler',
                    defaultFormat => '0softnz',
                    arithmetic    => '=A1*A4',
                    arguments     => {
                        A1 => Dataset(
                            name => 'Maximum scaling'
                              . ' (set to zero to prohibit scaling up)',
                            defaultFormat => '%hard',
                            number        => 1012,
                            appendTo      => $model->{inputTables},
                            dataset       => $model->{dataset},
                            validation    => {
                                validate => 'decimal',
                                criteria => '>=',
                                minimum  => 0,
                                error_message =>
                                  'Must be a non-negative value.',
                            },
                            data => [
                                [ $model->{scaler} =~ /([0-9.]+)/ ? $1 : 0.0 ]
                            ],
                        ),
                        A4 => $totalRevenuesFromScalable,
                    }
                );

                $capped = Arithmetic(
                    name          => 'Revenue to be recovered by scaler',
                    defaultFormat => '0softnz',
                    arithmetic    => '=MIN(A1,A2)',
                    arguments     => {
                        A1 => $revenueShortfall,
                        A2 => $max,
                    }
                );

                Columnset(
                    name    => 'Application of scaler cap',
                    columns => [ $max, $capped ]
                );

            }

            $scalerTable = {
                map {
                    $_ => Arithmetic(
                        name       => "$_ scaler",
                        cols       => $assetScalerPot,
                        arithmetic => $dontScaleGeneration
                        ? '=IF(A4<0,0,A1/A2*A3)'
                        : '=A1/A2*A3',
                        arguments => {
                            A1 => $assetElements->{$_},
                            A2 => $totalRevenuesFromScalable,
                            A3 => $capped || $revenueShortfall,
                            A4 => $loadCoefficients
                        }
                    );
                } @$scaledComponents
            };

            if ($totalSiteSpecificReplacement) {
                $siteSpecificCharges = Arithmetic(
                    name => 'Total site specific sole use asset charges (£)',
                    defaultFormat => '0softnz',
                    cols          => Labelset(
                        name => 'Site specific levels',
                        list => [
                            map { "Site-specific $_" }
                              @{ $siteSpecificReplacement->{cols}{list} }
                        ],
                        accepts => [
                            $siteSpecificReplacement->{cols},
                            $siteSpecificOperatingCost->{cols}
                        ]
                    ),
                    arithmetic => '=A2*(1+A3/A4)+A1',
                    arguments  => {
                        A1 => $siteSpecificOperatingCost,
                        A2 => $siteSpecificReplacement,
                        A3 => $revenueShortfall,
                        A4 => $totalRevenuesFromScalable
                    }
                );
            }

        }    # not madcapping

        my $revenuesFromScaler;

        {
            my @termsNoDays;
            my @termsWithDays;
            my %args = ( A400 => $daysInYear );
            my $i = 1;

            foreach ( grep { $scalerTable->{$_} } @$nonExcludedComponents ) {
                ++$i;
                my $pad = "$i";
                $pad = "0$pad" while length $pad < 3;
                if (m#/day#) {
                    push @termsWithDays, "A2$pad*A3$pad";
                }
                else {
                    push @termsNoDays, "A2$pad*A3$pad";
                }
                $args{"A2$pad"} = $scalerTable->{$_};
                $args{"A3$pad"} = $volumeData->{$_};
            }

            $revenuesFromScaler = Arithmetic(
                name       => 'Net revenues by tariff from scaler',
                rows       => $allTariffsByEndUser,
                cols       => $assetScalerPot,
                arithmetic => '='
                  . join( '+',
                    @termsWithDays
                    ? ( '0.01*A400*(' . join( '+', @termsWithDays ) . ')' )
                    : ('0'),
                    @termsNoDays ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
                    : ('0'),
                  ),
                arguments     => \%args,
                defaultFormat => '0softnz'
            );

            push @{ $model->{assetScaler} }, Columnset
              name    => 'Scaler',
              columns => [
                ( grep { $_ } @{$scalerTable}{@$scaledComponents} ),
                $revenuesFromScaler
              ]

        };

        $totalRevenuesFromScaler = GroupBy(
            name          => 'Total net revenues from scaler (£)',
            rows          => 0,
            cols          => 0,
            source        => $revenuesFromScaler,
            defaultFormat => '0softnz',
        );

        if ($totalRevenuesFromScalable) {

            my $assetRate =
              $levelled ? Arithmetic(
                name       => 'Scaler £/kW/year (by level)',
                arithmetic => '=A2/A3',
                arguments  => {
                    A2 => $totalRevenuesFromScaler,
                    A3 => $totalRevenuesFromScalable
                }
              )
              : $model->{scaler} && $model->{scaler} =~ /pick/i ? Arithmetic(
                name       => 'Scaling factor',
                arithmetic => '=A2/A3',
                arguments  => {
                    A2 => $totalRevenuesFromScaler,
                    A3 => $totalRevenuesFromScalable
                }
              )
              : Arithmetic(
                name          => 'Effective asset annuity rate',
                defaultFormat => '%softnz',
                arithmetic    => '=A8*(1+A2/A3)',
                arguments     => {
                    A8 => $annuityRate,
                    A2 => $totalRevenuesFromScaler,
                    A3 => $totalRevenuesFromScalable
                }
              );

            unless ($levelled) {
                push @{ $model->{summaryColumns} }, $assetRate;
                $totalRevenuesFromScaler->{mustCopy} = 1;
            }

        }

    }    # end of code that only runs if not an adder

    my ( $totalRevenuesFromAdder, $adderTable, $excludeScaler,
        $excludeReplacement );

    my $totalRevenueRemoved;

    if (   $model->{scaler} && $model->{scaler} !~ /levelled|pick|capped/i
        || $model->{noReplacement} && $model->{noReplacement} =~ /hybrid/i
        || $model->{scaler} && $model->{scaler} =~ /cappedwithadder/i )
    {

        my $tariffsByAtw = Labelset(
            name   => 'All tariffs by ATW tariff',
            groups => [
                map { Labelset( name => $_->{list}[0], list => $_->{list} ) }
                  @{
                         $allTariffsByEndUser->{groups}
                      || $allTariffsByEndUser->{list}
                  }
            ]
        );

        my $atwTariffs = Labelset(
            name => 'All ATW tariffs',
            list => $tariffsByAtw->{groups}
        );

        if ( $model->{scaler} && $model->{scaler} =~ /hybrid/i ) {

            my $excludeScalerPot = Labelset( list => ['Remove ATW scaler'] );

            $excludeScaler = {
                map {
                    $_ => Arithmetic(
                        name       => "$_ remove ATW scaler",
                        cols       => $excludeScalerPot,
                        rows       => $tariffsByAtw,
                        arithmetic => '=0-A1',
                        arguments  => {
                            A1 => Arithmetic(
                                name       => "$_ ATW scaler",
                                rows       => $atwTariffs,
                                arithmetic => '=A1',
                                arguments  => { A1 => $scalerTable->{$_} }
                            ),
                        }
                      )
                } grep { $scalerTable->{$_} } @$allComponents
            };

            push @{ $model->{optionLines} },
              'Adder to recover removed revenues from scaler';

        }

        if (   $model->{noReplacement}
            && $model->{noReplacement} =~ /hybrid/i )
        {

            my $excludeReplacementPot =
              Labelset( list => ['Remove ATW replacement'] );

            $excludeReplacement = {
                map {
                    $_ => Arithmetic(
                        name       => "$_ remove ATW replacement",
                        cols       => $excludeReplacementPot,
                        rows       => $tariffsByAtw,
                        arithmetic => '=0-A1',
                        arguments  => {
                            A1 => SumProduct(
                                name   => "$_ ATW replacement",
                                cols   => 0,
                                rows   => $atwTariffs,
                                matrix => $tariffsExMatching->{$_}{source},
                                vector => $replacementShare
                            ),
                        }
                      )
                } grep { $tariffsExMatching->{$_}{source} } @$allComponents
            };

            push @{ $model->{optionLines} },
              'Adder to recover removed replacement cost amounts';

        }

        if ( $excludeScaler || $excludeReplacement ) {

            push @{ $model->{adderResults} },
              Columnset(
                name => 'Identification of '
                  . 'all-the-way tariff elements'
                  . ' to be removed',
                columns => [
                    map { ${ $_->{arguments} }{A1} }
                      $excludeScaler
                    ? @{$excludeScaler}{ grep { $excludeScaler->{$_} }
                          @$allComponents }
                    : (),
                    $excludeReplacement
                    ? @{$excludeReplacement}{ grep { $excludeReplacement->{$_} }
                          @$allComponents }
                    : (),
                ]
              );

            my $revenuesRemoved;

            {
                my @termsNoDays;
                my @termsWithDays;
                my %args = ( A400 => $daysInYear );
                my $i = 1;

                if ($excludeScaler) {
                    foreach ( grep { $excludeScaler->{$_} }
                        @$nonExcludedComponents )
                    {
                        ++$i;
                        my $pad = "$i";
                        $pad = "0$pad" while length $pad < 3;
                        if (m#/day#) {
                            push @termsWithDays, "A2$pad*A3$pad";
                        }
                        else {
                            push @termsNoDays, "A2$pad*A3$pad";
                        }
                        $args{"A2$pad"} = $excludeScaler->{$_};
                        $args{"A3$pad"} = $volumeData->{$_};
                    }
                }

                if ($excludeReplacement) {
                    foreach ( grep { $excludeReplacement->{$_} }
                        @$nonExcludedComponents )
                    {
                        ++$i;
                        my $pad = "$i";
                        $pad = "0$pad" while length $pad < 3;
                        if (m#/day#) {
                            push @termsWithDays, "A2$pad*A3$pad";
                        }
                        else {
                            push @termsNoDays, "A2$pad*A3$pad";
                        }
                        $args{"A2$pad"} = $excludeReplacement->{$_};
                        $args{"A3$pad"} = $volumeData->{$_};
                    }
                }

                $revenuesRemoved = Arithmetic(
                    name       => 'Lost revenues by tariff from deductions',
                    rows       => $tariffsByAtw,
                    arithmetic => '=0-'
                      . join( '-',
                        @termsWithDays
                        ? ( '0.01*A400*(' . join( '+', @termsWithDays ) . ')' )
                        : ('0'),
                        @termsNoDays
                        ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
                        : ('0'),
                      ),
                    arguments     => \%args,
                    defaultFormat => '0softnz'
                );

                push @{ $model->{adderResults} },
                  Columnset(
                    name => 'Removal of all-the-way tariff elements'
                      . ' from all tariffs',
                    columns => [
                        $excludeScaler
                        ? @{$excludeScaler}{ grep { $excludeScaler->{$_} }
                              @$allComponents }
                        : (),
                        $excludeReplacement ? @{$excludeReplacement}{
                            grep { $excludeReplacement->{$_} } @$allComponents
                          }
                        : (),
                        $revenuesRemoved
                    ]
                  );

            };

            push @{ $model->{adderResults} },
              $totalRevenueRemoved = GroupBy(
                name          => 'Total net revenues transferred to adder (£)',
                rows          => 0,
                cols          => 0,
                source        => $revenuesRemoved,
                defaultFormat => '0softnz'
              );

        }

        my $adderAmount;

        if ( $model->{scaler} && $model->{scaler} =~ /cappedwithadder/i ) {
            $adderAmount = Arithmetic(
                name          => 'Amount to be recovered from adder (£)',
                defaultFormat => '0softnz',
                arithmetic    => '=A1-A2',
                arguments =>
                  { A1 => $revenueShortfall, A2 => $totalRevenuesFromScaler }
            );
        }
        elsif ( $model->{scaler} && $model->{scaler} =~ /adder/i ) {
            $adderAmount = $revenueShortfall;
            push @{ $model->{optionLines} }, 'Revenue matching by adder';
        }

        push @{ $model->{optionLines} },
          $model->{scaler} !~ /ppu/i ? 'Adder: single £/kW/year'
          : $model->{scaler} =~ /ppupercent/i
          ? 'Adder: single p/kWh converted to percentage'
          : $model->{scaler} =~ /ppumultiple/i ? 'Adder: p/kWh at each level'
          :                                      'Adder: single p/kWh';

        if ($totalRevenueRemoved) {
            if ($adderAmount) {
                $adderAmount = Arithmetic(
                    name          => 'Total net revenue needed from adder (£)',
                    defaultFormat => '0softnz',
                    arithmetic    => '=A1+A2',
                    arguments =>
                      { A1 => $revenueShortfall, A2 => $totalRevenueRemoved }
                );
            }
            else { $adderAmount = $totalRevenueRemoved; }
        }

        if ( $model->{scaler} =~ /ppugeneral/i )
        {    # single adder with caps and collars, on demand only by default

            my @columns = grep { /kWh/ } @$nonExcludedComponents;
            my @slope = map {
                Arithmetic(
                    name          => "Effect through $_",
                    arithmetic    => '=IF(A3<0,0,A1*10)',
                    defaultFormat => '0soft',
                    arguments     => {
                        A3 => $loadCoefficients,
                        A1 => $volumeData->{$_},
                    },
                );
            } @columns;

            my $slopeSet = Columnset(
                name    => 'Marginal revenue effect of adder',
                columns => \@slope,
            );

            my ( %minAdder, $minAdderSet, %maxAdder, $maxAdderSet );

            if ( $model->{scaler} =~ /min/i ) {
                my %min = map {
                    my $tariffComponent = $_;
                    $_ => $model->{scaler} =~ /zeroexplicit/i
                      ? Constant(
                        name => "Minimum $_",
                        rows => $allTariffsByEndUser,
                        data => [ map { 0 } @{ $allTariffsByEndUser->{list} } ],
                        rowFormats => [
                            map {
                                $componentMap->{$_}{$tariffComponent}
                                  ? undef
                                  : 'unavailable';
                            } @{ $allTariffsByEndUser->{list} }
                        ]
                      )
                      : $model->{scaler} =~ /zero/i ? undef
                      : Dataset(
                        name       => "Minimum $_",
                        rows       => $allTariffsByEndUser,
                        validation => {
                            validate => 'decimal',
                            criteria => 'between',
                            minimum  => -999_999.999,
                            maximum  => 999_999.999,
                        },
                        usePlaceholderData => 1,
                        data               => [
                            map {
                                $componentMap->{$_}{$tariffComponent} ? ''
                                  : undef;
                            } @{ $allTariffsByEndUser->{list} }
                        ],
                      );
                } @columns;
                if ( my @cols = grep { $_ } @min{@columns} ) {
                    Columnset(
                        name     => 'Minimum rates',
                        number   => 1077,
                        columns  => \@cols,
                        dataset  => $model->{dataset},
                        appendTo => $model->{inputTables},
                    );
                }
                %minAdder = map {
                    my $tariffComponent = $_;

                    $_ => $min{$_}
                      ? Arithmetic(
                        name       => "Adder threshold for $_",
                        arithmetic => '=IF(ISNUMBER(A3),A2-A1,"")',
                        arguments  => {
                            A1 => $tariffsExMatching->{$_},
                            A2 => $min{$_},
                            A3 => $min{$_}
                        },
                        rowFormats => [
                            map {
                                $componentMap->{$_}{$tariffComponent}
                                  ? undef
                                  : 'unavailable';
                            } @{ $allTariffsByEndUser->{list} }
                        ]
                      )
                      : Arithmetic(
                        name       => "Adder threshold for $_",
                        arithmetic => '=0-A1',
                        arguments  => {
                            A1 => $tariffsExMatching->{$_},
                        },
                        rowFormats => [
                            map {
                                $componentMap->{$_}{$tariffComponent} ? undef
                                  : 'unavailable';
                            } @{ $allTariffsByEndUser->{list} }
                        ]
                      );
                } @columns;
                $minAdderSet = Columnset(
                    name    => 'Adder value at which the minimum is breached',
                    columns => [ @minAdder{@columns} ]
                );
            }

            if ( $model->{scaler} =~ /max/i ) {
                my %max = map {
                    my $tariffComponent = $_;
                    $_ => Dataset(
                        name       => "Maximum $_",
                        rows       => $allTariffsByEndUser,
                        validation => {
                            validate => 'decimal',
                            criteria => 'between',
                            minimum  => -999_999.999,
                            maximum  => 999_999.999,
                        },
                        usePlaceholderData => 1,
                        data               => [
                            map {
                                $componentMap->{$_}{$tariffComponent}
                                  ? ''
                                  : undef;
                            } @{ $allTariffsByEndUser->{list} }
                        ],
                    );
                } @columns;
                Columnset(
                    name     => 'Maximum rates',
                    number   => 1078,
                    columns  => [ @max{@columns} ],
                    dataset  => $model->{dataset},
                    appendTo => $model->{inputTables},
                );
                %maxAdder = map {
                    my $tariffComponent = $_;

                    $_ => Arithmetic(
                        name       => "Adder threshold for $_",
                        arithmetic => '=IF(ISNUMBER(A3),A2-A1,"")',
                        arguments  => {
                            A1 => $tariffsExMatching->{$_},
                            A2 => $max{$_},
                            A3 => $max{$_}
                        },
                        rowFormats => [
                            map {
                                $componentMap->{$_}{$tariffComponent}
                                  ? undef
                                  : 'unavailable';
                            } @{ $allTariffsByEndUser->{list} }
                        ]
                    );
                } @columns;
                $maxAdderSet = Columnset(
                    name    => 'Adder value at which the maximum is breached',
                    columns => [ @maxAdder{@columns} ]
                );
            }

            push @{ $model->{optionLines} },
              'p/kWh revenue matching with capping'
              if $minAdderSet || $maxAdderSet;

            my $adderRate = new SpreadsheetModel::SegmentRoot(
                name   => 'General adder rate (p/kWh)',
                slopes => $slopeSet,
                target => $adderAmount,
                min    => $minAdderSet,
                max    => $maxAdderSet
            );

            $adderTable = {
                map {
                    my $tariffComponent = $_;
                    my $iv              = 'A1';
                    $iv = "MAX($iv,A5)" if $minAdder{$_};
                    $iv = "MIN($iv,A6)" if $maxAdder{$_};
                    $_ => Arithmetic(
                        name       => "Adder on $_",
                        rows       => $allTariffsByEndUser,
                        cols       => Labelset( list => ['Adder'] ),
                        arithmetic => "=IF(A3<0,0,$iv)",
                        arguments  => {
                            A1 => $adderRate,
                            A3 => $loadCoefficients,
                            $minAdder{$_} ? ( A5 => $minAdder{$_} ) : (),
                            $maxAdder{$_} ? ( A6 => $maxAdder{$_} ) : ()
                        },
                        rowFormats => [
                            map {
                                $componentMap->{$_}{$tariffComponent}
                                  ? undef
                                  : 'unavailable';
                            } @{ $allTariffsByEndUser->{list} }
                        ]
                    );
                } @columns
            };

        }

        elsif ( $model->{scaler} =~ /ppuflex/i ) {
            $adderTable =
              $model->flexibleAdder( $allTariffsByEndUser,
                $nonExcludedComponents, $componentMap, $volumeData,
                $tariffsExMatching, $adderAmount, );
        }

        else {    #  various legacy adder options but no caps/collars

            my $includeGenerators = Arithmetic(
                name       => 'Are generator tariffs subject to adder?',
                arithmetic => !$model->{genAdder} ? 'FALSE'
                : $model->{genAdder} =~ /charge/ ? '=A1>0'
                : $model->{genAdder} =~ /pay/    ? '=A1<0'
                : $model->{genAdder} =~ /always/ ? 'TRUE'
                : $model->{genAdder} =~ /never/  ? 'FALSE'
                : $model->{genAdder},
                arguments => { A1 => $adderAmount },
            );

            push @{ $model->{optionLines} },
              'Adder for generators: '
              . (
                 !$model->{genAdder}             ? 'never'
                : $model->{genAdder} =~ /charge/ ? 'if it is a charge'
                : $model->{genAdder} =~ /pay/    ? 'if it is a credit'
                : $model->{genAdder} =~ /always/ ? 'always'
                : $model->{genAdder} =~ /never/  ? 'never'
                :                                  'never'
              );

            my $gspPot = Labelset( list => [ $chargingLevels->{list}[0] ] );

            my $fixedAdderPot = Labelset(
                list    => ['Adder'],
                accepts => [$gspPot]
            );

            my ( $sml, $numberOfLevels );

            if ( $model->{scaler} =~ /ppumultiple/i ) {
                $numberOfLevels = GroupBy(
                    rows   => $unitsInYear->{rows},
                    source => $routeingFactors,
                    name   => 'Number of network levels'
                );
                $sml = GroupBy(
                    name   => 'Total MWh eligible for adder',
                    source => Arithmetic(
                        name       => 'MWh eligible for adder, by tariff',
                        cols       => $gspPot,
                        arithmetic => '=A4*IF(OR(A2,A3>=0),ABS(A1),0)',
                        arguments  => {
                            A1 => $unitsInYear,
                            A2 => $includeGenerators,
                            A3 => $loadCoefficients,
                            A4 => $numberOfLevels
                        }
                    )
                );
            }

            elsif ( $model->{scaler} =~ /ppu/i ) {
                $sml = GroupBy(
                    name   => 'Total MWh eligible for adder',
                    source => Arithmetic(
                        name       => 'MWh eligible for adder, by tariff',
                        cols       => $gspPot,
                        arithmetic => '=IF(OR(A2,A3>=0),ABS(A1),0)',
                        arguments  => {
                            A1 => $unitsInYear,
                            A2 => $includeGenerators,
                            A3 => $loadCoefficients
                        }
                    )
                );
            }

            else {
                $sml = GroupBy(
                    name => 'Total kW eligible for adder from'
                      . ' tariffs without multiple unit rate calculation',
                    source => Arithmetic(
                        name => 'Total kW eligible for adder from'
                          . ' tariffs without multiple unit rate calculation',
                        cols => $gspPot,
                        rows => Labelset(
                            name =>
                              'Tariffs without multiple unit rate calculation',
                            groups => [
                                grep {
                                    !( $componentMap->{$_}{'Unit rate 2 p/kWh'}
                                        || $componentMap->{$_}
                                        {'Unit rate 0 p/kWh'} )
                                  } @{
                                    $allTariffsByEndUser->{groups}
                                      || $allTariffsByEndUser->{list}
                                  }
                            ]
                        ),
                        arithmetic => '=IF(OR(A2,A3>=0),ABS(A1),0)',
                        arguments  => {
                            A1 => $simultaneousMaximumLoadUnits,
                            A2 => $includeGenerators,
                            A3 => $loadCoefficients
                        }
                    )
                );
            }

            push @{ $model->{adderResults} }, $sml;

            my @sml = ($sml);

            unless ( $model->{scaler} =~ /ppu/i ) {

                foreach my $r ( 0 .. $#unitRateSystemLoadCoefficients ) {
                    my $r1 = 1 + $r;

                    my $sml = GroupBy(
                        name => 'Total kW eligible for adder from'
                          . " rate $r1 units",
                        source => Arithmetic(
                            name => 'Total kW eligible for adder from'
                              . " rate $r1 units",
                            rows =>
                              $unitRateSystemLoadCoefficients[$r]{tariffs},
                            cols       => $gspPot,
                            arithmetic => '=IF(OR(A2,A3>=0),'
                              . 'ABS(A1),0)*A4*1000/(24*A6)*A5',
                            arguments => {
                                A1 => $unitRateSystemLoadCoefficients[$r],
                                A2 => $includeGenerators,
                                A3 => $unitRateSystemLoadCoefficients[$r],
                                A4 => $volumeData->{"Unit rate $r1 p/kWh"},
                                A5 => $lineLossFactorsToGsp,
                                A6 => $daysInYear,
                            }
                        )
                    );

                    push @{ $model->{adderResults} }, $sml;

                    push @sml, $sml;

                }

                if ( @{ $generationCapacityTariffsByEndUser->{list} } ) {

                    my $volumeForAdderCapacity = Arithmetic
                      name       => 'Volume eligible for adder',
                      cols       => $gspPot,
                      rows       => $generationCapacityTariffsByEndUser,
                      arithmetic => '=IF(OR(A2,A3>=0),ABS(A1),0)',
                      arguments  => {
                        A1 => $simultaneousMaximumLoadCapacity,
                        A2 => $includeGenerators,
                        A3 => $simultaneousMaximumLoadCapacity
                      };

                    push @{ $model->{adderResults} },
                      my $volumeForAdderCapacitySum = GroupBy
                      name => 'Total volume eligible for adder'
                      . ' from generation capacity rates',
                      source => $volumeForAdderCapacity;

                    push @sml, $volumeForAdderCapacitySum;

                }

            }

            my $volumeForAdder = @sml == 1 ? $sml[0] : Arithmetic(
                name       => 'Total volume eligible for adder',
                arithmetic => '=' . join( '+', map { "A$_" } 1 .. @sml ),
                arguments  => { map { ; "A$_" => $sml[ $_ - 1 ]; } 1 .. @sml }
            );

            push @{ $model->{adderResults} }, Columnset
              name    => 'Aggregation of total volume eligible for adder',
              columns => [ @sml, $volumeForAdder ]
              if @sml > 1;

            my $adderUnitYardstick =
              $model->{scaler} !~ /ppu/i
              ? Arithmetic(
                name => 'Adder yardstick p/kWh',
                rows => $allTariffsByEndUser,
                cols => $fixedAdderPot,
                arithmetic =>
                  '=IF(OR(A5,A2>=0),ABS(A1)*A7*A4/A3/(24*A6)*100,0)',
                arguments => {
                    A1 => $loadCoefficients,
                    A2 => $loadCoefficients,
                    A3 => $volumeForAdder,
                    A4 => $adderAmount,
                    A5 => $includeGenerators,
                    A6 => $daysInYear,
                    A7 => $lineLossFactorsToGsp
                }
              )
              : Arithmetic(
                name       => 'Adder yardstick p/kWh',
                rows       => $allTariffsByEndUser,
                cols       => $fixedAdderPot,
                arithmetic => '=IF(OR(A5,A2>=0),A4/A3/10,0)'
                  . ( $numberOfLevels ? '*A6' : '' ),
                arguments => {
                    A2 => $loadCoefficients,
                    A3 => $volumeForAdder,
                    A4 => $adderAmount,
                    A5 => $includeGenerators,
                    $numberOfLevels ? ( A6 => $numberOfLevels ) : ()
                }
              );

            push @{ $model->{adderResults} }, $adderUnitYardstick;

            if ( $model->{scaler} =~ /ppupercent/i ) {

                $adderTable = {
                    map {
                        $_ => Arithmetic(
                            name       => "$_ adder",
                            cols       => $fixedAdderPot,
                            arithmetic => '=IF(A94,A1/A2*A3/A4*A6,'
                              . ( /kWh/ ? 'A5)' : '0)' ),
                            arguments => {
                                A1  => $sml[0]->{source},
                                A2  => $volumeForAdder,
                                A3  => $revenueShortfall,
                                A4  => $revenuesSoFar,
                                A94 => $revenuesSoFar,
                                A5  => $adderUnitYardstick,
                                A6  => $tariffsExMatching->{$_}
                            }
                        );
                    } @$scaledComponents
                };

            }

            else {    # not ppupercent

                $adderTable = {
                    map {
                        my $r  = $_;
                        my $r1 = 1 + $r;
                        my $rateAdder =
                            $model->{scaler} =~ /ppu/i
                          ? $adderUnitYardstick
                          : Stack(
                            name    => "Unit rate $r1 p/kWh adder",
                            rows    => $allTariffsByEndUser,
                            cols    => $fixedAdderPot,
                            sources => [
                                Arithmetic(
                                    name => "Adder on rate $r1 (p/kWh)",
                                    rows => $unitRateSystemLoadCoefficients[$r]
                                      {tariffs},
                                    cols       => $fixedAdderPot,
                                    arithmetic => '=IF(OR(A5,A2>=0),'
                                      . 'ABS(A1)*A7*A4/A3/(24*A6)*100' . ',0)',
                                    arguments => {
                                        A1 =>
                                          $unitRateSystemLoadCoefficients[$r],
                                        A2 =>
                                          $unitRateSystemLoadCoefficients[$r],
                                        A3 => $volumeForAdder,
                                        A4 => $adderAmount,
                                        A5 => $includeGenerators,
                                        A6 => $daysInYear,
                                        A7 => $lineLossFactorsToGsp
                                    }
                                ),
                                $r ? () : ($adderUnitYardstick)
                            ]
                          );
                        "Unit rate $r1 p/kWh" => $rateAdder;
                      } 0 .. (
                        $model->{maxUnitRates} > 1
                        ? $model->{maxUnitRates} - 1
                        : -1
                      )
                };

                if ( @{ $generationCapacityTariffsByEndUser->{list} } ) {

                    my $adderCapacityYardstick =
                      $model->{scaler} =~ /ppu/i ? undef : Arithmetic(
                        name       => 'Generation capacity rate p/kW/day adder',
                        rows       => $generationCapacityTariffsByEndUser,
                        cols       => $fixedAdderPot,
                        arithmetic => '=IF(A5,ABS(A1)*A7*A4/A3/A6*100,0)',
                        arguments  => {
                            A1 => $fFactors,
                            A2 => $fFactors,
                            A3 => $volumeForAdder,
                            A4 => $adderAmount,
                            A5 => $includeGenerators,
                            A6 => $daysInYear,
                            A7 => $lineLossFactorsToGsp
                        }
                      );

                    $adderTable->{'Generation capacity rate p/kW/day'} = Stack(
                        name    => 'Generation capacity rate p/kW/day adder',
                        rows    => $allTariffsByEndUser,
                        cols    => $fixedAdderPot,
                        sources => [$adderCapacityYardstick]
                    );

                }
            }

            my $adderRate =
              $model->{scaler} =~ /ppu/i
              ? Arithmetic(
                name       => 'Adder rate p/kWh',
                arithmetic => '=A1/A2/10',
                arguments  => { A1 => $adderAmount, A2 => $volumeForAdder }
              )
              : Arithmetic(
                name          => 'Adder rate £/kW/year',
                defaultFormat => '0.00soft',
                arithmetic    => '=A1/A2',
                arguments     => { A1 => $adderAmount, A2 => $volumeForAdder }
              );

        }    # end of if ppuminmax

        my $revenuesFromAdder;

        {
            my @termsNoDays;
            my @termsWithDays;
            my %args = ( A400 => $daysInYear );
            my $i = 1;

            foreach ( grep { $adderTable->{$_} } @$nonExcludedComponents ) {
                ++$i;
                my $pad = "$i";
                $pad = "0$pad" while length $pad < 3;
                if (m#/day#) {
                    push @termsWithDays, "A2$pad*A3$pad";
                }
                else {
                    push @termsNoDays, "A2$pad*A3$pad";
                }
                $args{"A2$pad"} = $adderTable->{$_};
                $args{"A3$pad"} = $volumeData->{$_};
            }

            $revenuesFromAdder = Arithmetic(
                name       => 'Net revenues by tariff from adder',
                rows       => $allTariffsByEndUser,
                arithmetic => '='
                  . join( '+',
                    @termsWithDays
                    ? ( '0.01*A400*(' . join( '+', @termsWithDays ) . ')' )
                    : ('0'),
                    @termsNoDays ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
                    : ('0'),
                  ),
                arguments     => \%args,
                defaultFormat => '0softnz'
            );

            push @{ $model->{adderResults} },
              Columnset(
                name    => 'Adder',
                columns => [
                    ( grep { $_ } @{$adderTable}{@$nonExcludedComponents} ),
                    $revenuesFromAdder
                ]
              );

        };

        $totalRevenuesFromAdder = GroupBy(
            name          => 'Total net revenues from adder (£)',
            rows          => 0,
            cols          => 0,
            source        => $revenuesFromAdder,
            defaultFormat => '0softnz'
        );

    }

    $totalRevenuesFromScaler
      || (
        $totalRevenueRemoved
        ? Arithmetic(
            name          => 'Revenue from adder, net of revenue removed (£)',
            defaultFormat => '0softnz',
            arithmetic    => '=A1-A2',
            arguments =>
              { A1 => $totalRevenuesFromAdder, A2 => $totalRevenueRemoved }
        )
        : $totalRevenuesFromAdder
      ),
      $siteSpecificCharges, grep { $_ } $scalerTable, $excludeScaler,
      $excludeReplacement,  $adderTable;

}

1;
