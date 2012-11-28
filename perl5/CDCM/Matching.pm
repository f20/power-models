package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2012 DCUSA Limited and others. All rights reserved.

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
                arithmetic => '=IV2+IV1',
                arguments  => {
                    IV1 => $siteSpecificOperatingCost,
                    IV2 => $siteSpecificReplacement,
                }
            );
        }
    }

    else {    # not an adder

        my $dontScaleGeneration =
          $model->{scaler} && $model->{scaler} =~ /nogen/i;

        my $scalableTariffs =
            $dontScaleGeneration
          ? $demandTariffsByEndUser
          : $allTariffsByEndUser;

        my $levelled = $model->{scaler} && $model->{scaler} =~ /levelled/i;

        my $capped = $model->{scaler} && $model->{scaler} =~ /capped/i;

        push @{ $model->{optionLines} }, $levelled
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
            : ' asset levels'
          )
          : 'Revenue matching by scaler'
          . (
            $model->{scaler} && $model->{scaler} =~ /pick/i
            ? ' at '
              . ( $model->{scaler} =~ /ehv/i ? 'EHV' : 'selected' )
              . ' levels'
            : ''
          );

        push @{ $model->{optionLines} }, 'Scaler is capped'
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
                    name    => 'Applicability factor for £1/kW scaler',
                    cols    => $chargingLevels,
                    sources => [
                        $model->{scaler} =~ /exit|ehv/i
                        ? Arithmetic(
                            name => 'Factor to scale to £1/kW '
                              . (
                                $model->{scaler} =~ /exit/i
                                ? 'at transmission exit level'
                                : 'at each level'
                              ),
                            arithmetic => '=IF(IV1,1/IV2,0)',
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
                                IV1 => $costToSml,
                                IV2 => $costToSml,
                            }
                          )

                        : Arithmetic(
                            name => 'Factor to scale to £1/kW '
                              . ' at each level',
                            arithmetic => '=IF(IV1,IV3/IV2,0)',
                            rows       => 0,
                            cols       => $costToSml->{cols},
                            arguments  => {
                                IV1 => $costToSml,
                                IV2 => $costToSml,
                                IV3 => Dataset(
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
                        name => 'Factor to scale to £1/kW '
                          . '(operating) at each level',
                        arithmetic => '=IF(IV1,1/IV2,0)',
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
                            IV1 => $operatingCostToSml,
                            IV2 => $operatingCostToSml,
                        }
                      );
                }
                else {
                    push @factors, $model->{scaler} =~ /ehv/
                      ? Arithmetic(
                        name =>
                          'Factor to scale to £1/kW (assets) at each level',
                        arithmetic => '=IF(IV1,IV3/IV2,0)',
                        cols       => $modelCostToSml->{rows},
                        rows       => 0,
                        arguments  => {
                            IV1 => $modelCostToSml,
                            IV2 => $modelCostToSml,
                            IV3 => Constant(
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
                        name =>
                          'Factor to scale to £1/kW (assets) at each level',
                        arithmetic => '=IF(IV1,1/IV2,0)',
                        cols       => $modelCostToSml->{rows},
                        rows       => 0,
                        arguments =>
                          { IV1 => $modelCostToSml, IV2 => $modelCostToSml, }
                      );
                }
                $assetFlag = Stack(
                    name    => 'Applicability factor for £1/kW (assets) scaler',
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
                    arithmetic    => '=IV1*IV4',
                    arguments     => {
                        IV1 => Dataset(
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
                        IV4 => $totalRevenuesSoFar,
                    }
                );

                $revenueShortfall = Arithmetic(
                    name          => 'Revenue to be recovered by scaler',
                    defaultFormat => '0softnz',
                    arithmetic    => '=MIN(IV1,IV2)',
                    arguments     => {
                        IV1 => $revenueShortfall,
                        IV2 => $max,
                    }
                );

                Columnset(
                    name    => 'Application of scaler cap',
                    columns => [ $max, $revenueShortfall ]
                );

            }

            # now get on with madcapping

            my @slope = map {
                Arithmetic(
                    name       => "Effect through $_",
                    arithmetic => /day/
                    ? '=IV4*IV2*IV1/100'
                    : '=IF(IV3<0,0,IV4*IV1*10)',    # IV4*
                    arguments => {
                        IV3 => $loadCoefficients,
                        IV2 => $daysInYear,
                        IV1 => $volumeData->{$_},
                        IV4 => $assetElements->{$_},
                    }
                );
            } @$scaledComponents;

            my $slopeSet = Columnset(
                name    => 'Marginal revenue effect of scaler',
                columns => \@slope
            );

            my %minScaler = map {
                my $tariffComponent = $_;
                $_ => Arithmetic(
                    name       => "Scaler threshold for $_",
                    arithmetic => '=IF(IV3,0-IV1/IV2,0)',
                    arguments  => {
                        IV1 => $tariffsExMatching->{$_},
                        IV2 => $assetElements->{$_},
                        IV3 => $assetElements->{$_}
                    },
                    rowFormats => [
                        map {
                            $componentMap->{$_}{$tariffComponent}
                              ? undef
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
                    $_ => Arithmetic(
                        name       => "$_ scaler",
                        cols       => $assetScalerPot,
                        arithmetic => $dontScaleGeneration
                        ? '=IF(IV4<0,0,IF(IV1*IV3+IV9>0,IV81*IV83,0-IV89))'
                        : '=IF(IV1*IV3+IV9>0,IV81*IV83,0-IV89)',
                        arguments => {
                            IV1  => $assetElements->{$_},
                            IV3  => $scalerRate,
                            IV4  => $loadCoefficients,
                            IV9  => $tariffsExMatching->{$_},
                            IV81 => $assetElements->{$_},
                            IV83 => $scalerRate,
                            IV89 => $tariffsExMatching->{$_},
                        }
                    );
                } @$scaledComponents
            };

            if ($totalSiteSpecificReplacement) {
                $siteSpecificCharges = Arithmetic(
                    name => 'Total site specific sole use asset charges (£)',
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
                    arithmetic => '=IV2*(1+IV3)+IV1',
                    arguments  => {
                        IV1 => $siteSpecificOperatingCost,
                        IV2 => $siteSpecificReplacement,
                        IV3 => $scalerRate,
                    }
                );
            }

        }

        else {

            my $revenuesFromAsset;

            {
                my @termsNoDays;
                my @termsWithDays;
                my %args = ( IV400 => $daysInYear );
                my $i = 1;
                foreach (@$nonExcludedComponents) {
                    ++$i;
                    my $pad = "$i";
                    $pad = "0$pad" while length $pad < 3;
                    if (m#/day#) {
                        push @termsWithDays, "IV2$pad*IV3$pad";
                    }
                    else {
                        push @termsNoDays, "IV2$pad*IV3$pad";
                    }
                    $args{"IV2$pad"} = $assetElements->{$_};
                    $args{"IV3$pad"} = $volumeData->{$_};
                }
                $revenuesFromAsset = Arithmetic(
                    name       => 'Net revenues to which the scaler applies',
                    rows       => $scalableTariffs,
                    arithmetic => '='
                      . join(
                        '+',
                        @termsWithDays
                        ? ( '0.01*IV400*(' . join( '+', @termsWithDays ) . ')' )
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
                arithmetic => '=IV1+IV2',
                arguments  => {
                    IV1 => $totalRevenuesFromScalable,
                    IV2 => $totalSiteSpecificReplacement
                },
                defaultFormat => '0soft'
            ) if $totalSiteSpecificReplacement;

            push @{ $model->{assetScaler} }, $totalRevenuesFromScalable;

            if ($capped) {

                my $max = Arithmetic(
                    name => 'Maximum revenue that can be recovered by scaler',
                    defaultFormat => '0softnz',
                    arithmetic    => '=IV1*IV4',
                    arguments     => {
                        IV1 => Dataset(
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
                        IV4 => $totalRevenuesFromScalable,
                    }
                );

                $capped = Arithmetic(
                    name          => 'Revenue to be recovered by scaler',
                    defaultFormat => '0softnz',
                    arithmetic    => '=MIN(IV1,IV2)',
                    arguments     => {
                        IV1 => $revenueShortfall,
                        IV2 => $max,
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
                        ? '=IF(IV4<0,0,IV1/IV2*IV3)'
                        : '=IV1/IV2*IV3',
                        arguments => {
                            IV1 => $assetElements->{$_},
                            IV2 => $totalRevenuesFromScalable,
                            IV3 => $capped || $revenueShortfall,
                            IV4 => $loadCoefficients
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
                    arithmetic => '=IV2*(1+IV3/IV4)+IV1',
                    arguments  => {
                        IV1 => $siteSpecificOperatingCost,
                        IV2 => $siteSpecificReplacement,
                        IV3 => $revenueShortfall,
                        IV4 => $totalRevenuesFromScalable
                    }
                );
            }

        }    # not madcapping

        my $revenuesFromScaler;

        {
            my @termsNoDays;
            my @termsWithDays;
            my %args = ( IV400 => $daysInYear );
            my $i = 1;

            foreach ( grep { $scalerTable->{$_} } @$nonExcludedComponents ) {
                ++$i;
                my $pad = "$i";
                $pad = "0$pad" while length $pad < 3;
                if (m#/day#) {
                    push @termsWithDays, "IV2$pad*IV3$pad";
                }
                else {
                    push @termsNoDays, "IV2$pad*IV3$pad";
                }
                $args{"IV2$pad"} = $scalerTable->{$_};
                $args{"IV3$pad"} = $volumeData->{$_};
            }

            $revenuesFromScaler = Arithmetic(
                name       => 'Net revenues by tariff from scaler',
                rows       => $allTariffsByEndUser,
                cols       => $assetScalerPot,
                arithmetic => '='
                  . join( '+',
                    @termsWithDays
                    ? ( '0.01*IV400*(' . join( '+', @termsWithDays ) . ')' )
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
                arithmetic => '=IV2/IV3',
                arguments  => {
                    IV2 => $totalRevenuesFromScaler,
                    IV3 => $totalRevenuesFromScalable
                }
              )
              : $model->{scaler} && $model->{scaler} =~ /pick/i ? Arithmetic(
                name       => 'Scaling factor',
                arithmetic => '=IV2/IV3',
                arguments  => {
                    IV2 => $totalRevenuesFromScaler,
                    IV3 => $totalRevenuesFromScalable
                }
              )
              : Arithmetic(
                name          => 'Effective asset annuity rate',
                defaultFormat => '%softnz',
                arithmetic    => '=IV8*(1+IV2/IV3)',
                arguments     => {
                    IV8 => $annuityRate,
                    IV2 => $totalRevenuesFromScaler,
                    IV3 => $totalRevenuesFromScalable
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
                        arithmetic => '=0-IV1',
                        arguments  => {
                            IV1 => Arithmetic(
                                name       => "$_ ATW scaler",
                                rows       => $atwTariffs,
                                arithmetic => '=IV1',
                                arguments  => { IV1 => $scalerTable->{$_} }
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
                        arithmetic => '=0-IV1',
                        arguments  => {
                            IV1 => SumProduct(
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
                    map { ${ $_->{arguments} }{IV1} }
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
                my %args = ( IV400 => $daysInYear );
                my $i = 1;

                if ($excludeScaler) {
                    foreach ( grep { $excludeScaler->{$_} }
                        @$nonExcludedComponents )
                    {
                        ++$i;
                        my $pad = "$i";
                        $pad = "0$pad" while length $pad < 3;
                        if (m#/day#) {
                            push @termsWithDays, "IV2$pad*IV3$pad";
                        }
                        else {
                            push @termsNoDays, "IV2$pad*IV3$pad";
                        }
                        $args{"IV2$pad"} = $excludeScaler->{$_};
                        $args{"IV3$pad"} = $volumeData->{$_};
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
                            push @termsWithDays, "IV2$pad*IV3$pad";
                        }
                        else {
                            push @termsNoDays, "IV2$pad*IV3$pad";
                        }
                        $args{"IV2$pad"} = $excludeReplacement->{$_};
                        $args{"IV3$pad"} = $volumeData->{$_};
                    }
                }

                $revenuesRemoved = Arithmetic(
                    name       => 'Lost revenues by tariff from deductions',
                    rows       => $tariffsByAtw,
                    arithmetic => '=0-'
                      . join(
                        '-',
                        @termsWithDays
                        ? ( '0.01*IV400*(' . join( '+', @termsWithDays ) . ')' )
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

            push @{ $model->{adderResults} }, $totalRevenueRemoved = GroupBy(
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
                arithmetic    => '=IV1-IV2',
                arguments =>
                  { IV1 => $revenueShortfall, IV2 => $totalRevenuesFromScaler }
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
                    arithmetic    => '=IV1+IV2',
                    arguments =>
                      { IV1 => $revenueShortfall, IV2 => $totalRevenueRemoved }
                );
            }
            else { $adderAmount = $totalRevenueRemoved; }
        }

        if ( $model->{scaler} =~ /ppugeneral/i ) {

            my @columns = grep { /kWh/ } @$nonExcludedComponents;
            my @slope = map {
                Arithmetic(
                    name       => "Effect through $_",
                    arithmetic => '=IF(IV3<0,0,IV1*10)',
                    arguments  => {
                        IV3 => $loadCoefficients,
                        IV2 => $daysInYear,
                        IV1 => $volumeData->{$_},
                    }
                );
            } @columns;

            my $slopeSet = Columnset(
                name    => 'Marginal revenue effect of adder',
                columns => \@slope
            );

            my ( %minAdder, $minAdderSet, %maxAdder, $maxAdderSet );

            if ( $model->{scaler} =~ /min/i ) {
                my %min = map {
                    my $tariffComponent = $_;
                    $_ => $model->{scaler} =~ /zero/i
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
                      : Dataset(
                        name       => "Minimum $_",
                        rows       => $allTariffsByEndUser,
                        validation => {
                            validate => 'decimal',
                            criteria => 'between',
                            minimum  => -999_999.999,
                            maximum  => 999_999.999,
                        },
                        data => [ map { 0 } @{ $allTariffsByEndUser->{list} } ],
                        rowFormats => [
                            map {
                                $componentMap->{$_}{$tariffComponent} ? undef
                                  :   'unavailable';
                            } @{ $allTariffsByEndUser->{list} }
                        ]
                      );
                } @columns;
                Columnset(
                    name    => 'Minimum rates',
                    columns => [ @min{@columns} ]
                );
                %minAdder = map {
                    my $tariffComponent = $_;
                    $_ => Arithmetic(
                        name       => "Adder threshold for $_",
                        arithmetic => '=IF(ISNUMBER(IV3),IV2-IV1,0)',
                        arguments  => {
                            IV1 => $tariffsExMatching->{$_},
                            IV2 => $min{$_},
                            IV3 => $min{$_}
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
                        data => [
                            map { 999_999.999 }
                              @{ $allTariffsByEndUser->{list} }
                        ],
                        rowFormats => [
                            map {
                                $componentMap->{$_}{$tariffComponent}
                                  ? undef
                                  : 'unavailable';
                            } @{ $allTariffsByEndUser->{list} }
                        ]
                    );
                } @columns;
                Columnset(
                    name    => 'Maximum rates',
                    columns => [ @max{@columns} ]
                );
                %maxAdder = map {
                    my $tariffComponent = $_;
                    $_ => Arithmetic(
                        name       => "Adder threshold for $_",
                        arithmetic => '=IF(ISNUMBER(IV3),IV2-IV1,0)',
                        arguments  => {
                            IV1 => $tariffsExMatching->{$_},
                            IV2 => $max{$_},
                            IV3 => $max{$_}
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
                    my $iv              = 'IV1';
                    $iv = "MAX($iv,IV5)" if $minAdder{$_};
                    $iv = "MIN($iv,IV6)" if $maxAdder{$_};
                    $_ => Arithmetic(
                        name       => "Adder on $_",
                        rows       => $allTariffsByEndUser,
                        cols       => Labelset( list => ['Adder'] ),
                        arithmetic => "=IF(IV3<0,0,$iv)",
                        arguments  => {
                            IV1 => $adderRate,
                            IV3 => $loadCoefficients,
                            $minAdder{$_} ? ( IV5 => $minAdder{$_} ) : (),
                            $maxAdder{$_} ? ( IV6 => $maxAdder{$_} ) : ()
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

        else {    # not ppugeneral

            my $includeGenerators =
              Arithmetic
              name       => 'Are generator tariffs subject to adder?',
              arithmetic => !$model->{genAdder} ? 'FALSE'
              : $model->{genAdder} =~ /charge/ ? '=IV1>0'
              : $model->{genAdder} =~ /pay/    ? '=IV1<0'
              : $model->{genAdder} =~ /always/ ? 'TRUE'
              : $model->{genAdder} =~ /never/  ? 'FALSE'
              : 'FALSE',
              arguments => { IV1 => $adderAmount };

            push @{ $model->{optionLines} }, 'Adder for generators: '
              . (
                 !$model->{genAdder} ? 'never'
                : $model->{genAdder} =~ /charge/ ? 'if it is a charge'
                : $model->{genAdder} =~ /pay/    ? 'if it is a credit'
                : $model->{genAdder} =~ /always/ ? 'always'
                : $model->{genAdder} =~ /never/  ? 'never'
                : 'never'
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
                        arithmetic => '=IV4*IF(OR(IV2,IV3>=0),ABS(IV1),0)',
                        arguments  => {
                            IV1 => $unitsInYear,
                            IV2 => $includeGenerators,
                            IV3 => $loadCoefficients,
                            IV4 => $numberOfLevels
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
                        arithmetic => '=IF(OR(IV2,IV3>=0),ABS(IV1),0)',
                        arguments  => {
                            IV1 => $unitsInYear,
                            IV2 => $includeGenerators,
                            IV3 => $loadCoefficients
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
                        arithmetic => '=IF(OR(IV2,IV3>=0),ABS(IV1),0)',
                        arguments  => {
                            IV1 => $simultaneousMaximumLoadUnits,
                            IV2 => $includeGenerators,
                            IV3 => $loadCoefficients
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
                            arithmetic => '=IF(OR(IV2,IV3>=0),'
                              . 'ABS(IV1),0)*IV4*1000/(24*IV6)*IV5',
                            arguments => {
                                IV1 => $unitRateSystemLoadCoefficients[$r],
                                IV2 => $includeGenerators,
                                IV3 => $unitRateSystemLoadCoefficients[$r],
                                IV4 => $volumeData->{"Unit rate $r1 p/kWh"},
                                IV5 => $lineLossFactorsToGsp,
                                IV6 => $daysInYear,
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
                      arithmetic => '=IF(OR(IV2,IV3>=0),ABS(IV1),0)',
                      arguments  => {
                        IV1 => $simultaneousMaximumLoadCapacity,
                        IV2 => $includeGenerators,
                        IV3 => $simultaneousMaximumLoadCapacity
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
                arithmetic => '=' . join( '+', map { "IV$_" } 1 .. @sml ),
                arguments  => { map { ; "IV$_" => $sml[ $_ - 1 ]; } 1 .. @sml }
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
                  '=IF(OR(IV5,IV2>=0),ABS(IV1)*IV7*IV4/IV3/(24*IV6)*100,0)',
                arguments => {
                    IV1 => $loadCoefficients,
                    IV2 => $loadCoefficients,
                    IV3 => $volumeForAdder,
                    IV4 => $adderAmount,
                    IV5 => $includeGenerators,
                    IV6 => $daysInYear,
                    IV7 => $lineLossFactorsToGsp
                }
              )
              : Arithmetic(
                name       => 'Adder yardstick p/kWh',
                rows       => $allTariffsByEndUser,
                cols       => $fixedAdderPot,
                arithmetic => '=IF(OR(IV5,IV2>=0),IV4/IV3/10,0)'
                  . ( $numberOfLevels ? '*IV6' : '' ),
                arguments => {
                    IV2 => $loadCoefficients,
                    IV3 => $volumeForAdder,
                    IV4 => $adderAmount,
                    IV5 => $includeGenerators,
                    $numberOfLevels ? ( IV6 => $numberOfLevels ) : ()
                }
              );

            push @{ $model->{adderResults} }, $adderUnitYardstick;

            if ( $model->{scaler} =~ /ppupercent/i ) {

                $adderTable = {
                    map {
                        $_ => Arithmetic(
                            name       => "$_ adder",
                            cols       => $fixedAdderPot,
                            arithmetic => '=IF(IV94,IV1/IV2*IV3/IV4*IV6,'
                              . ( /kWh/ ? 'IV5)' : '0)' ),
                            arguments => {
                                IV1  => $sml[0]->{source},
                                IV2  => $volumeForAdder,
                                IV3  => $revenueShortfall,
                                IV4  => $revenuesSoFar,
                                IV94 => $revenuesSoFar,
                                IV5  => $adderUnitYardstick,
                                IV6  => $tariffsExMatching->{$_}
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
                                    arithmetic => '=IF(OR(IV5,IV2>=0),'
                                      . 'ABS(IV1)*IV7*IV4/IV3/(24*IV6)*100'
                                      . ',0)',
                                    arguments => {
                                        IV1 =>
                                          $unitRateSystemLoadCoefficients[$r],
                                        IV2 =>
                                          $unitRateSystemLoadCoefficients[$r],
                                        IV3 => $volumeForAdder,
                                        IV4 => $adderAmount,
                                        IV5 => $includeGenerators,
                                        IV6 => $daysInYear,
                                        IV7 => $lineLossFactorsToGsp
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
                        arithmetic => '=IF(IV5,ABS(IV1)*IV7*IV4/IV3/IV6*100,0)',
                        arguments  => {
                            IV1 => $fFactors,
                            IV2 => $fFactors,
                            IV3 => $volumeForAdder,
                            IV4 => $adderAmount,
                            IV5 => $includeGenerators,
                            IV6 => $daysInYear,
                            IV7 => $lineLossFactorsToGsp
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
                arithmetic => '=IV1/IV2/10',
                arguments  => { IV1 => $adderAmount, IV2 => $volumeForAdder }
              )
              : Arithmetic(
                name          => 'Adder rate £/kW/year',
                defaultFormat => '0.00soft',
                arithmetic    => '=IV1/IV2',
                arguments     => { IV1 => $adderAmount, IV2 => $volumeForAdder }
              );

        }    # end of if ppuminmax

        my $revenuesFromAdder;

        {
            my @termsNoDays;
            my @termsWithDays;
            my %args = ( IV400 => $daysInYear );
            my $i = 1;

            foreach ( grep { $adderTable->{$_} } @$nonExcludedComponents ) {
                ++$i;
                my $pad = "$i";
                $pad = "0$pad" while length $pad < 3;
                if (m#/day#) {
                    push @termsWithDays, "IV2$pad*IV3$pad";
                }
                else {
                    push @termsNoDays, "IV2$pad*IV3$pad";
                }
                $args{"IV2$pad"} = $adderTable->{$_};
                $args{"IV3$pad"} = $volumeData->{$_};
            }

            $revenuesFromAdder = Arithmetic(
                name       => 'Net revenues by tariff from adder',
                rows       => $allTariffsByEndUser,
                arithmetic => '='
                  . join( '+',
                    @termsWithDays
                    ? ( '0.01*IV400*(' . join( '+', @termsWithDays ) . ')' )
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
            arithmetic    => '=IV1-IV2',
            arguments =>
              { IV1 => $totalRevenuesFromAdder, IV2 => $totalRevenueRemoved }
        )
        : $totalRevenuesFromAdder
      ),
      $siteSpecificCharges, grep { $_ } $scalerTable, $excludeScaler,
      $excludeReplacement,  $adderTable;

}

sub matching2012 {

    my (
        $model,                  $adderAmount,
        $componentMap,           $allTariffsByEndUser,
        $demandTariffsByEndUser, $allEndUsers,
        $chargingLevels,         $nonExcludedComponents,
        $allComponents,          $daysAfter,
        $volumeAfter,            $loadCoefficients,
        $tariffsExMatching,      $daysFullYear,
        $volumeFullYear
    ) = @_;

    my $fixedAdderPot = Labelset( list => ['Adder'] );

    my $totalRevenueFromUnits = GroupBy(
        name          => 'Total revenues from demand unit rates (£/year)',
        defaultFormat => '0soft',
        source        => Arithmetic(
            name => Label(
                'From unit rates',
                'Revenues from demand unit rates before matching (£/year)'
            ),
            rows => $allTariffsByEndUser,
            arithmetic =>    # '=IF(IV5<0,0,'
              '=10*('
              . join( '+',
                'IV701*IV702',
                map { "IV$_*IV90$_" } 1 .. $model->{maxUnitRates} )
              . ')',         # . '))',
            arguments => {
                IV5   => $loadCoefficients,
                IV701 => $tariffsExMatching->{'Reactive power charge p/kVArh'},
                IV702 => $volumeFullYear->{'Reactive power charge p/kVArh'},
                map {
                    my $name = "Unit rate $_ p/kWh";
                    "IV$_"     => $tariffsExMatching->{$name},
                      "IV90$_" => $volumeFullYear->{$name};
                } 1 .. $model->{maxUnitRates}
            },
            defaultFormat => '0soft'
        ),
    );

    my $totalRevenueFromFixed = GroupBy(
        name          => 'Total revenues from demand fixed charges (£/year)',
        defaultFormat => '0soft',
        source        => Arithmetic(
            name => Label(
                'From fixed charges',
                'Revenues from demand fixed charges before matching (£/year)'
            ),
            rows       => $allTariffsByEndUser,
            arithmetic => '=0.01*IV6*IV1*IV2', # '=IF(IV5<0,0,0.01*IV6*IV1*IV2)'
            arguments  => {
                IV5 => $loadCoefficients,
                IV6 => $daysFullYear,
                IV1 => $tariffsExMatching->{'Fixed charge p/MPAN/day'},
                IV2 => $volumeFullYear->{'Fixed charge p/MPAN/day'},
            },
            defaultFormat => '0soft'
        ),
    );

    my $totalRevenueFromCapacity = GroupBy(
        name          => 'Total revenues from demand capacity charges (£/year)',
        defaultFormat => '0soft',
        source        => Arithmetic(
            name => Label(
                'From capacity charges',
                'Revenues from demand capacity charges before matching (£/year)'
            ),
            rows       => $allTariffsByEndUser,
            arithmetic => '=0.01*IV6*IV1*IV2', # '=IF(IV5<0,0,0.01*IV6*IV1*IV2)'
            arguments  => {
                IV5 => $loadCoefficients,
                IV6 => $daysFullYear,
                IV1 => $tariffsExMatching->{'Capacity charge p/kVA/day'},
                IV2 => $volumeFullYear->{'Capacity charge p/kVA/day'},
            },
            defaultFormat => '0soft'
        ),
    );

    Columnset(
        name    => 'Analysis of annual revenue before matching (£/year)',
        columns => [
            map { $_->{source} } $totalRevenueFromUnits,
            $totalRevenueFromFixed,
            $totalRevenueFromCapacity,
        ]
    );

    Columnset(
        name    => 'Total analsyis of annual revenue before matching  (£/year)',
        columns => [
            $totalRevenueFromUnits, $totalRevenueFromFixed,
            $totalRevenueFromCapacity,
        ]
    );

    my $adderTable = {};

    {    # units

        my @columns = grep { /kWh/ } @$nonExcludedComponents;

        my @slope = map {
            Arithmetic(
                name       => "Effect through $_",
                arithmetic => '=IF(IV3<0,0,IV1*10)',
                arguments  => {
                    IV3 => $loadCoefficients,
                    IV2 => $daysAfter,
                    IV1 => $volumeAfter->{$_},
                }
            );
        } @columns;

        my $slopeSet = Columnset(
            name    => 'Marginal revenue effect of adder',
            columns => \@slope,
        );

        my %minAdder = map {
            my $tariffComponent = $_;
            $_ => Arithmetic(
                name       => "Adder threshold for $_",
                arithmetic => '=IF(IV4<0,0,0-IV1)',
                arguments  => {
                    IV4 => $loadCoefficients,
                    IV1 => $tariffsExMatching->{$_},
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

        my $minAdderSet = Columnset(
            name    => 'Adder value at which the minimum is breached',
            columns => [ @minAdder{@columns} ]
        );

        my $adderRate = new SpreadsheetModel::SegmentRoot(
            name   => 'General adder rate (p/kWh)',
            slopes => $slopeSet,
            target => $model->{scaler} =~ /opt3/i
            ? Arithmetic(
                name       => 'Revenue matching target from unit rates (£)',
                arithmetic => '=IV1*IV2/(IV3+IV4+IV5)',
                arguments  => {
                    IV1 => $adderAmount,
                    IV2 => $totalRevenueFromUnits,
                    IV3 => $totalRevenueFromUnits,
                    IV4 => $totalRevenueFromFixed,
                    IV5 => $totalRevenueFromCapacity,
                }
              )
            : $adderAmount,
            min => $minAdderSet,
        );

        foreach (@columns) {
            my $tariffComponent = $_;
            $adderTable->{$_} = Arithmetic(
                name       => "Adder on $_",
                rows       => $allTariffsByEndUser,
                cols       => Labelset( list => ['Adder'] ),
                arithmetic => "=IF(IV3<0,0,MAX(IV6,IV1))",
                arguments  => {
                    IV1 => $adderRate,
                    IV3 => $loadCoefficients,
                    IV6 => $minAdder{$_},
                },
                rowFormats => [
                    map {
                        $componentMap->{$_}{$tariffComponent}
                          ? undef
                          : 'unavailable';
                    } @{ $allTariffsByEndUser->{list} }
                ]
            );
        }

    }

    if ( $model->{scaler} =~ /opt3/i ) {

        {    # fixed

            my @columns = grep { /fixed/i } @$nonExcludedComponents;

            my @slope = map {
                Arithmetic(
                    name       => "Effect through $_",
                    arithmetic => '=IF(IV3>0,IV1*IV2*0.01,0)',
                    arguments  => {
                        IV3 => $loadCoefficients,
                        IV2 => $daysAfter,
                        IV1 => $volumeAfter->{$_},
                    }
                );
            } @columns;

            my $slopeSet = Columnset(
                name    => 'Marginal revenue effect of adder',
                columns => \@slope,
            );

            my %minAdder = map {
                my $tariffComponent = $_;
                $_ => Arithmetic(
                    name       => "Adder threshold for $_",
                    arithmetic => '=IF(IV4>0,0-IV1,0)',
                    arguments  => {
                        IV4 => $loadCoefficients,
                        IV1 => $tariffsExMatching->{$_},
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

            my $minAdderSet = Columnset(
                name    => 'Adder value at which the minimum is breached',
                columns => [ @minAdder{@columns} ]
            );

            my $adderRate = new SpreadsheetModel::SegmentRoot(
                name   => 'General adder rate (p/MPAN/day)',
                slopes => $slopeSet,
                target => Arithmetic(
                    name => 'Revenue matching target from fixed charges (£)',
                    arithmetic => '=IV1*IV2/(IV3+IV4+IV5)',
                    arguments  => {
                        IV1 => $adderAmount,
                        IV2 => $totalRevenueFromFixed,
                        IV3 => $totalRevenueFromUnits,
                        IV4 => $totalRevenueFromFixed,
                        IV5 => $totalRevenueFromCapacity,
                    }
                ),
                min => $minAdderSet,
            );

            foreach (@columns) {
                my $tariffComponent = $_;
                $adderTable->{$_} = Arithmetic(
                    name       => "Adder on $_",
                    rows       => $allTariffsByEndUser,
                    cols       => Labelset( list => ['Adder'] ),
                    arithmetic => "=IF(IV3<0,0,MAX(IV6,IV1))",
                    arguments  => {
                        IV1 => $adderRate,
                        IV3 => $loadCoefficients,
                        IV6 => $minAdder{$_},
                    },
                    rowFormats => [
                        map {
                            $componentMap->{$_}{$tariffComponent}
                              ? undef
                              : 'unavailable';
                        } @{ $allTariffsByEndUser->{list} }
                    ]
                );
            }

        }

        {    # capacity

            my @columns = grep { /kVA\/day/ } @$nonExcludedComponents;

            my @slope = map {
                Arithmetic(
                    name       => "Effect through $_",
                    arithmetic => '=IF(IV3<0,0,IV1*IV2*0.01)',
                    arguments  => {
                        IV3 => $loadCoefficients,
                        IV2 => $daysAfter,
                        IV1 => $volumeAfter->{$_},
                    }
                );
            } @columns;

            my $slopeSet = Columnset(
                name    => 'Marginal revenue effect of adder',
                columns => \@slope,
            );

            my %minAdder = map {
                my $tariffComponent = $_;
                $_ => Arithmetic(
                    name       => "Adder threshold for $_",
                    arithmetic => '=IF(IV4<0,0,0-IV1)',
                    arguments  => {
                        IV4 => $loadCoefficients,
                        IV1 => $tariffsExMatching->{$_},
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

            my $minAdderSet = Columnset(
                name    => 'Adder value at which the minimum is breached',
                columns => [ @minAdder{@columns} ]
            );

            my $adderRate = new SpreadsheetModel::SegmentRoot(
                name   => 'General adder rate (p/kVA/day)',
                slopes => $slopeSet,
                target => Arithmetic(
                    name => 'Revenue matching target from capacity charges (£)',
                    arithmetic => '=IV1*IV2/(IV3+IV4+IV5)',
                    arguments  => {
                        IV1 => $adderAmount,
                        IV2 => $totalRevenueFromCapacity,
                        IV3 => $totalRevenueFromUnits,
                        IV4 => $totalRevenueFromFixed,
                        IV5 => $totalRevenueFromCapacity,
                    }
                ),
                min => $minAdderSet,
            );

            foreach (@columns) {
                my $tariffComponent = $_;
                $adderTable->{$_} = Arithmetic(
                    name       => "Adder on $_",
                    rows       => $allTariffsByEndUser,
                    cols       => Labelset( list => ['Adder'] ),
                    arithmetic => "=IF(IV3<0,0,MAX(IV6,IV1))",
                    arguments  => {
                        IV1 => $adderRate,
                        IV3 => $loadCoefficients,
                        IV6 => $minAdder{$_},
                    },
                    rowFormats => [
                        map {
                            $componentMap->{$_}{$tariffComponent}
                              ? undef
                              : 'unavailable';
                        } @{ $allTariffsByEndUser->{list} }
                    ]
                );
            }

        }

    }    # end of option 3-only SegmentRoots

    my $revenuesFromAdder;

    {
        my @termsNoDays;
        my @termsWithDays;
        my %args = ( IV400 => $daysAfter );
        my $i = 1;

        foreach ( grep { $adderTable->{$_} } @$nonExcludedComponents ) {
            ++$i;
            my $pad = "$i";
            $pad = "0$pad" while length $pad < 3;
            if (m#/day#) {
                push @termsWithDays, "IV2$pad*IV3$pad";
            }
            else {
                push @termsNoDays, "IV2$pad*IV3$pad";
            }
            $args{"IV2$pad"} = $adderTable->{$_};
            $args{"IV3$pad"} = $volumeAfter->{$_};
        }

        $revenuesFromAdder = Arithmetic(
            name       => 'Net revenues by tariff from adder',
            rows       => $allTariffsByEndUser,
            arithmetic => '='
              . join( '+',
                @termsWithDays
                ? ( '0.01*IV400*(' . join( '+', @termsWithDays ) . ')' )
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

    }

    my $totalRevenuesFromAdder = GroupBy(
        name          => 'Total net revenues from adder (£)',
        rows          => 0,
        cols          => 0,
        source        => $revenuesFromAdder,
        defaultFormat => '0softnz'
    );

    my $siteSpecificCharges;    # not implemented in this option

    $totalRevenuesFromAdder, $siteSpecificCharges, $adderTable;

}

1;
