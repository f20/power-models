package CDCM;

# Copyright 2009-2011 Energy Networks Association Limited and others.
# Copyright 2011-2025 Franck Latrémolière and others.
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

sub timeOfDay {

    my ( $model, $networkLevels, $componentMap, $allEndUsers, $daysInYear,
        $loadCoefficients, $volumeByEndUser, $unitsByEndUser )
      = @_;

    my $timebandSet;

    if ( $model->{customTimebandNames} || $model->{timebands} != 3 ) {
        $timebandSet = Labelset(
            name     => 'Distribution time bands',
            editable => Dataset(
                cols => Labelset(
                    name => 'Distribution time band list',
                    list =>
                      [ map { "Name of band $_" } 1 .. $model->{timebands} ]
                ),
                name          => 'Names of distribution time bands',
                number        => 1050,
                appendTo      => $model->{inputTables},
                dataset       => $model->{dataset},
                defaultFormat => 'puretexthard',
                data          => [ map { "Band $_" } 1 .. $model->{timebands} ],
                lines         =>
                  'Source: user preference (does not affect calculations).'
            )
        );
    }
    else {
        $timebandSet = Labelset(
            name => 'Distribution time bands',
            list => [qw(Red Amber Green)],
        );
    }

    my $annualHoursByTimebandRaw = Dataset(
        name       => 'Typical annual hours by distribution time band',
        validation => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
        lines => [
            'Source: definition of distribution time bands.',
            'The figures in this table will be automatically'
              . ' adjusted to match the number of days in the charging period.',
        ],
        singleRowName => 'Annual hours',
        number        => 1068,
        appendTo      => $model->{inputTables},
        dataset       => $model->{dataset},
        cols          => $timebandSet,
        data => [ qw(650 3224 4862), map { 0 } 4 .. $model->{timebands} ],
        defaultFormat => '0.0hardnz'
    );

    my $annualHoursByTimebandTotal = GroupBy(
        rows => 0,
        cols => 0,
        name => Label(
            'Hours aggregate',
            'Total hours in the year'
              . ' according to time band hours input data'
        ),
        defaultFormat => '0.0softnz',
        source        => $annualHoursByTimebandRaw,
    );

    my $annualHoursByTimeband = Arithmetic(
        name =>
          'Annual hours by distribution time band (reconciled to days in year)',
        singleRowName => 'Annual hours',
        defaultFormat => '0.0softnz',
        arithmetic    => '=A1*24*A3/A2',
        arguments     => {
            A2 => $annualHoursByTimebandTotal,
            A3 => $daysInYear,
            A1 => $annualHoursByTimebandRaw,
        }
    );

    $model->{hoursByRedAmberGreen} = $annualHoursByTimeband;

    Columnset(
        name =>
          'Adjust annual hours by distribution time band to match days in year',
        columns => [ $annualHoursByTimebandTotal, $annualHoursByTimeband ]
    );

    my $networkLevelsTimeband = Labelset(
        name   => 'Network levels and time bands',
        groups => [
            map {
                my $lev = "$_";
                Labelset(
                    name => $lev,
                    list => $timebandSet->{list}
                )
            } @{ $networkLevels->{list} }
        ]
    );

    my $networkLevelsTimebandAware = Labelset
      name    => 'Network levels aware of time band',
      list    => $networkLevelsTimeband->{groups},
      accepts => [$networkLevels];

    my $peakingProbabilitiesTable = Dataset(
        name       => 'Peaking probabilities by network level',
        validation => {
            validate => 'decimal',
            criteria => 'between',
            minimum  => 0,
            maximum  => 1,
        },
        lines    => 'Source: analysis of network operation data.',
        number   => 1069,
        appendTo => $model->{inputTables},
        dataset  => $model->{dataset},
        cols     => $timebandSet,
        rows     => $networkLevelsTimebandAware,
        byrow    => 1,
        data     => [
            map {
                /(132|GSP|Trans)/i
                  ? (
                    @{ $timebandSet->{list} } > 2
                    ? [ .95, .05, map { 0 } 3 .. @{ $timebandSet->{list} } ]
                    : [ 1, 0 ]
                  )
                  : /^EHV$/i ? (
                    @{ $timebandSet->{list} } > 2
                    ? [
                        .69, .29,
                        .02, map { 0 } 4 .. @{ $timebandSet->{list} }
                      ]
                    : [ .98, .02 ]
                  )
                  : (
                    @{ $timebandSet->{list} } > 2
                    ? [
                        .52, .39,
                        .09, map { 0 } 4 .. @{ $timebandSet->{list} }
                      ]
                    : [ .84, .16 ]
                  )
            } @{ $networkLevelsTimebandAware->{list} }
        ],
        defaultFormat => '%hard'
    );

    my $totalProbability = GroupBy(
        name          => 'Total probability (should be 100%)',
        rows          => $networkLevelsTimebandAware,
        defaultFormat => '%soft',
        source        => $peakingProbabilitiesTable
    );

    $peakingProbabilitiesTable = Arithmetic(
        name => (
                 $model->{otneiErrors}
              || $model->{lvDiversityWrong} ? 'Non-normalised ' : 'Normalised '
          )
          . 'peaking probabilities',
        defaultFormat => '%soft',
        arithmetic    => $model->{otneiErrors} || $model->{lvDiversityWrong}
        ? '=IF(A3,A1,A8/A9)'
        : '=IF(A3,A1/A2,A8/A9)',
        arguments => {
            A8 => $annualHoursByTimebandRaw,
            A9 => $annualHoursByTimebandTotal,
            A1 => $peakingProbabilitiesTable,
            A2 => $totalProbability,
            A3 => $totalProbability,
        }
    );

    Columnset(
        name    => 'Normalisation of peaking probabilities',
        columns => [ $totalProbability, $peakingProbabilitiesTable ]
    );

    my $peakingProbability = new SpreadsheetModel::Reshape(
        name          => 'Peaking probabilities by network level (reshaped)',
        singleRowName => 'Probability of peak within timeband',
        cols          => $networkLevelsTimeband,
        rows          => 0,
        defaultFormat => '%copy',
        source        => $peakingProbabilitiesTable
    );

    my @relevantEndUsersByRate;
    my @relevantTariffsByRate;

    my $usersWithDistTimeBands = Labelset(
        name => 'Users with distribution time band tariff',
        list => [
            grep { $componentMap->{$_}{'Unit rates p/kWh'} }
              @{ $allEndUsers->{list} }
        ]
    );

    my $tariffsWithDistTimeBands =
      $model->{pcd} ? $usersWithDistTimeBands : Labelset(
        name   => 'Time band tariffs',
        groups => $usersWithDistTimeBands->{list}
      );

    my $relevantEndUsers = Labelset(
        name => 'End users for multiple unit rate calculation',
        list => [
            grep {
                     $componentMap->{$_}{'Unit rate 2 p/kWh'}
                  || $componentMap->{$_}{'Unit rate 0 p/kWh'}
            } @{ $allEndUsers->{list} }
        ]
    );

    my $relevantTariffs = $model->{pcd} ? $relevantEndUsers : Labelset(
        name   => 'Tariffs for multiple unit rate calculation',
        groups => $relevantEndUsers->{list}
    );

    {
        my $prevC = -1;
        my $dtbC  = @{ $usersWithDistTimeBands->{list} };
        foreach ( 0 .. $model->{maxUnitRates} - 1 ) {
            my $r        = $_ + 1;
            my $rateDesc = "Unit rate $r p/kWh";
            my @us =
              grep { $componentMap->{$_}{$rateDesc} }
              @{ $relevantEndUsers->{list} };

            last unless @us;

            if ( @us == $prevC ) {
                $relevantEndUsersByRate[$_] = $relevantEndUsersByRate[ $_ - 1 ];
                $relevantTariffsByRate[$_]  = $relevantTariffsByRate[ $_ - 1 ];
            }
            elsif ( @us == $dtbC ) {
                $relevantEndUsersByRate[$_] = $usersWithDistTimeBands;
                $relevantTariffsByRate[$_]  = $tariffsWithDistTimeBands;
            }
            else {
                $prevC = @us;
                $relevantEndUsersByRate[$_] = Labelset(
                    name => "End users which have a unit rate $r",
                    list => \@us
                );
                $relevantTariffsByRate[$_] =
                  $model->{pcd} ? $relevantEndUsersByRate[$_] : Labelset(
                    name   => "Tariffs which have a unit rate $r",
                    groups => \@us
                  );
            }
        }
    }

    my @timebandUseByRate = map {
        my $usersWithThisRate = $relevantEndUsersByRate[$_];
        my $r                 = 1 + $_;
        my $xst =
          $r > 9
          ? "${r}th"
          : (qw(first second third fourth fifth sixth seventh eigth ninth))[$_];
        my $usersWithInput = Labelset(
            name => "Users with non-obvious split of $xst TPR",
            list => [
                grep { !$componentMap->{$_}{'Unit rates p/kWh'} }
                  @{ $usersWithThisRate->{list} }
            ]
        );

        my ( $inData, $conData );

        if ( @{ $usersWithInput->{list} } ) {

            $inData = Dataset(
                name => 'Average split of rate '
                  . $r
                  . ' units by distribution time band',
                validation => {
                    validate => 'decimal',
                    criteria => 'between',
                    minimum  => 0,
                    maximum  => 1,
                },
                $r < 8 ? ( number => 1060 + $r ) : (),
                appendTo => $model->{inputTables},
                dataset  => $model->{dataset},
                lines    => <<EOT
Source: load data analysis.
This table relates to the $xst TPR or charging period for each tariff.
EOT
                  . ( $r > 1 ? 'See notes on table 1061.' : <<'EOT'),
Each line relates to a different user type and tariff structure. For each user type and tariff structure, the
figure entered against each of the time bands used for network analysis is the proportion of the units recorded
on the relevant TPR or within the relevant charging period that would fall with each time band.
EOT
                rows  => $usersWithInput,
                cols  => $timebandSet,
                byrow => 1,
                data  => [
                    map {
                        $componentMap->{$_}{'Unit rates p/kWh'}
                          ? [ map { $_ == $r ? 1 : 0 }
                              1 .. $model->{timebands} ]
                          : $componentMap->{$_}{'Unit rate 0 p/kWh'}
                          || $r != 1 ? [qw(0 0 1 0 0 0 0 0 0)]
                          : [qw(0.15 0.54 0.31 0 0 0 0 0 0)]
                    } @{ $usersWithInput->{list} }
                ],
                defaultFormat => '%hard'
            );

            my $totals = GroupBy(
                name          => 'Total split',
                rows          => $usersWithInput,
                defaultFormat => '%soft',
                source        => $inData
            );

            $inData = Arithmetic(
                name =>
                  "Normalised split of rate $r units by distribution time band",
                defaultFormat => '%soft',
                arithmetic    => "=IF(A3,A1/A2,A8/A9/24)",
                arguments     => {
                    A8 => $annualHoursByTimeband,
                    A9 => $daysInYear,
                    A1 => $inData,
                    A2 => $totals,
                    A3 => $totals,
                }
            );

            Columnset(
                name    => "Normalisation of split of rate $r units",
                columns => [ $totals, $inData ]
            );

        }

        $conData = Constant(
            name => 'Split of rate '
              . $r
              . ' units between distribution time bands'
              . ' (default)',
            rows  => $usersWithDistTimeBands,
            cols  => $timebandSet,
            byrow => 1,
            data  => [
                map {
                    [ map { $_ == $r ? 1 : 0 } 1 .. $model->{timebands} ]
                } @{ $usersWithDistTimeBands->{list} }
            ],
            defaultFormat => '%connz'
        ) if @{ $usersWithDistTimeBands->{list} };
        $inData && $conData ? Stack(
            name => 'Split of rate '
              . $r
              . ' units between distribution time bands',
            rows          => $usersWithThisRate,
            cols          => $timebandSet,
            sources       => [ $inData, $conData ],
            defaultFormat => '%copynz'
        ) : $inData ? $inData : $conData;
    } 0 .. $model->{maxUnitRates} - 1;

    push @{ $model->{timeOfDayResults} }, @timebandUseByRate
      unless $model->{coincidenceAdj}
      && $model->{coincidenceAdj} =~ /none/i;

    my $timebandLoadCoefficient;

    if ( $model->{timebandCoef} ) {

        push @{ $model->{optionLines} }, 'Time band load coefficients used';

        if ( $model->{timebandCoef} =~ /detail/i ) {
            $timebandLoadCoefficient = Dataset(
                name     => 'Unscaled time band load coefficients',
                number   => 1067,
                appendTo => $model->{inputTables},
                dataset  => $model->{dataset},
                rows     => $relevantEndUsersByRate[0],
                cols     => $networkLevelsTimeband,
                data     => [
                    map {
                        [ map { /generat/i ? -1 : 1 }
                              @{ $relevantEndUsersByRate[0]{list} } ]
                    } @{ $networkLevelsTimeband->{list} }
                ],
                defaultFormat => '%hardnz'
            );
        }
        else {
            $timebandLoadCoefficient = Constant(
                name => 'Default time band load coefficient value',
                rows => $relevantEndUsersByRate[0],
                cols => 0,
                data => [
                    map { /generat/i ? -1 : 1 }
                      @{ $relevantEndUsersByRate[0]{list} }
                ],
                defaultFormat => '%connz'
            );
        }
    }

    # unadjusted; to be replaced below if there is a coincidence adjustment
    my $pseudoLoadCoefficientBreakdown = Arithmetic(
        name       => 'Pseudo load coefficient by time band and network level',
        rows       => $relevantEndUsersByRate[0],
        cols       => $networkLevelsTimeband,
        arithmetic => '=IF(A6>0,A7*24*A9/A5,0)*IF(A2<0,-1,1)',
        arguments  => {
            A5 => $annualHoursByTimeband,
            A6 => $annualHoursByTimeband,
            A7 => $peakingProbability,
            A9 => $daysInYear,
            A2 => $loadCoefficients,
        }
    );

    my @pseudoLoadCoefficientsAgainstSystemPeak;

    unless ( $model->{coincidenceAdj} && $model->{coincidenceAdj} =~ /none/i ) {
        my $peakBand = Labelset( list => [ $timebandSet->{list}[0] ] );

        my $timebandLoadCoefficientAccording = Stack(
            name    => 'First-time-band load coefficient',
            rows    => $relevantEndUsersByRate[0],
            cols    => $peakBand,
            sources => [
                map {

                    my $r  = $_ + 1;
                    my $r2 = $_ + 2;
                    my $rt =
                        $_ > 9
                      ? $r
                      : qw(one two three four five six seven eight nine ten)
                      [$_];
                    if (
                        my @relevant =
                        grep {
                                  !/(?:related|additional|gener)/i
                              and !$componentMap->{$_}{"Unit rate $r2 p/kWh"}
                        } @{ $relevantEndUsersByRate[$_]{list} }
                      )
                    {

                        my $relevantUsers = Labelset
                          name => "Demand end users with $rt-rate tariffs",
                          list => \@relevant;

                        my $timebandUseByRateTotal = Arithmetic(
                            rows => $relevantUsers,
                            cols => $timebandSet,
                            name => 'Use of distribution time bands by units'
                              . " in demand forecast for $rt-rate tariffs",
                            arithmetic => '=IF(A501>0,(' . join(
                                '+',
                                map {
                                    my $pad = "$_";
                                    $pad = "0$pad" while length $pad < 3;
                                    "A1$pad*A2$pad"
                                } 1 .. $r
                              )
                              . ')/A502,0)',
                            arguments => {
                                A501 => $unitsByEndUser,
                                A502 => $unitsByEndUser,
                                map {
                                    my $pad = $_;
                                    $pad = "0$pad" while length $pad < 3;
                                    (
                                        "A1$pad" => $volumeByEndUser->{
                                            "Unit rate $_ p/kWh"},
                                        "A2$pad" => $timebandUseByRate[ $_ - 1 ]
                                    )
                                } 1 .. $r
                            },
                            defaultFormat => '%softnz'
                        );

                        my $implied = Arithmetic(
                            name => 'First-time-band load coefficient for '
                              . $rt
                              . '-rate tariffs',
                            arithmetic => $timebandLoadCoefficient
                            ? '=IF(A6>0,A1*A2*A4*24/A5,0)'
                            : '=IF(A6>0,A1*A4*24/A5,0)',
                            cols      => $peakBand,
                            rows      => $relevantUsers,
                            arguments => {
                                $timebandLoadCoefficient
                                ? ( A2 => $timebandLoadCoefficient )
                                : (),
                                A1 => $timebandUseByRateTotal,
                                A4 => $daysInYear,
                                A5 => $annualHoursByTimeband,
                                A6 => $annualHoursByTimeband
                            }
                        );

                        Columnset
                          name =>
"Calculation of implied load coefficients for $rt-rate users",
                          columns => [ $timebandUseByRateTotal, $implied ];

                        $model->{noCoincidenceForHalfHourly}
                          ? View(
                            sources => [$implied],
                            rows    => Labelset(
                                list => [
                                    grep {
                                        !$componentMap->{$_}{"Unit rates p/kWh"}
                                    } @relevant
                                ]
                            )
                          )
                          : $implied;

                    }
                    else { (); }

                } 0 .. $model->{maxUnitRates} - 1
            ]
        );

        @pseudoLoadCoefficientsAgainstSystemPeak = map {
            my $r = 1 + $_;
            Arithmetic(
                name => "Unit rate $r pseudo load coefficient at system level",
                arithmetic =>
                  '=IF(A6>0,A1*IF(A7<>0,A3/A2,IF(A9<0,-1,1))*24*A4/A5,0)',
                cols      => $peakBand,
                arguments => {
                    A1 => $timebandUseByRate[$_],
                    A2 => $timebandLoadCoefficientAccording,
                    A3 => $loadCoefficients,
                    A9 => $loadCoefficients,
                    A4 => $daysInYear,
                    A5 => $annualHoursByTimeband,
                    A6 => $annualHoursByTimeband,
                    A7 => $timebandLoadCoefficientAccording
                },
                rows    => $relevantEndUsersByRate[$_],
                tariffs => $relevantTariffsByRate[$_],
            );
        } 0 .. $model->{maxUnitRates} - 1;

        my $timebandLoadCoefficientAdjusted =
          Arithmetic
          name => 'Load coefficient correction factor'
          . ' (kW at peak in band / band average kW)',
          arithmetic => $timebandLoadCoefficient ? '=IF(A5<>0,A4/A2/A1,A6)'
          : '=IF(A5<>0,A4/A2,IF(A8<0,-1,1))',
          rows => $relevantEndUsersByRate[0],
          $model->{timebandCoef} && $model->{timebandCoef} =~ /detail/i
          ? ( cols => $networkLevelsTimeband )
          : (),
          arguments => {
            $timebandLoadCoefficient
            ? (
                A1 => $timebandLoadCoefficient,
                A6 => $timebandLoadCoefficient
              )
            : (),
            A2 => $timebandLoadCoefficientAccording,
            A5 => $timebandLoadCoefficientAccording,
            A4 => $loadCoefficients,
            A8 => $loadCoefficients,
          };

        if ( $model->{coincidenceAdj} && $model->{coincidenceAdj} =~ /group/i )
        {

            my $relevantUsers =
              Labelset( list =>
                  [ grep { !/gener/i } @{ $relevantEndUsersByRate[0]{list} } ]
              );

            my ( $tariffGroupset, $mapping );

            if ( $model->{coincidenceAdj} =~ /voltage/i ) {

                push @{ $model->{optionLines} },
                  'Coincidence correction factors by'
                  . ' voltage level tariff group';

                # And now override the groupings

                $relevantUsers =
                  Labelset( list =>
                      [ grep { !/^hv sub/i } @{ $relevantUsers->{list} } ] );

                $tariffGroupset =
                  Labelset(
                    list => [ 'LV network', 'LV substation', 'HV network', ] );

                $mapping = Constant(
                    name => 'Mapping of tariffs to '
                      . 'tariff groups for coincidence adjustment factor',
                    defaultFormat => '0connz',
                    rows          => $relevantUsers,
                    cols          => $tariffGroupset,
                    data          => [
                        map {
                                /^lv sub/i ? [qw(0 1 0)]
                              : /^lv/i     ? [qw(1 0 0)]
                              :              [qw(0 0 1)];
                        } @{ $relevantUsers->{list} }
                    ],
                    byrow => 1,
                );

            }
            elsif ( $model->{coincidenceAdj} =~ /3|three/i ) {

                $tariffGroupset = Labelset(
                    list => [
                        'Domestic and/or single-phase '
                          . 'and/or non-half-hourly UMS',
                        'Non-domestic and/or three-phase whole current metered',
                        'Large and/or half-hourly',
                    ]
                );

                $mapping = Constant(
                    name => 'Mapping of tariffs to '
                      . 'tariff groups for coincidence adjustment factor',
                    defaultFormat => '0connz',
                    rows          => $relevantUsers,
                    cols          => $tariffGroupset,
                    data          => [
                        map {
                            /domestic|1p|single/i && !/non.?dom/i
                              || !$componentMap->{$_}{'Fixed charge p/MPAN/day'}
                              && !$componentMap->{$_}{'Unit rates p/kWh'}
                              ? [qw(1 0 0)]
                              : /wc|small/i ? [qw(0 1 0)]
                              :               [qw(0 0 1)];
                        } @{ $relevantUsers->{list} }
                    ],
                    byrow => 1,
                );

            }
            elsif ( $model->{coincidenceAdj} =~ /all/i ) {
                push @{ $model->{optionLines} },
                  'Single coincidence correction factor';

                $tariffGroupset = Labelset( list => [ 'All demand tariffs', ] );

                $mapping = Constant(
                    name => 'Mapping of tariffs to '
                      . 'tariff groups for coincidence adjustment factor',
                    defaultFormat => '0connz',
                    rows          => $relevantUsers,
                    cols          => $tariffGroupset,
                    data  => [ map { [1] } @{ $relevantUsers->{list} } ],
                    byrow => 1,
                );
            }

            if ($mapping) {

                my $red = Arithmetic(
                    name          => 'Contribution to first-band peak kW',
                    defaultFormat => '0softnz',
                    arithmetic    => $timebandLoadCoefficient
                    ? '=A1*A9*A2/24/A3*1000'
                    : '=A1*A2/24/A3*1000',
                    rows      => $relevantUsers,
                    arguments => {
                        A1 => $timebandLoadCoefficientAccording,
                        A2 => $unitsByEndUser,
                        A3 => $daysInYear,
                        $timebandLoadCoefficient
                        ? ( A9 => $timebandLoadCoefficient )
                        : (),
                    },
                );

                my $coin = Arithmetic(
                    name          => 'Contribution to system-peak-time kW',
                    defaultFormat => '0softnz',
                    arithmetic    => '=A1*A2/24/A3*1000',
                    rows          => $relevantUsers,
                    arguments     => {
                        A1 => $loadCoefficients,
                        A2 => $unitsByEndUser,
                        A3 => $daysInYear,
                    },
                );

                $timebandLoadCoefficientAccording->{dontcolumnset} = 1;

                Columnset(
                    name    => 'Estimated contributions to peak demand',
                    columns => [    # $timebandLoadCoefficientAccording,
                        $red, $coin,
                    ]
                );

                my $redG = SumProduct(
                    name          => 'Group contribution to first-band peak kW',
                    defaultFormat => '0softnz',
                    matrix        => $mapping,
                    vector        => $red,
                );

                my $coinG = SumProduct(
                    name => 'Group contribution to system-peak-time kW',
                    defaultFormat => '0softnz',
                    matrix        => $mapping,
                    vector        => $coin,
                );

                $timebandLoadCoefficientAdjusted = Stack(
                    name    => 'Load coefficient correction factor (combined)',
                    rows    => $relevantEndUsersByRate[0],
                    cols    => 0,
                    sources => [
                        SumProduct(
                            name => 'Load coefficient correction factor '
                              . '(based on group)',
                            matrix => $mapping,
                            vector => Arithmetic(
                                name => 'Load coefficient correction factor'
                                  . ' for each group',
                                arithmetic => '=IF(A1,A2/A3,0)',
                                rows       => 0,
                                arguments  => {
                                    A1 => $redG,
                                    A2 => $coinG,
                                    A3 => $redG,
                                }
                            ),
                        ),
                        $timebandLoadCoefficientAdjusted,
                    ]
                );

            }
        }

        push @{ $model->{timeOfDayResults} },
          Columnset(
            name    => 'Calculation of adjusted time band load coefficients',
            columns => [
                $timebandLoadCoefficientAccording,
                $timebandLoadCoefficientAdjusted
            ]
          ) unless $timebandLoadCoefficientAccording->{dontcolumnset};

        $pseudoLoadCoefficientBreakdown = Arithmetic(
            name => 'Pseudo load coefficient by time band and network level',
            rows => $relevantEndUsersByRate[0],
            cols => $networkLevelsTimeband,
            arithmetic => $model->{coincidenceAdj}
              && $model->{coincidenceAdj} =~ /redonly/i
            ? '=IF(A6>0,(1+A8*(A2-1))*A7*24*A9/A5,0)'
            : '=IF(A6>0,A2*A7*24*A9/A5,0)',
            arguments => {
                A2 => $timebandLoadCoefficientAdjusted,
                A5 => $annualHoursByTimeband,
                A6 => $annualHoursByTimeband,
                A7 => $peakingProbability,
                A9 => $daysInYear,
                $model->{coincidenceAdj}
                  && $model->{coincidenceAdj} =~ /redonly/i
                ? (
                    A8 => Constant(
                        name =>
                          'Time bands to apply the time band load coefficient',
                        defaultFormat => '0con',
                        cols          => $timebandSet,
                        data => [ 1, map { 0 } 2 .. @{ $timebandSet->{list} } ],
                    )
                  )
                : (),
            },
        );

    }

    push @{ $model->{timeOfDayResults} }, $pseudoLoadCoefficientBreakdown;

    my @paygUnitRates;
    my @pseudoLoadCoefficients = map {

        my $pseudoLoadCoefficient = SumProduct(
            name => 'Unit rate '
              . ( 1 + $_ )
              . ' pseudo load coefficient by network level',
            rows    => $relevantEndUsersByRate[$_],
            cols    => $networkLevelsTimebandAware,
            matrix  => $pseudoLoadCoefficientBreakdown,
            vector  => $timebandUseByRate[$_],
            tariffs => $relevantTariffsByRate[$_],
        );

        push @{ $model->{timeOfDayResults} }, $pseudoLoadCoefficient;

        $pseudoLoadCoefficient;

    } 0 .. $model->{maxUnitRates} - 1;

    \@pseudoLoadCoefficientsAgainstSystemPeak, \@pseudoLoadCoefficients;

}

1;
