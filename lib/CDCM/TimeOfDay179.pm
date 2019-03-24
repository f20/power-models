package CDCM;

# Copyright 2009-2011 Energy Networks Association Limited and others.
# Copyright 2011-2016 Franck Latrémolière, Reckon LLP and others.
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

sub timeOfDay179 {

    my ( $model, $networkLevels, $componentMap, $allEndUsers, $daysInYear,
        $loadCoefficients, $volumeByEndUser, $unitsByEndUser )
      = @_;

    my (@pseudoLoadCoeffMetered) = $model->timeOfDay179Runner(
        $networkLevels,
        $componentMap,
        Labelset(
            name =>
              'Metered end users with directly calculated multi-rate tariffs',
            list => [
                grep {
                    !/unmeter/i
                      and $componentMap->{$_}{'Unit rate 0 p/kWh'}
                      || $componentMap->{$_}{'Unit rate 2 p/kWh'};
                } @{ $allEndUsers->{list} }
            ]
        ),
        $daysInYear,
        $loadCoefficients,
        $volumeByEndUser,
        $unitsByEndUser,
        '',
        ( $model->{agghhequalisation} ? 'equalisation' : '' )
          . (
              $model->{coincidenceAdj} && $model->{coincidenceAdj} !~ /ums/i
            ? $model->{coincidenceAdj}
            : 'none'
          ),
    );

    my (@pseudoLoadCoeffUnmetered) = $model->timeOfDay179Runner(
        $networkLevels,
        $componentMap,
        Labelset(
            name => 'Unmetered end users',
            list => [ grep { /unmeter/i } @{ $allEndUsers->{list} } ]
        ),
        $daysInYear,
        $loadCoefficients,
        $volumeByEndUser,
        $unitsByEndUser,
        'special ',
        'groupAll',
    );

    my @pseudoLoadCoeffs =
      map {
        my $r0      = $_;
        my $r       = $r0 + 1;
        my @sources = grep { $_ }
          map { $_->[$r0] } @pseudoLoadCoeffMetered,
          @pseudoLoadCoeffUnmetered;
        my %userSet;
        undef $userSet{$_} foreach map { @{ $_->{rows}{list} } } @sources;
        my $endUsers = Labelset(
            name => "Users which have a unit rate $r",
            list => [ grep { exists $userSet{$_} } @{ $allEndUsers->{list} } ]
        );
        Stack(
            name => "Unit rate $r"
              . ' pseudo load coefficient by network level (combined)',
            rows    => $endUsers,
            cols    => $networkLevels,
            tariffs => $model->{pcd} ? $endUsers : Labelset(
                name   => "Tariffs which have a unit rate $r",
                groups => $endUsers->{list}
            ),
            sources => \@sources,
        );
      } 0 .. $model->{maxUnitRates} - 1;

    push @{ $model->{timeOfDayResults} }, @pseudoLoadCoeffs;

    \@pseudoLoadCoeffs;

}

sub timeOfDay179Runner {

    my (
        $model,           $networkLevels,  $componentMap,
        $allEndUsers,     $daysInYear,     $loadCoefficients,
        $volumeByEndUser, $unitsByEndUser, $blackYellowGreen,
        $correctionRules,
    ) = @_;

    my $timebandSet = Labelset(
        name => ucfirst( $blackYellowGreen . 'distribution time bands' ),
        list => $blackYellowGreen
        ? [qw(Black Yellow Green)]
        : [qw(Red Amber Green)],
    );

    my $annualHoursByTimebandRaw = Dataset(
        name => 'Typical annual hours by '
          . $blackYellowGreen
          . 'distribution time band',
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
        number        => $blackYellowGreen ? 1066 : 1068,
        appendTo      => $model->{inputTables},
        dataset       => $model->{dataset},
        cols          => $timebandSet,
        data          => [qw(650 3224 4862)],
        defaultFormat => '0.0hardnz'
    );

    my $annualHoursByTimebandTotal = GroupBy(
        rows => 0,
        cols => 0,
        name => Label(
            'Hours aggregate',
            'Total hours in the year'
              . ' according to '
              . $blackYellowGreen
              . 'time band hours input data'
        ),
        defaultFormat => '0.0softnz',
        source        => $annualHoursByTimebandRaw,
    );

    my $annualHoursByTimeband = Arithmetic(
        name => 'Annual hours by '
          . $blackYellowGreen
          . 'distribution time band (reconciled to days in year)',
        singleRowName => 'Annual hours',
        defaultFormat => '0.0softnz',
        arithmetic    => '=A1*24*A3/A2',
        arguments     => {
            A2 => $annualHoursByTimebandTotal,
            A3 => $daysInYear,
            A1 => $annualHoursByTimebandRaw,
        }
    );

    $model->{hoursByRedAmberGreen} = $annualHoursByTimeband
      unless $blackYellowGreen;

    Columnset(
        name => 'Adjust annual hours by '
          . $blackYellowGreen
          . 'distribution time band to match days in year',
        columns => [ $annualHoursByTimebandTotal, $annualHoursByTimeband ]
    );

    my $networkLevelsTimeband = Labelset(
        name   => 'Network levels and ' . $blackYellowGreen . 'time bands',
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
      name    => 'Network levels aware of ' . $blackYellowGreen . 'time bands',
      list    => $networkLevelsTimeband->{groups},
      accepts => [$networkLevels];

    my $peakingProbabilitiesTable;

    if ($blackYellowGreen) {
        my $redPeaking = Arithmetic(
            name          => 'Red peaking probabilities',
            defaultFormat => '%copy',
            cols          => $model->{redAmberGreenRed},
            arithmetic    => '=A1',
            arguments     => { A1 => $model->{redAmberGreenPeaking}, }
        );
        my $amberPeaking = Arithmetic(
            name          => 'Amber peaking probabilities',
            defaultFormat => '%copy',
            cols          => $model->{redAmberGreenAmber},
            arithmetic    => '=A1',
            arguments     => { A1 => $model->{redAmberGreenPeaking}, }
        );
        my $greenPeaking = Arithmetic(
            name          => 'Green peaking probabilities',
            defaultFormat => '%copy',
            cols          => $model->{redAmberGreenGreen},
            arithmetic    => '=A1',
            arguments     => { A1 => $model->{redAmberGreenPeaking}, }
        );
        my $amberPeakingRate;
        $amberPeakingRate = Arithmetic(
            name       => 'Amber peaking rates',
            arithmetic => '=A1*24*A3/A2',
            arguments  => {
                A1 => $amberPeaking,
                A2 => $model->{redAmberGreenHours},
                A3 => $daysInYear,
            }
        ) unless $model->{blackPeakingProbabilityRequired};
        my $yellowPeaking = Arithmetic(
            name          => 'Yellow peaking probabilities',
            defaultFormat => '%soft',
            rows          => $amberPeaking->{rows},
            cols          => Labelset( list => [ $timebandSet->{list}[1] ] ),
            arithmetic    => $amberPeakingRate
            ? '=IF(A1,MAX(0,A2+A3-A4),A6*A7/A8/24)'
            : 1 ? '=A2+A3-A4'    # trust black probability even if blank
            : '=IF(A1,MAX(0,A2+A3-A4),IF(A5,1/0,0))',
            arguments => {
                A1 => $model->{blackPeaking},
                A2 => $amberPeaking,
                A3 => $redPeaking,
                A4 => $model->{blackPeaking},
                A5 => $model->{totalProbability},
                $amberPeakingRate
                ? (
                    A6 => $amberPeakingRate,
                    A7 => $annualHoursByTimeband,
                    A8 => $daysInYear,
                  )
                : (),
            }
        );
        my $blackPeaking = Arithmetic(
            name          => 'Black peaking probabilities',
            defaultFormat => '%soft',
            rows          => $greenPeaking->{rows},
            cols          => Labelset( list => [ $timebandSet->{list}[0] ] ),
            arithmetic    => '=1-A1-A2',
            arguments     => {
                A1 => $yellowPeaking,
                A2 => $greenPeaking,
            }
        );
        $peakingProbabilitiesTable = Stack(
            name => $peakingProbabilitiesTable->{name},
            name => ( $blackYellowGreen ? 'Special peaking' : 'Peaking' )
              . ' probabilities by network level',
            cols          => $timebandSet,
            rows          => $networkLevelsTimebandAware,
            sources       => [ $greenPeaking, $yellowPeaking, $blackPeaking, ],
            defaultFormat => '%copynz',
        );

        Columnset(
            name    => 'Calculation of special peaking probabilities',
            columns => [
                grep { $_; } $redPeaking, $amberPeaking,
                $greenPeaking,  $amberPeakingRate,
                $yellowPeaking, $blackPeaking,
            ]
        );

    }

    else {

        $peakingProbabilitiesTable = Dataset(
            name       => 'Red, amber and green peaking probabilities',
            validation => {
                validate => 'decimal',
                criteria => 'between',
                minimum  => 0,
                maximum  => 1,
            },
            cols  => $timebandSet,
            rows  => $networkLevelsTimebandAware,
            byrow => 1,
            data  => [
                map {
                    /(132|GSP|Trans)/i
                      ? (
                        @{ $timebandSet->{list} } > 2
                        ? [
                            .95, .05, map { 0 } 3 .. @{ $timebandSet->{list} }
                          ]
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

        $model->{blackPeaking} = Dataset(
            name       => 'Black peaking probabilities',
            validation => {
                validate => 'decimal',
                criteria => 'between',
                minimum  => 0,
                maximum  => 1,
            },
            rows => $networkLevelsTimebandAware,
            data => [ map { '' } @{ $networkLevelsTimebandAware->{list} } ],
            defaultFormat => '%hard',
        );

        Columnset(
            name     => 'Peaking probabilities by network level',
            lines    => 'Source: analysis of network operation data.',
            number   => 1069,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            columns  => [ $peakingProbabilitiesTable, $model->{blackPeaking}, ],
        );

        $model->{totalProbability} = GroupBy(
            name => 'Total '
              . $blackYellowGreen
              . 'probability (should be 100%)',
            rows          => $networkLevelsTimebandAware,
            defaultFormat => '%soft',
            source        => $peakingProbabilitiesTable
        );

        $peakingProbabilitiesTable = Arithmetic(
            name => 'Normalised ' . $blackYellowGreen . 'peaking probabilities',
            defaultFormat => '%soft',
            arithmetic    => "=IF(A3,A1/A2,A8/A9)",
            arguments     => {
                A8 => $annualHoursByTimebandRaw,
                A9 => $annualHoursByTimebandTotal,
                A1 => $peakingProbabilitiesTable,
                A2 => $model->{totalProbability},
                A3 => $model->{totalProbability},
            }
        );

        Columnset(
            name => 'Normalisation of '
              . $blackYellowGreen
              . 'peaking probabilities',
            columns =>
              [ $model->{totalProbability}, $peakingProbabilitiesTable ]
        );

        unless ($blackYellowGreen) {
            $model->{redAmberGreenPeaking} = $peakingProbabilitiesTable;
            $model->{redAmberGreenHours}   = $annualHoursByTimeband;
            $model->{redAmberGreenRed} =
              Labelset( list => [ $timebandSet->{list}[0] ] );
            $model->{redAmberGreenAmber} =
              Labelset( list => [ $timebandSet->{list}[1] ] );
            $model->{redAmberGreenGreen} =
              Labelset( list => [ $timebandSet->{list}[2] ] );
        }

    }

    my $peakingProbability = new SpreadsheetModel::Reshape(
        name => ( $blackYellowGreen ? 'Special peaking' : 'Peaking' )
          . ' probabilities by network level (reshaped)',
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
                  . ' units by '
                  . $blackYellowGreen
                  . 'distribution time band',
                validation => {
                    validate => 'decimal',
                    criteria => 'between',
                    minimum  => 0,
                    maximum  => 1,
                },
                number => ( $blackYellowGreen ? 1063 : 1060 ) + $r,
                appendTo => $model->{inputTables},
                dataset  => $model->{dataset},
                rows     => $usersWithInput,
                cols     => $timebandSet,
                byrow    => 1,
                data     => [
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
                name => "Normalised split of rate $r units by "
                  . $blackYellowGreen
                  . 'distribution time band',
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
                name => "Normalisation of split of rate $r units by "
                  . $blackYellowGreen
                  . 'time band',
                columns => [ $totals, $inData ]
            );

        }

        $conData = Constant(
            name => 'Split of rate '
              . $r
              . ' units between '
              . $blackYellowGreen
              . 'distribution time bands'
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
              . ' units between '
              . $blackYellowGreen
              . 'distribution time bands',
            rows          => $usersWithThisRate,
            cols          => $timebandSet,
            sources       => [ $inData, $conData ],
            defaultFormat => '%copynz'
          )
          : $inData ? $inData
          :           $conData;
    } 0 .. $model->{maxUnitRates} - 1;

    push @{ $model->{timeOfDayResults} }, @timebandUseByRate
      unless $model->{coincidenceAdj}
      && $model->{coincidenceAdj} =~ /none/i;

    # unadjusted; will be replaced below if there is a coincidence adjustment
    my $pseudoLoadCoefficientBreakdown = Arithmetic(
        name => 'Pseudo load coefficient by '
          . $blackYellowGreen
          . 'time band and network level',
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

    my @timebandUseByTariff;

    unless ( $model->{coincidenceAdj} && $model->{coincidenceAdj} =~ /none/i ) {

        my $peakBand = Labelset( list => [ $timebandSet->{list}[0] ] );
        my $peakMatrix;
        $peakMatrix = Stack(
            sources => [$peakingProbabilitiesTable],
            cols    => $timebandSet,
            rows    => (
                $peakBand = Labelset(
                    list => [ $peakingProbabilitiesTable->{rows}{list}[0] ]
                )
            ),
          )
          if $model->{coincidenceAdj}
          and $model->{coincidenceAdj} =~ /gspIsPeak/i
          || !$blackYellowGreen
          && $model->{coincidenceAdj} =~ /gspInsteadOfRed/i;

        my $timebandLoadCoefficientAccording = Stack(
            name    => 'Peak band ' . $blackYellowGreen . 'load coefficient',
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
                          name => 'Non-related-MPAN demand end users'
                          . " with $rt-rate tariffs",
                          list => \@relevant;

                        $timebandUseByTariff[$_] = Arithmetic(
                            rows => $relevantUsers,
                            cols => $timebandSet,
                            name => 'Use of '
                              . $blackYellowGreen
                              . 'distribution time bands by units'
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

                        my $implied =
                          $peakMatrix
                          ? SumProduct(
                            name => 'Top network level '
                              . $blackYellowGreen
                              . 'load coefficient for '
                              . $rt
                              . '-rate tariffs',
                            vector => $peakMatrix,
                            matrix => Arithmetic(
                                name => ucfirst(
                                        $blackYellowGreen
                                      . 'load coefficient for '
                                      . $rt
                                      . '-rate tariffs'
                                ),
                                arithmetic => '=IF(A6>0,A1*A4*24/A5,0)',
                                cols       => $timebandSet,
                                rows       => $relevantUsers,
                                arguments  => {
                                    A1 => $timebandUseByTariff[$_],
                                    A4 => $daysInYear,
                                    A5 => $annualHoursByTimeband,
                                    A6 => $annualHoursByTimeband
                                }
                            )
                          )
                          : Arithmetic(
                            name => 'Peak band '
                              . $blackYellowGreen
                              . 'load coefficient for '
                              . $rt
                              . '-rate tariffs',
                            arithmetic => '=IF(A6>0,A1*A4*24/A5,0)',
                            cols       => $peakBand,
                            rows       => $relevantUsers,
                            arguments  => {
                                A1 => $timebandUseByTariff[$_],
                                A4 => $daysInYear,
                                A5 => $annualHoursByTimeband,
                                A6 => $annualHoursByTimeband
                            }
                          );

                        Columnset
                          name => 'Calculation of implied '
                          . $blackYellowGreen
                          . "load coefficients for $rt-rate users",
                          columns => [ $timebandUseByTariff[$_], $implied ]
                          unless $peakMatrix;

                        $implied;

                    }
                    else { (); }

                } 0 .. $model->{maxUnitRates} - 1
            ]
        );

        my $timebandLoadCoefficientAdjusted = Arithmetic(
            name => 'Load coefficient correction factor'
              . ' (kW at peak in band / band average kW)',
            arithmetic => '=IF(A5<>0,A4/A2,IF(A8<0,-1,1))',
            rows       => $relevantEndUsersByRate[0],
            $model->{timebandCoef} && $model->{timebandCoef} =~ /detail/i
            ? ( cols => $networkLevelsTimeband )
            : (),
            arguments => {
                A2 => $timebandLoadCoefficientAccording,
                A5 => $timebandLoadCoefficientAccording,
                A4 => $loadCoefficients,
                A8 => $loadCoefficients,
            }
        );

        if ( $correctionRules =~ /groupAll/i ) {

            my $relevantUsers = $relevantEndUsersByRate[0];

            my $red = Arithmetic(
                name          => 'Contribution to peak band kW',
                defaultFormat => '0softnz',
                arithmetic    => '=A1*A2/24/A3*1000',
                rows          => $relevantUsers,
                arguments     => {
                    A1 => $timebandLoadCoefficientAccording,
                    A2 => $unitsByEndUser,
                    A3 => $daysInYear,
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

            push @{ $model->{timeOfDayResults} },
              Columnset(
                name    => 'Estimated contributions to peak demand',
                columns => [ $timebandLoadCoefficientAccording, $red, $coin, ]
              );

            $timebandLoadCoefficientAdjusted = Arithmetic(
                name => 'Load coefficient correction factor for the group',
                arithmetic => '=IF(SUM(A1_A5),SUM(A2_A6)/SUM(A3_A7),0)',
                rows       => 0,
                arguments  => {
                    A1_A5 => $red,
                    A2_A6 => $coin,
                    A3_A7 => $red,
                }
            );

        }

        elsif ( $correctionRules =~ /group/i ) {

            my $relevantUsers =
              Labelset( list =>
                  [ grep { !/gener/i } @{ $relevantEndUsersByRate[0]{list} } ]
              );

            my ( $tariffGroupset, $mapping );

            if ( $correctionRules =~ /voltage/i ) {

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
            elsif ( $correctionRules =~ /3|three/i ) {

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

            elsif (
                $correctionRules =~ /hhcap/i
                && (
                    my @relevantToGrouping2 =
                    grep {
                             $componentMap->{$_}{'Capacity charge p/kVA/day'}
                          && $componentMap->{$_}{'Unit rates p/kWh'};
                    } @{ $relevantUsers->{list} }
                )
              )
            {

                $relevantUsers =
                  @{ $relevantEndUsersByRate[0]{list} } == @relevantToGrouping2
                  ? $relevantEndUsersByRate[0]    # hack
                  : Labelset( list => \@relevantToGrouping2 );

                if ( $correctionRules =~ /level/i ) {

                    $tariffGroupset =
                      Labelset( list => [ 'LV HH', 'LV Sub HH', 'HV HH', ] );

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
                else {

                    $tariffGroupset =
                      Labelset(
                        list => [ 'Half hourly with capacity charges', ] );

                    $mapping = Constant(
                        name => 'Mapping of tariffs to '
                          . 'tariff groups for coincidence adjustment factor',
                        defaultFormat => '0connz',
                        rows          => $relevantUsers,
                        cols          => $tariffGroupset,
                        data => [ map { 1; } @{ $relevantUsers->{list} } ],
                    );

                }

            }

            if ($mapping) {

                my $red = Arithmetic(
                    name          => 'Contribution to peak band kW',
                    defaultFormat => '0softnz',
                    arithmetic    => '=A1*A2/24/A3*1000',
                    rows          => $relevantUsers,
                    arguments     => {
                        A1 => $timebandLoadCoefficientAccording,
                        A2 => $unitsByEndUser,
                        A3 => $daysInYear,
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
                    name => 'Estimated contributions to peak demand',
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
          $timebandLoadCoefficientAccording->{dontcolumnset}
          ? $timebandLoadCoefficientAdjusted
          : Columnset(
            name => 'Calculation of adjusted'
              . ( $blackYellowGreen ? ' special ' : ' ' )
              . 'time band load coefficients',
            columns => [
                $timebandLoadCoefficientAccording,
                $timebandLoadCoefficientAdjusted
            ]
          );

        $pseudoLoadCoefficientBreakdown = Arithmetic(
            name => 'Pseudo load coefficient by'
              . ( $blackYellowGreen ? ' special ' : ' ' )
              . 'time band and network level',
            $correctionRules =~ /groupAll/i
            ? ()
            : ( rows => $relevantEndUsersByRate[0] ),
            cols       => $networkLevelsTimeband,
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
                        name => 'Time bands to apply the'
                          . ( $blackYellowGreen ? ' special ' : ' ' )
                          . 'time band load coefficient',
                        defaultFormat => '0con',
                        cols          => $timebandSet,
                        data => [ 1, map { 0 } 2 .. @{ $timebandSet->{list} } ],
                    )
                  )
                : (),
            },
        );

    }

    if ( $correctionRules =~ /equalisation/i ) {

        my $groupset = Labelset(
            name => 'HH/NHH equalisation groups',
            list => [
                'Domestic equalisation group',
                'Non-domestic equalisation group'
            ]
        );

        my %eq;

        $eq{prefix} = {
            NHH1 => 'Single rate non half hourly',
            NHH2 => 'Multi rate non half hourly',
            NHH0 => 'Off-peak non half hourly',
            AHH  => 'Aggregated half hourly',
        };

        $eq{userSet}{NHH1} = Labelset(
            list => [
                grep { /domestic/i && /unrestricted/i }
                  @{ $allEndUsers->{list} }
            ]
        );

        $eq{userSet}{NHH2} = Labelset(
            list => [
                grep {
                         /domestic/i
                      && /two.?rate/i
                      && ( !/non.?dom/i || /small/i )
                } @{ $allEndUsers->{list} }
            ]
        );

        $eq{userSet}{NHH0} = Labelset(
            list => [
                grep {
                         /domestic/i
                      && /off.?peak|additional|related/i
                      && ( !/non.?dom/i || /small/i )
                } @{ $allEndUsers->{list} }
            ]
        );

        $eq{userSet}{AHH} = Labelset(
            list => [
                grep {
                         !/gener/i
                      && !/netting/i
                      && $componentMap->{$_}{'Unit rates p/kWh'}
                      && !$componentMap->{$_}{'Capacity charge p/kVA/day'}
                } @{ $allEndUsers->{list} }
            ]
        );

        foreach ( values %{ $eq{userSet} } ) {
            die "Mismatch: @{ $_->{list} } != @{ $groupset->{list} }"
              if @{ $_->{list} } != @{ $groupset->{list} };
            push @{ $_->{accepts} },        $groupset;
            push @{ $groupset->{accepts} }, $_;
        }

        push @{ $eq{userSet}{AHH}{accepts} },
          map { $eq{userSet}{$_} } qw(NHH1 NHH2 NHH0);

        $eq{units}{$_} = Stack(
            name          => $eq{prefix}{$_} . ' units (MWh)',
            defaultFormat => '0softnz',
            rows          => $eq{userSet}{$_},
            sources       => [$unitsByEndUser]
        ) foreach qw(NHH1 NHH2 NHH0 AHH);

        $eq{timebandUse}{$_} = Stack(
            name          => $eq{prefix}{$_} . ' timeband use',
            defaultFormat => '%softnz',
            rows          => $eq{userSet}{$_},
            sources       => [ $timebandUseByRate[0] ]
        ) foreach qw(NHH1 NHH0);

        $eq{timebandUse}{$_} = Stack(
            name          => $eq{prefix}{$_} . ' timeband use',
            defaultFormat => '%softnz',
            rows          => $eq{userSet}{$_},
            sources => [ $timebandUseByTariff[ /NHH2/ ? 1 : /AHH/ ? 2 : 0 ] ]
        ) foreach qw(NHH2 AHH);

        foreach ( $model->{agghhequalisation} =~ /rag/i ? () : qw(NHH1) ) {
            $eq{timebandCoef}{$_} = $eq{tariffCoef}{$_} = Stack(
                name    => $eq{prefix}{$_} . ' tariff load coefficient',
                rows    => $eq{userSet}{$_},
                sources => [$loadCoefficients],
            );
        }

        foreach ( $model->{agghhequalisation} =~ /rag/i ? qw(NHH1) : (),
            qw(NHH2 NHH0 AHH) )
        {
            $eq{timebandCoef}{$_} = Stack(
                name => $eq{prefix}{$_} . ' pseudo timeband load coefficients',
                rows => $eq{userSet}{$_},
                sources => [$pseudoLoadCoefficientBreakdown]
            );
            $eq{tariffCoef}{$_} = SumProduct(
                name   => $eq{prefix}{$_} . ' tariff pseudo load coefficient',
                rows   => $eq{userSet}{$_},
                cols   => $networkLevelsTimebandAware,
                vector => $eq{timebandUse}{$_},
                matrix => $eq{timebandCoef}{$_},
            );
        }

        $eq{tariffCoef}{hybrid} = SumProduct(
            name => 'Aggregated half hourly tariff pseudo load coefficient'
              . ' using average non half hourly unit mix',
            rows   => $eq{userSet}{AHH},
            cols   => $networkLevelsTimebandAware,
            vector => Arithmetic(
                name          => 'Average non half hourly timeband use',
                defaultFormat => '%softnz',
                rows          => $eq{userSet}{AHH},
                cols          => $timebandSet,
                arithmetic    => $model->{agghhequalisation} =~ /nooffpeak/i
                ? '=(A1*A2+A3*A4)/(A7+A8)'
                : '=(A1*A2+A3*A4+A5*A6)/(A7+A8+A9)',
                arguments => {
                    A1 => $eq{units}{NHH1},
                    A2 => $eq{timebandUse}{NHH1},
                    A3 => $eq{units}{NHH2},
                    A4 => $eq{timebandUse}{NHH2},
                    A7 => $eq{units}{NHH1},
                    A8 => $eq{units}{NHH2},
                    $model->{agghhequalisation} =~ /nooffpeak/i ? ()
                    : (
                        A5 => $eq{units}{NHH0},
                        A6 => $eq{timebandUse}{NHH0},
                        A9 => $eq{units}{NHH0}
                    ),
                },
            ),
            matrix => $eq{timebandCoef}{AHH},
        );

        $eq{tariffCoef}{NHH} = Arithmetic(
            name => 'Average non half hourly tariff pseudo load coefficient',
            rows => $groupset,
            cols => $networkLevelsTimebandAware,
            arithmetic => $model->{agghhequalisation} =~ /nooffpeak/i
            ? '=(A1*A2+A3*A4)/(A7+A8)'
            : '=(A1*A2+A3*A4+A5*A6)/(A7+A8+A9)',
            arguments => {
                A1 => $eq{units}{NHH1},
                A2 => $eq{tariffCoef}{NHH1},
                A3 => $eq{units}{NHH2},
                A4 => $eq{tariffCoef}{NHH2},
                A7 => $eq{units}{NHH1},
                A8 => $eq{units}{NHH2},
                $model->{agghhequalisation} =~ /nooffpeak/i ? ()
                : (
                    A5 => $eq{units}{NHH0},
                    A6 => $eq{tariffCoef}{NHH0},
                    A9 => $eq{units}{NHH0},
                ),
            },
        );

        $eq{relCorr} = Arithmetic(
            name =>
              'Relative correction factor for aggregated half hourly tariff',
            rows       => $groupset,
            cols       => $networkLevelsTimebandAware,
            arithmetic => '=A1/A9',
            arguments  => {
                A1 => $eq{tariffCoef}{NHH},
                A9 => $eq{tariffCoef}{hybrid},
            },
        );

        $eq{nhhCorr} = Arithmetic(
            name       => 'Correction factor for non half hourly tariffs',
            rows       => $groupset,
            cols       => $networkLevelsTimebandAware,
            arithmetic => $model->{agghhequalisation} =~ /nooffpeak/i
            ? '=(A11*A21+A31*A41+A71*A81)' . '/(A12*A22+A32*A42+A72*A82*A9)'
            : '=(A11*A21+A31*A41+A51*A61+A71*A81)'
              . '/(A12*A22+A32*A42+A52*A62+A72*A82*A9)',
            arguments => {
                A11 => $eq{units}{NHH1},
                A12 => $eq{units}{NHH1},
                A21 => $eq{tariffCoef}{NHH1},
                A22 => $eq{tariffCoef}{NHH1},
                A31 => $eq{units}{NHH2},
                A32 => $eq{units}{NHH2},
                A41 => $eq{tariffCoef}{NHH2},
                A42 => $eq{tariffCoef}{NHH2},
                A71 => $eq{units}{AHH},
                A72 => $eq{units}{AHH},
                A81 => $eq{tariffCoef}{AHH},
                A82 => $eq{tariffCoef}{AHH},
                A9  => $eq{relCorr},
                $model->{agghhequalisation} =~ /nooffpeak/i ? ()
                : (
                    A51 => $eq{units}{NHH0},
                    A52 => $eq{units}{NHH0},
                    A61 => $eq{tariffCoef}{NHH0},
                    A62 => $eq{tariffCoef}{NHH0},
                ),
            },
        );

        $pseudoLoadCoefficientBreakdown = Stack(
            name => 'Pseudo load coefficient by time band'
              . ' and network level (equalised)',
            rows    => $pseudoLoadCoefficientBreakdown->{rows},
            cols    => $pseudoLoadCoefficientBreakdown->{cols},
            sources => [
                (
                    map {
                        Arithmetic(
                            name => $eq{prefix}{$_}
                              . ' corrected pseudo timeband load coefficient',
                            rows       => $eq{userSet}{$_},
                            cols       => $networkLevelsTimeband,
                            arithmetic => '=A1*A2',
                            arguments  => {
                                A1 => $eq{timebandCoef}{$_},
                                A2 => $eq{nhhCorr},
                            },
                          )
                    } qw(NHH1 NHH2),
                    $model->{agghhequalisation} =~ /nooffpeak/i
                    ? ()
                    : qw(NHH0),
                ),
                Arithmetic(
                    name => $eq{prefix}{AHH}
                      . ' corrected pseudo timeband load coefficient',
                    rows       => $eq{userSet}{AHH},
                    arithmetic => '=A1*A2*A3',
                    arguments  => {
                        A1 => $eq{timebandCoef}{AHH},
                        A2 => $eq{nhhCorr},
                        A3 => $eq{relCorr},
                    },
                ),
                $pseudoLoadCoefficientBreakdown,
            ],
        );

    }

    push @{ $model->{timeOfDayResults} }, $pseudoLoadCoefficientBreakdown;

    my @paygUnitRates;
    my @pseudoLoadCoefficients = map {

        my $pseudoLoadCoefficient = SumProduct(
            name => 'Unit rate '
              . ( 1 + $_ )
              . ' pseudo load coefficient by network level'
              . ( $blackYellowGreen ? ' (special)' : '' ),
            rows    => $relevantEndUsersByRate[$_],
            cols    => $networkLevelsTimebandAware,
            matrix  => $pseudoLoadCoefficientBreakdown,
            vector  => $timebandUseByRate[$_],
            tariffs => $relevantTariffsByRate[$_],
        );

        push @{ $model->{timeOfDayResults} }, $pseudoLoadCoefficient;

        $pseudoLoadCoefficient;

    } 0 .. $model->{maxUnitRates} - 1;

    \@pseudoLoadCoefficients;

}

1;
