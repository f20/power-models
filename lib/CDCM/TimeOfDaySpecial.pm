﻿package CDCM;

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

sub timeOfDaySpecial {

    my ( $model, $networkLevels, $componentMap, $allEndUsers, $daysInYear,
        $loadCoefficients, $volumeByEndUser, $unitsByEndUser )
      = @_;

    if ( $model->{dcp130preprocess}
        && ( my $d = $model->{dataset} ) )
    {
        if ( $d->{1070}[1]{'Annual hours'} ) {

            foreach my $c ( 1, 2 ) {
                $d->{1041}[$c]{"NHH UMS category $_"} =
                  $d->{1042}[$c]{"NHH UMS category $_"}
                  foreach qw(A B C D);
            }

            unless ( $d->{1053}[1]{'NHH UMS category B'} ) {
                my $units = 0;
                foreach my $c ( map { $d->{1053}[$_] } 1 .. 3 ) {
                    foreach my $t ( grep { /\bums\b/i } keys %$c ) {
                        $units += $c->{$t} if $c->{$t};
                        $c->{$t} = '';
                    }
                }
                $d->{1053}[1]{"NHH UMS category A"}  = 0.25 * $units;
                $d->{1053}[1]{"NHH UMS category B"}  = 0.50 * $units;
                $d->{1053}[1]{"NHH UMS category C"}  = 0.20 * $units;
                $d->{1053}[1]{"NHH UMS category D"}  = 0.05 * $units;
                $d->{1053}[4]{"NHH UMS category $_"} = '' foreach qw(A B C D);
                foreach my $ldno ( 'LDNO LV', 'LDNO HV', 'QNO LV', 'QNO HV' ) {
                    $d->{1053}[4]{"$ldno NHH UMS category $_"} =
                      $d->{1053}[1]{"$ldno NHH UMS category $_"} = ''
                      foreach qw(A B C D);
                }
            }

            $d->{1201}[1]{"NHH UMS category $_"} = '' foreach qw(A B C D);
            $d->{1201}[2]{"NHH UMS category $_"} = $d->{1201}[2]{"NHH UMS"}
              foreach qw(A B C D);

            die "Mismatch in 1068 for $model->{datafile}"
              if grep {
                $d->{1070}[1]{'Annual hours'} != $d->{1068}[1]{'Annual hours'}
              } 1 .. 3;

            foreach ( keys %{ $d->{1069}[1] } ) {
                $d->{1069}[4]{$_} = ''
                  unless $d->{1069}[4]{$_}
                  && $d->{1069}[4]{$_} =~ /^[0-9\.\s]+$/
                  && !( $d->{1069}[4]{$_} > $d->{1069}[1]{$_} );
            }

        }
    }

    my ( $ignore1, $plca1 ) = $model->timeOfDaySpecialRunner(
        $networkLevels,
        $componentMap,
        Labelset(
            name => 'Metered end users',
            list => [ grep { !/unmeter/i } @{ $allEndUsers->{list} } ]
        ),
        $daysInYear,
        $loadCoefficients,
        $volumeByEndUser,
        $unitsByEndUser,
    );

    my ( $ignore2, $plca2 ) = $model->timeOfDaySpecialRunner(
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
        1
    );

    my @pseudos =
      map {
        my $r        = $_ + 1;
        my $pl1      = $plca1->[$_];
        my $pl2      = $plca2->[$_];
        my $endUsers = Labelset(
            name => "Users which have a unit rate $r",
            list => [ @{ $pl1->{rows}{list} }, @{ $pl2->{rows}{list} } ]
        );
        Stack(
            name => 'Unit rate '
              . ( 1 + $_ )
              . ' pseudo load coefficient by network level (combined)',
            rows    => $endUsers,
            cols    => $networkLevels,
            tariffs => $model->{pcd} ? $endUsers : Labelset(
                name   => "Tariffs which have a unit rate $r",
                groups => $endUsers->{list}
            ),
            sources => [ $pl1, $pl2 ],
        );
      } 0 .. $model->{maxUnitRates} - 1;

    if ( $model->{coincidenceAdj} && $model->{coincidenceAdj} =~ /group2/i ) {

        my $relevantUsers = $pseudos[0]{rows};

        $relevantUsers = Labelset(
            list => [ grep { !/gener/i } @{ $relevantUsers->{list} } ] );

        my $tariffGroupset = Labelset( list => [ 'All demand tariffs', ] );

        my $mapping = Constant(
            name => 'Mapping of tariffs to '
              . 'tariff groups for coincidence adjustment factor',
            defaultFormat => '0connz',
            rows          => $relevantUsers,
            cols          => $tariffGroupset,
            data          => [ map { [1] } @{ $relevantUsers->{list} } ],
            byrow         => 1,
        );

        if ( $model->{coincidenceAdj} =~ /voltage/i ) {

            $relevantUsers =
              Labelset(
                list => [ grep { !/^hv sub/i } @{ $relevantUsers->{list} } ] );

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

        }    # end of overrides if $model->{coincidenceAdj} =~ /voltage/i

        my $red = Stack(
            name          => 'Contribution to peak band kW',
            defaultFormat => '0softnz',
            rows          => $relevantUsers,
            sources       => $model->{timeOfDayGroupRedSources},
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

        push @{ $model->{timeOfDayResults} },
          Columnset(
            name    => 'Estimated contributions to peak demand',
            columns => [ $red, $coin, ]
          );

        push @{ $model->{timeOfDayResults} },
          my $coinG = SumProduct(
            name          => 'Group contribution to system-peak-time kW',
            defaultFormat => '0softnz',
            matrix        => $mapping,
            vector        => $coin,
          );

        my $redG = SumProduct(
            name          => 'Group contribution to peak band kW',
            defaultFormat => '0softnz',
            matrix        => $mapping,
            vector        => $red,
        );

        my $adjust = Stack(
            name    => 'Load coefficient correction factor (combined)',
            rows    => $pseudos[0]{rows},
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
                Constant(
                    name => 'Default to 1',
                    data => [ [ map { 1; } @{ $pseudos[0]{rows}{list} } ] ],
                    rows => $pseudos[0]{rows},
                ),
            ]
        );

        $adjust = $adjust->{sources}[0]
          if $adjust->lastRow == $adjust->{sources}[0]->lastRow;    # hacky

        $_ = Arithmetic(
            name       => "$_->{name} adjusted",
            arithmetic => '=A1*A2',
            arguments  => { A1 => $_, A2 => $adjust },
            tariffs    => $_->{tariffs},
        ) foreach @pseudos;

    }

    push @{ $model->{timeOfDayResults} }, @pseudos;

    \@pseudos;

}

sub timeOfDaySpecialRunner {

    my (
        $model,           $networkLevels,  $componentMap,
        $allEndUsers,     $daysInYear,     $loadCoefficients,
        $volumeByEndUser, $unitsByEndUser, $blackYellowGreen
    ) = @_;

    my $timebandSet =
      $blackYellowGreen
      ? Labelset(
        name => 'Special distribution time bands',
        list => [qw(Black Yellow Green)],
      )
      : Labelset(
        name => 'Distribution time bands',
        list => [qw(Red Amber Green)],
      );

    my $annualHoursByTimebandRaw = Dataset(
        name => 'Typical annual hours by'
          . ( $blackYellowGreen ? ' special ' : ' ' )
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
              . ' according to'
              . ( $blackYellowGreen ? ' special ' : ' ' )
              . 'time band hours input data'
        ),
        defaultFormat => '0.0softnz',
        source        => $annualHoursByTimebandRaw,
    );

    my $annualHoursByTimeband = Arithmetic(
        name => 'Annual hours by'
          . ( $blackYellowGreen ? ' special ' : ' ' )
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
        name => 'Adjust annual hours by'
          . ( $blackYellowGreen ? ' special ' : ' ' )
          . 'distribution time band to match days in year',
        columns => [ $annualHoursByTimebandTotal, $annualHoursByTimeband ]
    );

    my $networkLevelsTimeband = Labelset(
        name => 'Network levels and'
          . ( $blackYellowGreen ? ' special ' : ' ' )
          . 'time bands',
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
      name => 'Network levels aware of'
      . ( $blackYellowGreen ? ' special ' : ' ' )
      . 'time bands',
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
            : 1 ? '=A2+A3-A4'    # allow use of blank peaking probability
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
            arithmetic    => '=A1+A2-A3',
            arguments     => {
                A1 => $amberPeaking,
                A2 => $redPeaking,
                A3 => $yellowPeaking,
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
                $greenPeaking,            $amberPeakingRate,
                $yellowPeaking,           $blackPeaking,
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
            data => [ map { ''; } @{ $networkLevelsTimebandAware->{list} } ],
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
            name => 'Total'
              . ( $blackYellowGreen ? ' special ' : ' ' )
              . 'probability (should be 100%)',
            rows          => $networkLevelsTimebandAware,
            defaultFormat => '%soft',
            source        => $peakingProbabilitiesTable
        );

        $peakingProbabilitiesTable = Arithmetic(
            name => (
                $model->{otneiErrors}
                  || $model->{lvDiversityWrong} ? 'Non-normalised '
                : 'Normalised '
              )
              . ( $blackYellowGreen ? ' special ' : ' ' )
              . 'peaking probabilities',
            defaultFormat => '%soft',
            arithmetic    => $model->{otneiErrors}
              || $model->{lvDiversityWrong} ? '=IF(A3,A1,A8/A9)'
            : '=IF(A3,A1/A2,A8/A9)',
            arguments => {
                A8 => $annualHoursByTimebandRaw,
                A9 => $annualHoursByTimebandTotal,
                A1 => $peakingProbabilitiesTable,
                A2 => $model->{totalProbability},
                A3 => $model->{totalProbability},
            }
        );

        Columnset(
            name => 'Normalisation of'
              . ( $blackYellowGreen ? ' special ' : ' ' )
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
                  . ' units by'
                  . ( $blackYellowGreen ? ' special ' : ' ' )
                  . 'distribution time band',
                validation => {
                    validate => 'decimal',
                    criteria => 'between',
                    minimum  => 0,
                    maximum  => 1,
                },
                number   => ( $blackYellowGreen ? 1063 : 1060 ) + $r,
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
                name => "Normalised split of rate $r units by"
                  . ( $blackYellowGreen ? ' special ' : ' ' )
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
                name => "Normalisation of split of rate $r units by"
                  . ( $blackYellowGreen ? ' special ' : ' ' )
                  . 'time band',
                columns => [ $totals, $inData ]
            );

        }

        $conData = Constant(
            name => 'Split of rate '
              . $r
              . ' units between'
              . ( $blackYellowGreen ? ' special ' : ' ' )
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
        $inData && $conData
          ? Stack(
            name => 'Split of rate '
              . $r
              . ' units between'
              . ( $blackYellowGreen ? ' special ' : ' ' )
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

    my $timebandLoadCoefficient;    # never used?

    # unadjusted; to be replaced below if there is a coincidence adjustment
    my $pseudoLoadCoefficientBreakdown = Arithmetic(
        name => 'Pseudo load coefficient by'
          . ( $blackYellowGreen ? ' special ' : ' ' )
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

    my @pseudoLoadCoefficientsAgainstSystemPeak;

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
            name => 'Peak band'
              . ( $blackYellowGreen ? ' special ' : ' ' )
              . 'load coefficient',
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
                                  !/gener/i
                              and !/(?:related|additional)/i
                              || $model->{coincidenceAdj}
                              && $model->{coincidenceAdj} =~ /268/
                              and !$componentMap->{$_}{"Unit rate $r2 p/kWh"}
                        } @{ $relevantEndUsersByRate[$_]{list} }
                      )
                    {

                        my $relevantUsers = Labelset
                          name => 'Relevant demand end users'
                          . " with $rt-rate tariffs",
                          list => \@relevant;

                        my $timebandUseByRateTotal = Arithmetic(
                            rows => $relevantUsers,
                            cols => $timebandSet,
                            name => 'Use of'
                              . ( $blackYellowGreen ? ' special ' : ' ' )
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
                            name => 'Top network level'
                              . ( $blackYellowGreen ? ' special ' : ' ' )
                              . 'load coefficient for '
                              . $rt
                              . '-rate tariffs',
                            vector => $peakMatrix,
                            matrix => Arithmetic(
                                name =>
                                  ( $blackYellowGreen ? 'Special l' : 'L' )
                                  . 'oad coefficient for '
                                  . $rt
                                  . '-rate tariffs',
                                arithmetic => $timebandLoadCoefficient
                                ? '=IF(A6>0,A1*A2*A4*24/A5,0)'
                                : '=IF(A6>0,A1*A4*24/A5,0)',
                                cols      => $timebandSet,
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
                            )
                          )
                          : Arithmetic(
                            name => 'Peak band'
                              . ( $blackYellowGreen ? ' special ' : ' ' )
                              . 'load coefficient for '
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
                          name => 'Calculation of implied'
                          . ( $blackYellowGreen ? ' special ' : ' ' )
                          . "load coefficients for $rt-rate users",
                          columns => [ $timebandUseByRateTotal, $implied ]
                          unless $peakMatrix;

                        $implied;

                    }
                    else { (); }

                } 0 .. $model->{maxUnitRates} - 1
            ]
        );

        $model->{coincidenceAdj} = 'group'
          unless defined $model->{coincidenceAdj};

        if ( $model->{coincidenceAdj} =~ /group2/i ) {

            my $red = Arithmetic(
                name          => 'Contribution to peak band kW',
                defaultFormat => '0softnz',
                arithmetic => $timebandLoadCoefficient ? '=A1*A9*A2/24/A3*1000'
                : '=A1*A2/24/A3*1000',
                arguments => {
                    A1 => $timebandLoadCoefficientAccording,
                    A2 => $unitsByEndUser,
                    A3 => $daysInYear,
                    $timebandLoadCoefficient
                    ? ( A9 => $timebandLoadCoefficient )
                    : (),
                },
            );

            push @{ $model->{timeOfDayGroupRedSources} },
              Arithmetic(
                name       => Label( $red->{name}, "$red->{name} (copy)" ),
                arguments  => { A1 => $red },
                cols       => 0,
                arithmetic => '=A1',
              );

        }

        elsif ($model->{coincidenceAdj} =~ /group/i
            && $model->{coincidenceAdj} !~ /group2/i )
        {

            @pseudoLoadCoefficientsAgainstSystemPeak = map {
                my $r = 1 + $_;
                Arithmetic(
                    name => "Unit rate $r"
                      . ( $blackYellowGreen ? ' special ' : ' ' )
                      . 'pseudo load coefficient at system level',
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
              } 0 .. $model->{maxUnitRates} - 1
              unless $peakMatrix;

            my $timebandLoadCoefficientAdjusted = Arithmetic(
                name => 'Load coefficient correction factor'
                  . ' (kW at peak in band / band average kW)',
                arithmetic => $timebandLoadCoefficient
                ? '=IF(A5<>0,A4/A2/A1,A6)'
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
                }
            );
            my $relevantUsers =
              Labelset( list =>
                  [ grep { !/gener/i } @{ $relevantEndUsersByRate[0]{list} } ]
              );

            my ( $tariffGroupset, $mapping );

            if ( $model->{coincidenceAdj} =~ /tcr/ && grep { /related mpan/i }
                @{ $relevantUsers->{list} } )
            {

                my ( @tariffGroups, @mappingData );

                for ( my $i = 0 ; $i < @{ $relevantUsers->{list} } ; ++$i ) {
                    local $_ = $relevantUsers->{list}[$i];
                    s/^.*\n//s;
                    s/(?: band [1-4]| no residual| with residual)$//i;
                    s/[ (]+related mpan\)*$//i;
                    push @tariffGroups, $_
                      unless @tariffGroups
                      && $tariffGroups[$#tariffGroups] eq $_;
                    $mappingData[$#tariffGroups][$i] = 1;
                }

                for ( my $i = 0 ; $i < @{ $relevantUsers->{list} } ; ++$i ) {
                    for ( my $g = 0 ; $g < @tariffGroups ; ++$g ) {
                        $mappingData[$g][$i] ||= 0;
                    }
                }

                $tariffGroupset = Labelset( list => \@tariffGroups );

                $mapping = Constant(
                    name => 'Mapping of tariffs to '
                      . 'tariff groups for coincidence adjustment factor',
                    defaultFormat => '0connz',
                    rows          => $relevantUsers,
                    cols          => $tariffGroupset,
                    data          => \@mappingData,
                );

            }

            elsif ( $model->{coincidenceAdj} =~ /268/
                && grep { /related mpan/i } @{ $relevantUsers->{list} } )
            {

                $relevantUsers = Labelset(
                    list => [
                        grep {
                            !$componentMap->{$_}{'Capacity charge p/kVA/day'};
                        } @{ $relevantUsers->{list} }
                    ]
                );

                $tariffGroupset =
                  Labelset( list =>
                      [ grep { !/related mpan/i } @{ $relevantUsers->{list} } ]
                  );

                my @mappingData =
                  map {
                    [ map { 0; } @{ $relevantUsers->{list} } ];
                  } @{ $tariffGroupset->{list} };
                {
                    my $j = -1;
                    for ( my $i = 0 ; $i < @{ $relevantUsers->{list} } ; ++$i )
                    {
                        ++$j
                          unless $relevantUsers->{list}[$i] =~ /related mpan/i;
                        $mappingData[$j][$i] = 1;
                    }
                }

                $mapping = Constant(
                    name => 'Mapping of tariffs to '
                      . 'tariff groups for coincidence adjustment factor',
                    defaultFormat => '0connz',
                    rows          => $relevantUsers,
                    cols          => $tariffGroupset,
                    data          => \@mappingData,
                );

            }

            elsif (
                $model->{coincidenceAdj} =~ /ums/i
                && (
                    my @relevantToGrouping =
                    grep { /\bUMS\b|un-?met/i } @{ $relevantUsers->{list} }
                )
              )
            {

                push @{ $model->{optionLines} },
                  'Coincidence correction factors grouped for UMS';

                $relevantUsers =
                  @{ $relevantEndUsersByRate[0]{list} } == @relevantToGrouping
                  ? $relevantEndUsersByRate[0]    # hack
                  : Labelset( list => \@relevantToGrouping );

                $tariffGroupset = Labelset( list => [ 'Unmetered', ] );

                $mapping = Constant(
                    name => 'Mapping of tariffs to '
                      . 'tariff groups for coincidence adjustment factor',
                    defaultFormat => '0connz',
                    rows          => $relevantUsers,
                    cols          => $tariffGroupset,
                    data          => [ map { 1; } @{ $relevantUsers->{list} } ],
                );

            }

            elsif (
                $model->{coincidenceAdj} =~ /hhcap/i
                && (
                    my @relevantToGrouping2 =
                    grep {
                             $componentMap->{$_}{'Capacity charge p/kVA/day'}
                          && $componentMap->{$_}{'Unit rates p/kWh'};
                    } @{ $relevantUsers->{list} }
                )
              )
            {

                push @{ $model->{optionLines} },
                  'Coincidence correction factors grouped for'
                  . ' half hourly tariffs with capacity charges';

                $relevantUsers =
                  @{ $relevantEndUsersByRate[0]{list} } == @relevantToGrouping2
                  ? $relevantEndUsersByRate[0]    # hack
                  : Labelset( list => \@relevantToGrouping2 );

                $tariffGroupset =
                  Labelset( list => [ 'Half hourly with capacity charges', ] );

                $mapping = Constant(
                    name => 'Mapping of tariffs to '
                      . 'tariff groups for coincidence adjustment factor',
                    defaultFormat => '0connz',
                    rows          => $relevantUsers,
                    cols          => $tariffGroupset,
                    data          => [ map { 1; } @{ $relevantUsers->{list} } ],
                );

            }

            elsif ( $model->{coincidenceAdj} =~ /voltage/i ) {

                push @{ $model->{optionLines} },
                  'Coincidence correction factors by'
                  . ' voltage level tariff group';

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
                    list => [ 'Domestic', 'Non-domestic non-CT', 'Other', ] );

                $mapping = Constant(
                    name => 'Mapping of tariffs to '
                      . 'tariff groups for coincidence adjustment factor',
                    defaultFormat => '0connz',
                    rows          => $relevantUsers,
                    cols          => $tariffGroupset,
                    data          => [
                        map {
                            /domestic/i && !/non.?dom/i ? [qw(1 0 0)]
                              : !$componentMap->{$_}{'Unit rates p/kWh'}
                              || /non.?ct/i ? [qw(0 1 0)]
                              : [qw(0 0 1)];
                        } @{ $relevantUsers->{list} }
                    ],
                    byrow => 1,
                );

            }

            elsif ( $model->{coincidenceAdj} =~ /pick/i ) {

                $tariffGroupset = Labelset( list =>
                      [ 'Domestic Unrestricted', 'LV Medium Non Domestic', ] );

                $relevantUsers = Labelset(
                    list => [
                        grep {
                                 /domestic/i
                              && !/non.?dom/i
                              && $componentMap->{$_}{'Fixed charge p/MPAN/day'}
                              && !$componentMap->{$_}{'Unit rate 2 p/kWh'}
                              || /(profile|pc).*[5-8]|medium/i
                              || $componentMap->{$_}{'Unit rates p/kWh'}
                              && !$componentMap->{$_}
                              {'Capacity charge p/kVA/day'};
                        } @{ $relevantUsers->{list} }
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
                            /domestic/i && !/non.?dom/i ? [qw(1 0)] : [qw(0 1)];
                        } @{ $relevantUsers->{list} }
                    ],
                    byrow => 1,
                );

            }

            elsif ( $model->{coincidenceAdj} =~ /domnondomvoltctnonct/i ) {

                $tariffGroupset = Labelset(
                    list => [
                        'Domestic',
                        'Non-domestic non-CT LV',
                        'Non-domestic CT LV',
                        'Non-domestic LV Sub',
                        'Non-domestic HV',
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
                                /domestic/i && !/non.?dom/i ? [qw(1 0 0 0 0)]
                              : /^hv/i                      ? [qw(0 0 0 0 1)]
                              : /^lv sub/i                  ? [qw(0 0 0 1 0)]
                              : !$componentMap->{$_}{'Unit rates p/kWh'}
                              || /non.?ct/i ? [qw(0 1 0 0 0)]
                              : [qw(0 0 1 0 0)];
                        } @{ $relevantUsers->{list} }
                    ],
                    byrow => 1,
                );

            }

            elsif ( $model->{coincidenceAdj} =~ /domnondomvolt/i ) {

                $tariffGroupset = Labelset(
                    list => [
                        'Domestic',
                        'Non-domestic LV',
                        'Non-domestic LV Sub',
                        'Non-domestic HV',
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
                                /domestic/i && !/non.?dom/i ? [qw(1 0 0 0)]
                              : /^hv/i                      ? [qw(0 0 0 1)]
                              : /^lv sub/i                  ? [qw(0 0 1 0)]
                              :                               [qw(0 1 0 0)];
                        } @{ $relevantUsers->{list} }
                    ],
                    byrow => 1,
                );

            }

            elsif ( $model->{coincidenceAdj} =~ /domnondom/i ) {

                $tariffGroupset =
                  Labelset( list => [ 'Domestic', 'Non-domestic', ] );

                $mapping = Constant(
                    name => 'Mapping of tariffs to '
                      . 'tariff groups for coincidence adjustment factor',
                    defaultFormat => '0connz',
                    rows          => $relevantUsers,
                    cols          => $tariffGroupset,
                    data          => [
                        map {
                            /domestic/i && !/non.?dom/i ? [qw(1 0)] : [qw(0 1)];
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
                    name          => 'Contribution to peak band kW',
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

                push @{ $model->{timeOfDayResults} },
                  Columnset(
                    name    => 'Estimated contributions to peak demand',
                    columns => [
                        $timebandLoadCoefficientAccording->{rows} ==
                          $red->{rows}
                        ? $timebandLoadCoefficientAccording
                        : (),
                        $red,
                        $coin,
                    ]
                  );

                push @{ $model->{timeOfDayResults} },
                  my $coinG = SumProduct(
                    name => 'Group contribution to system-peak-time kW',
                    defaultFormat => '0softnz',
                    matrix        => $mapping,
                    vector        => $coin,
                  );

                my $redG = SumProduct(
                    name          => 'Group contribution to peak band kW',
                    defaultFormat => '0softnz',
                    matrix        => $mapping,
                    vector        => $red,
                );

                $timebandLoadCoefficientAdjusted = Stack(
                    name    => 'Load coefficient correction factor (combined)',
                    rows    => $relevantEndUsersByRate[0],
                    cols    => $timebandLoadCoefficientAdjusted->{cols},
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

                push @{ $model->{timeOfDayResults} },
                  $timebandLoadCoefficientAdjusted;

            }

            push @{ $model->{timeOfDayResults} },
              Columnset(
                name => 'Calculation of adjusted'
                  . ( $blackYellowGreen ? ' special ' : ' ' )
                  . 'time band load coefficients',
                columns => [
                    $timebandLoadCoefficientAccording,
                    $timebandLoadCoefficientAdjusted
                ]
              ) unless $timebandLoadCoefficientAccording->{dontcolumnset};

            $pseudoLoadCoefficientBreakdown = Arithmetic(
                name => 'Pseudo load coefficient by'
                  . ( $blackYellowGreen ? ' special ' : ' ' )
                  . 'time band and network level',
                rows       => $relevantEndUsersByRate[0],
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
                            data          =>
                              [ 1, map { 0 } 2 .. @{ $timebandSet->{list} } ],
                        )
                      )
                    : (),
                },
            );

        }
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

    \@pseudoLoadCoefficientsAgainstSystemPeak, \@pseudoLoadCoefficients;

}

1;
