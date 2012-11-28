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
                foreach my $ldno (qw(LV HV)) {
                    $d->{1053}[4]{"LDNO $ldno NHH UMS category $_"} =
                      $d->{1053}[1]{"LDNO $ldno NHH UMS category $_"} = ''
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

    my ( $ignore1, $plca1 ) = $model->timeOfDayRunner(
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

    my ( $ignore2, $plca2 ) = $model->timeOfDayRunner(
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

    [
        map {
            my $r        = $_ + 1;
            my $pl1      = $plca1->[$_];
            my $pl2      = $plca2->[$_];
            my $endusers = Labelset(
                name => "Users which have a unit rate $r",
                list => [ @{ $pl1->{rows}{list} }, @{ $pl2->{rows}{list} } ]
            );
            push @{ $model->{timeOfDayResults} }, my $pl = Stack(
                name => 'Unit rate '
                  . ( 1 + $_ )
                  . ' pseudo load coefficient by network level (combined)',
                rows    => $endusers,
                cols    => $networkLevels,
                tariffs => $model->{pcd} ? $endusers : Labelset(
                    name   => "Tariffs which have a unit rate $r",
                    groups => $endusers->{list}
                ),
                sources => [ $pl1, $pl2 ],
            );
            $pl;
        } 0 .. $model->{maxUnitRates} - 1
    ];

}

sub timeOfDayRunner {

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
        arithmetic    => '=IV1*24*IV3/IV2',
        arguments     => {
            IV2 => $annualHoursByTimebandTotal,
            IV3 => $daysInYear,
            IV1 => $annualHoursByTimebandRaw,
        }
    );

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
            arithmetic    => '=IV1',
            arguments     => { IV1 => $model->{redAmberGreenPeaking}, }
        );
        my $amberPeaking = Arithmetic(
            name          => 'Amber peaking probabilities',
            defaultFormat => '%copy',
            cols          => $model->{redAmberGreenAmber},
            arithmetic    => '=IV1',
            arguments     => { IV1 => $model->{redAmberGreenPeaking}, }
        );
        my $greenPeaking = Arithmetic(
            name          => 'Green peaking probabilities',
            defaultFormat => '%copy',
            cols          => $model->{redAmberGreenGreen},
            arithmetic    => '=IV1',
            arguments     => { IV1 => $model->{redAmberGreenPeaking}, }
        );
        my $amberPeakingRate = Arithmetic(
            name       => 'Amber peaking rates',
            arithmetic => '=IV1*24*IV3/IV2',
            arguments  => {
                IV1 => $amberPeaking,
                IV2 => $model->{redAmberGreenHours},
                IV3 => $daysInYear,
            }
        );
        my $yellowPeaking = Arithmetic(
            name          => 'Yellow peaking probabilities',
            defaultFormat => '%soft',
            rows          => $amberPeaking->{rows},
            cols          => Labelset( list => [ $timebandSet->{list}[1] ] ),
            arithmetic    => '=IF(IV1,IV2+IV3-IV4,IV6*IV7/IV8/24)',
            arguments     => {
                IV1 => $model->{blackPeaking},
                IV2 => $amberPeaking,
                IV3 => $redPeaking,
                IV4 => $model->{blackPeaking},
                IV6 => $amberPeakingRate,
                IV7 => $annualHoursByTimeband,
                IV8 => $daysInYear,
            }
        );
        my $blackPeaking = Arithmetic(
            name          => 'Black peaking probabilities',
            defaultFormat => '%soft',
            rows          => $greenPeaking->{rows},
            cols          => Labelset( list => [ $timebandSet->{list}[0] ] ),
            arithmetic    => '=1-IV1-IV2',
            arguments     => {
                IV1 => $yellowPeaking,
                IV2 => $greenPeaking,
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
                $redPeaking,       $amberPeaking,  $greenPeaking,
                $amberPeakingRate, $yellowPeaking, $blackPeaking,
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

        my $totalProbability = GroupBy(
            name => 'Total'
              . ( $blackYellowGreen ? ' special ' : ' ' )
              . 'probability (should be 100%)',
            rows          => $networkLevelsTimebandAware,
            defaultFormat => '%soft',
            source        => $peakingProbabilitiesTable
        );

        $peakingProbabilitiesTable = Arithmetic(
            name => 'Normalised'
              . ( $blackYellowGreen ? ' special ' : ' ' )
              . 'peaking probabilities',
            defaultFormat => '%soft',
            arithmetic    => "=IF(IV3,IV1/IV2,IV8/IV9)",
            arguments     => {
                IV8 => $annualHoursByTimebandRaw,
                IV9 => $annualHoursByTimebandTotal,
                IV1 => $peakingProbabilitiesTable,
                IV2 => $totalProbability,
                IV3 => $totalProbability,
            }
        );

        Columnset(
            name => 'Normalisation of'
              . ( $blackYellowGreen ? ' special ' : ' ' )
              . 'peaking probabilities',
            columns => [ $totalProbability, $peakingProbabilitiesTable ]
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
        name   => 'Tariffs for multiple unit rate calulation',
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
                name => "Normalised split of rate $r units by"
                  . ( $blackYellowGreen ? ' special ' : ' ' )
                  . 'distribution time band',
                defaultFormat => '%soft',
                arithmetic    => "=IF(IV3,IV1/IV2,IV8/IV9/24)",
                arguments     => {
                    IV8 => $annualHoursByTimeband,
                    IV9 => $daysInYear,
                    IV1 => $inData,
                    IV2 => $totals,
                    IV3 => $totals,
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

    my $timebandLoadCoefficient;

    # unadjusted; to be replaced below if there is a coincidence adjustment
    my $pseudoLoadCoefficientBreakdown = Arithmetic(
        name => 'Pseudo load coefficient by'
          . ( $blackYellowGreen ? ' special ' : ' ' )
          . 'time band and network level',
        rows       => $relevantEndUsersByRate[0],
        cols       => $networkLevelsTimeband,
        arithmetic => '=IF(IV6>0,IV7*24*IV9/IV5,0)*IF(IV2<0,-1,1)',
        arguments  => {
            IV5 => $annualHoursByTimeband,
            IV6 => $annualHoursByTimeband,
            IV7 => $peakingProbability,
            IV9 => $daysInYear,
            IV2 => $loadCoefficients,
        }
    );

    my @pseudoLoadCoefficientsAgainstSystemPeak;

    unless ( $model->{coincidenceAdj} && $model->{coincidenceAdj} =~ /none/i ) {
        my $peakBand = Labelset( list => [ $timebandSet->{list}[0] ] );

        my $timebandLoadCoefficientAccording = Stack(
            name => 'First-time-band'
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
                            name => 'Use of'
                              . ( $blackYellowGreen ? ' special ' : ' ' )
                              . 'distribution time bands by units'
                              . " in demand forecast for $rt-rate tariffs",
                            arithmetic => '=IF(IV501>0,(' . join(
                                '+',
                                map {
                                    my $pad = "$_";
                                    $pad = "0$pad" while length $pad < 3;
                                    "IV1$pad*IV2$pad"
                                } 1 .. $r
                              )
                              . ')/IV502,0)',
                            arguments => {
                                IV501 => $unitsByEndUser,
                                IV502 => $unitsByEndUser,
                                map {
                                    my $pad = $_;
                                    $pad = "0$pad" while length $pad < 3;
                                    (
                                        "IV1$pad" => $volumeByEndUser->{
                                            "Unit rate $_ p/kWh"},
                                        "IV2$pad" =>
                                          $timebandUseByRate[ $_ - 1 ]
                                      )
                                } 1 .. $r
                            },
                            defaultFormat => '%softnz'
                        );

                        my $implied = Arithmetic(
                            name => 'First-time-band'
                              . ( $blackYellowGreen ? ' special ' : ' ' )
                              . 'load coefficient for '
                              . $rt
                              . '-rate tariffs',
                            arithmetic => $timebandLoadCoefficient
                            ? '=IF(IV6>0,IV1*IV2*IV4*24/IV5,0)'
                            : '=IF(IV6>0,IV1*IV4*24/IV5,0)',
                            cols      => $peakBand,
                            rows      => $relevantUsers,
                            arguments => {
                                $timebandLoadCoefficient
                                ? ( IV2 => $timebandLoadCoefficient )
                                : (),
                                IV1 => $timebandUseByRateTotal,
                                IV4 => $daysInYear,
                                IV5 => $annualHoursByTimeband,
                                IV6 => $annualHoursByTimeband
                            }
                        );

                        Columnset
                          name => 'Calculation of implied'
                          . ( $blackYellowGreen ? ' special ' : ' ' )
                          . "load coefficients for $rt-rate users",
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
                name => "Unit rate $r"
                  . ( $blackYellowGreen ? ' special ' : ' ' )
                  . 'pseudo load coefficient at system level',
                arithmetic =>
'=IF(IV6>0,IV1*IF(IV7<>0,IV3/IV2,IF(IV9<0,-1,1))*24*IV4/IV5,0)',
                cols      => $peakBand,
                arguments => {
                    IV1 => $timebandUseByRate[$_],
                    IV2 => $timebandLoadCoefficientAccording,
                    IV3 => $loadCoefficients,
                    IV9 => $loadCoefficients,
                    IV4 => $daysInYear,
                    IV5 => $annualHoursByTimeband,
                    IV6 => $annualHoursByTimeband,
                    IV7 => $timebandLoadCoefficientAccording
                },
                rows    => $relevantEndUsersByRate[$_],
                tariffs => $relevantTariffsByRate[$_],
            );
        } 0 .. $model->{maxUnitRates} - 1;

        my $timebandLoadCoefficientAdjusted =
          Arithmetic
          name => 'Load coefficient correction factor'
          . ' (kW at peak in band / band average kW)',
          arithmetic => $timebandLoadCoefficient
          ? '=IF(IV5<>0,IV4/IV2/IV1,IV6)'
          : '=IF(IV5<>0,IV4/IV2,IF(IV8<0,-1,1))',
          rows => $relevantEndUsersByRate[0],
          $model->{timebandCoef} && $model->{timebandCoef} =~ /detail/i
          ? ( cols => $networkLevelsTimeband )
          : (    $model->{coincidenceAdj}
              && $model->{coincidenceAdj} =~ /redonly/i )
          ? ( cols => $peakBand )
          : (), arguments => {
            $timebandLoadCoefficient
            ? (
                IV1 => $timebandLoadCoefficient,
                IV6 => $timebandLoadCoefficient
              )
            : (),
            IV2 => $timebandLoadCoefficientAccording,
            IV5 => $timebandLoadCoefficientAccording,
            IV4 => $loadCoefficients,
            IV8 => $loadCoefficients,
          };

        if ( $model->{coincidenceAdj} && $model->{coincidenceAdj} =~ /group/i )
        {
          GROUPING: {

                my $relevantUsers = Labelset(
                    list => [
                        grep { !/gener/i } @{ $relevantEndUsersByRate[0]{list} }
                    ]
                );

                my $tariffGroupset =
                  Labelset( list => [ 'All demand tariffs', ] );

                my $mapping = Constant(
                    name => 'Mapping of tariffs to '
                      . 'tariff groups for coincidence adjustment factor',
                    defaultFormat => '0connz',
                    rows          => $relevantUsers,
                    cols          => $tariffGroupset,
                    data  => [ map { [1] } @{ $relevantUsers->{list} } ],
                    byrow => 1,
                );

                if ( $model->{coincidenceAdj} =~ /ums/i ) {

                    my @relevantToGrouping =
                      grep { /\bUMS\b|un-?met/i } @{ $relevantUsers->{list} };
                    last GROUPING unless @relevantToGrouping;

                    push @{ $model->{optionLines} },
                      'Coincidence correction factors grouped for UMS';

                    $relevantUsers =
                      @{ $relevantEndUsersByRate[0]{list} } ==
                      @relevantToGrouping
                      ? $relevantEndUsersByRate[0]    # hack
                      : Labelset( list => \@relevantToGrouping );

                    $tariffGroupset = Labelset( list => [ 'Unmetered', ] );

                    $mapping = Constant(
                        name => 'Mapping of tariffs to '
                          . 'tariff groups for coincidence adjustment factor',
                        defaultFormat => '0connz',
                        rows          => $relevantUsers,
                        cols          => $tariffGroupset,
                        data => [ map { 1; } @{ $relevantUsers->{list} } ],
                    );

                }
                elsif ( $model->{coincidenceAdj} =~ /voltage/i ) {

                    push @{ $model->{optionLines} },
                      'Coincidence correction factors by'
                      . ' voltage level tariff group';

                    $relevantUsers =
                      Labelset( list =>
                          [ grep { !/^hv sub/i } @{ $relevantUsers->{list} } ]
                      );

                    $tariffGroupset =
                      Labelset( list =>
                          [ 'LV network', 'LV substation', 'HV network', ] );

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
                                  || !$componentMap->{$_}
                                  {'Fixed charge p/MPAN/day'}
                                  && !$componentMap->{$_}{'Unit rates p/kWh'}
                                  ? [qw(1 0 0)]
                                  : /wc|small/i ? [qw(0 1 0)]
                                  :               [qw(0 0 1)];
                            } @{ $relevantUsers->{list} }
                        ],
                        byrow => 1,
                    );

                }
                else {
                    push @{ $model->{optionLines} },
                      'Single coincidence correction factor';
                }

                my $red = Arithmetic(
                    name          => 'Contribution to first-band peak kW',
                    defaultFormat => '0softnz',
                    arithmetic    => $timebandLoadCoefficient
                    ? '=IV1*IV9*IV2/24/IV3*1000'
                    : '=IV1*IV2/24/IV3*1000',
                    rows      => $relevantUsers,
                    arguments => {
                        IV1 => $timebandLoadCoefficientAccording,
                        IV2 => $unitsByEndUser,
                        IV3 => $daysInYear,
                        $timebandLoadCoefficient
                        ? ( IV9 => $timebandLoadCoefficient )
                        : (),
                    },
                );

                my $coin = Arithmetic(
                    name          => 'Contribution to system-peak-time kW',
                    defaultFormat => '0softnz',
                    arithmetic    => '=IV1*IV2/24/IV3*1000',
                    rows          => $relevantUsers,
                    arguments     => {
                        IV1 => $loadCoefficients,
                        IV2 => $unitsByEndUser,
                        IV3 => $daysInYear,
                    },
                );

                $timebandLoadCoefficientAccording->{dontcolumnset} = 1;

                push @{ $model->{timeOfDayResults} },
                  Columnset(
                    name    => 'Estimated contributions to peak demand',
                    columns => [
                        $timebandLoadCoefficientAccording->{rows} ==
                          $red->{rows} ? $timebandLoadCoefficientAccording : (),
                        $red,
                        $coin,
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
                                arithmetic => '=IF(IV1,IV2/IV3,0)',
                                rows       => 0,
                                arguments  => {
                                    IV1 => $redG,
                                    IV2 => $coinG,
                                    IV3 => $redG,
                                }
                            ),
                        ),
                        $timebandLoadCoefficientAdjusted,
                    ]
                );

                $timebandLoadCoefficientAdjusted =
                  $timebandLoadCoefficientAdjusted->{sources}[0]
                  if $timebandLoadCoefficientAdjusted->lastRow ==
                  $timebandLoadCoefficientAdjusted->{sources}[0]
                  ->lastRow;    # hacky

            }
        }

        if (   $model->{coincidenceAdj}
            && $model->{coincidenceAdj} =~ /redonly/i )
        {
            $timebandLoadCoefficientAdjusted = Stack(
                name    => 'Rescaled time band load coefficient to be applied',
                rows    => $relevantEndUsersByRate[0],
                cols    => $timebandSet,
                sources => [
                    $timebandLoadCoefficientAdjusted,
                    Constant(
                        name => '1 for non-red',
                        rows => $relevantEndUsersByRate[0],
                        cols => $timebandSet,
                        data => [
                            map {
                                [ map { 1 }
                                      @{ $relevantEndUsersByRate[0]{list} } ]
                            } @{ $timebandSet->{list} }
                        ]
                    )
                ]
            );
            Columnset(
                name => 'Calculation of adjusted time band load coefficients',
                columns => [
                    $timebandLoadCoefficientAccording,
                    $timebandLoadCoefficientAdjusted->{sources}[0],
                    $timebandLoadCoefficientAdjusted,
                ]
              )
              unless $model->{coincidenceAdj}
              && $model->{coincidenceAdj} =~ /group/i;
        }
        else {
            Columnset(
                name => 'Calculation of adjusted time band load coefficients',
                columns => [
                    $timebandLoadCoefficientAccording,
                    $timebandLoadCoefficientAdjusted
                ]
            ) unless $timebandLoadCoefficientAccording->{dontcolumnset};
        }

        $pseudoLoadCoefficientBreakdown = Arithmetic(
            name => 'Pseudo load coefficient by time band and network level',
            rows => $relevantEndUsersByRate[0],
            cols => $networkLevelsTimeband,
            arithmetic => '=IF(IV6>0,IV2*IV7*24*IV9/IV5,0)',
            arguments  => {
                IV2 => $timebandLoadCoefficientAdjusted,
                IV5 => $annualHoursByTimeband,
                IV6 => $annualHoursByTimeband,
                IV7 => $peakingProbability,
                IV9 => $daysInYear,
            }
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

    \@pseudoLoadCoefficientsAgainstSystemPeak, \@pseudoLoadCoefficients;

}

1;
