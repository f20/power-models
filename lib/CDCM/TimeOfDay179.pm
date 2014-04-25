package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others.
Copyright 2011-2014 Franck Latrémolière, Reckon LLP and others.

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
        $model->{agghhequalisation} ? 'equalisation' : 'none',
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
        arithmetic    => '=IV1*24*IV3/IV2',
        arguments     => {
            IV2 => $annualHoursByTimebandTotal,
            IV3 => $daysInYear,
            IV1 => $annualHoursByTimebandRaw,
        }
    );

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
            name => 'Normalisation of '
              . $blackYellowGreen
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
        arithmetic => '=IF(IV6>0,IV7*24*IV9/IV5,0)*IF(IV2<0,-1,1)',
        arguments  => {
            IV5 => $annualHoursByTimeband,
            IV6 => $annualHoursByTimeband,
            IV7 => $peakingProbability,
            IV9 => $daysInYear,
            IV2 => $loadCoefficients,
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
                                arithmetic => '=IF(IV6>0,IV1*IV4*24/IV5,0)',
                                cols       => $timebandSet,
                                rows       => $relevantUsers,
                                arguments  => {
                                    IV1 => $timebandUseByTariff[$_],
                                    IV4 => $daysInYear,
                                    IV5 => $annualHoursByTimeband,
                                    IV6 => $annualHoursByTimeband
                                }
                            )
                          )
                          : Arithmetic(
                            name => 'Peak band '
                              . $blackYellowGreen
                              . 'load coefficient for '
                              . $rt
                              . '-rate tariffs',
                            arithmetic => '=IF(IV6>0,IV1*IV4*24/IV5,0)',
                            cols       => $peakBand,
                            rows       => $relevantUsers,
                            arguments  => {
                                IV1 => $timebandUseByTariff[$_],
                                IV4 => $daysInYear,
                                IV5 => $annualHoursByTimeband,
                                IV6 => $annualHoursByTimeband
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
            arithmetic => '=IF(IV5<>0,IV4/IV2,IF(IV8<0,-1,1))',
            rows       => $relevantEndUsersByRate[0],
            $model->{timebandCoef} && $model->{timebandCoef} =~ /detail/i
            ? ( cols => $networkLevelsTimeband )
            : (      $model->{coincidenceAdj}
                  && $model->{coincidenceAdj} =~ /redonly/i )
            ? ( cols => $peakBand )
            : (),
            arguments => {
                IV2 => $timebandLoadCoefficientAccording,
                IV5 => $timebandLoadCoefficientAccording,
                IV4 => $loadCoefficients,
                IV8 => $loadCoefficients,
            }
        );

        if ( $correctionRules =~ /groupAll/i ) {

            my $relevantUsers = $relevantEndUsersByRate[0];

            my $red = Arithmetic(
                name          => 'Contribution to peak band kW',
                defaultFormat => '0softnz',
                arithmetic    => '=IV1*IV2/24/IV3*1000',
                rows          => $relevantUsers,
                arguments     => {
                    IV1 => $timebandLoadCoefficientAccording,
                    IV2 => $unitsByEndUser,
                    IV3 => $daysInYear,
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
                columns => [ $timebandLoadCoefficientAccording, $red, $coin, ]
              );

            $timebandLoadCoefficientAdjusted = Arithmetic(
                name => 'Load coefficient correction factor for the group',
                arithmetic => '=IF(SUM(IV1_IV5),SUM(IV2_IV6)/SUM(IV3_IV7),0)',
                rows       => 0,
                arguments  => {
                    IV1_IV5 => $red,
                    IV2_IV6 => $coin,
                    IV3_IV7 => $red,
                }
            );

        }

        Columnset(
            name    => 'Calculation of adjusted time band load coefficients',
            columns => [
                $timebandLoadCoefficientAccording,
                $timebandLoadCoefficientAdjusted
            ]
        ) unless $timebandLoadCoefficientAccording->{dontcolumnset};

        $pseudoLoadCoefficientBreakdown = Arithmetic(
            name => 'Pseudo load coefficient by time band and network level',
            $correctionRules =~ /groupAll/i
            ? ()
            : ( rows => $relevantEndUsersByRate[0] ),
            cols       => $networkLevelsTimeband,
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
                      && $componentMap->{$_}{'Unit rates p/kWh'}
                      && !$componentMap->{$_}{'Capacity charge p/kVA/day'}
                } @{ $allEndUsers->{list} }
            ]
        );

        foreach ( values %{ $eq{userSet} } ) {
            die 'Mismatch' if @{ $_->{list} } != @{ $groupset->{list} };
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
                ? '=(IV1*IV2+IV3*IV4)/(IV7+IV8)'
                : '=(IV1*IV2+IV3*IV4+IV5*IV6)/(IV7+IV8+IV9)',
                arguments => {
                    IV1 => $eq{units}{NHH1},
                    IV2 => $eq{timebandUse}{NHH1},
                    IV3 => $eq{units}{NHH2},
                    IV4 => $eq{timebandUse}{NHH2},
                    IV7 => $eq{units}{NHH1},
                    IV8 => $eq{units}{NHH2},
                    $model->{agghhequalisation} =~ /nooffpeak/i ? ()
                    : (
                        IV5 => $eq{units}{NHH0},
                        IV6 => $eq{timebandUse}{NHH0},
                        IV9 => $eq{units}{NHH0}
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
            ? '=(IV1*IV2+IV3*IV4)/(IV7+IV8)'
            : '=(IV1*IV2+IV3*IV4+IV5*IV6)/(IV7+IV8+IV9)',
            arguments => {
                IV1 => $eq{units}{NHH1},
                IV2 => $eq{tariffCoef}{NHH1},
                IV3 => $eq{units}{NHH2},
                IV4 => $eq{tariffCoef}{NHH2},
                IV7 => $eq{units}{NHH1},
                IV8 => $eq{units}{NHH2},
                $model->{agghhequalisation} =~ /nooffpeak/i ? ()
                : (
                    IV5 => $eq{units}{NHH0},
                    IV6 => $eq{tariffCoef}{NHH0},
                    IV9 => $eq{units}{NHH0},
                ),
            },
        );

        $eq{relCorr} = Arithmetic(
            name =>
              'Relative correction factor for aggregated half hourly tariff',
            rows       => $groupset,
            cols       => $networkLevelsTimebandAware,
            arithmetic => '=IV1/IV9',
            arguments  => {
                IV1 => $eq{tariffCoef}{NHH},
                IV9 => $eq{tariffCoef}{hybrid},
            },
        );

        $eq{nhhCorr} = Arithmetic(
            name       => 'Correction factor for non half hourly tariffs',
            rows       => $groupset,
            cols       => $networkLevelsTimebandAware,
            arithmetic => $model->{agghhequalisation} =~ /nooffpeak/i
            ? '=(IV11*IV21+IV31*IV41+IV71*IV81)'
              . '/(IV12*IV22+IV32*IV42+IV72*IV82*IV9)'
            : '=(IV11*IV21+IV31*IV41+IV51*IV61+IV71*IV81)'
              . '/(IV12*IV22+IV32*IV42+IV52*IV62+IV72*IV82*IV9)',
            arguments => {
                IV11 => $eq{units}{NHH1},
                IV12 => $eq{units}{NHH1},
                IV21 => $eq{tariffCoef}{NHH1},
                IV22 => $eq{tariffCoef}{NHH1},
                IV31 => $eq{units}{NHH2},
                IV32 => $eq{units}{NHH2},
                IV41 => $eq{tariffCoef}{NHH2},
                IV42 => $eq{tariffCoef}{NHH2},
                IV71 => $eq{units}{AHH},
                IV72 => $eq{units}{AHH},
                IV81 => $eq{tariffCoef}{AHH},
                IV82 => $eq{tariffCoef}{AHH},
                IV9  => $eq{relCorr},
                $model->{agghhequalisation} =~ /nooffpeak/i ? ()
                : (
                    IV51 => $eq{units}{NHH0},
                    IV52 => $eq{units}{NHH0},
                    IV61 => $eq{tariffCoef}{NHH0},
                    IV62 => $eq{tariffCoef}{NHH0},
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
                            arithmetic => '=IV1*IV2',
                            arguments  => {
                                IV1 => $eq{timebandCoef}{$_},
                                IV2 => $eq{nhhCorr},
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
                    arithmetic => '=IV1*IV2*IV3',
                    arguments  => {
                        IV1 => $eq{timebandCoef}{AHH},
                        IV2 => $eq{nhhCorr},
                        IV3 => $eq{relCorr},
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
