﻿package CDCM;

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

sub networkUse {

    my (
        $model,                              $networkLevels,
        $drmLevels,                          $drmExitLevels,
        $operatingLevels,                    $operatingDrmExitLevels,
        $unitsInYear,                        $loadCoefficients,
        $pseudoLoadCoefficients,             $daysInYear,
        $lineLossFactors,                    $allTariffsByEndUser,
        $componentMap,                       $routeingFactors,
        $volumeData,                         $fFactors,
        $generationCapacityTariffsByEndUser, $modelGrossAssetsByLevel,
        $modelCostToSml,                     $modelSml,
        $serviceModelAssetsPerCustomer,      $serviceModelAssetsPerAnnualMwh
    ) = @_;

    my $simultaneousMaximumLoadUnits = Arithmetic(
        name => Label(
                'Estimated contributions of users on each tariff to '
              . 'system simultaneous maximum load by network level (kW)'
        ),
        arithmetic => '=A1*A3*A5/(24*A9)*1000',
        cols       => $drmExitLevels,
        arguments  => {
            A1 => $unitsInYear,
            A3 => $loadCoefficients,
            A9 => $daysInYear,
            A5 => $lineLossFactors
        },
        defaultFormat => '0softnz',
    );

    push @{ $model->{forecastSml} },
      Columnset(
        name => 'Conversion of demand forecast to estimated'
          . ' contribution to system simultaneous maximum load',
        columns => [ $unitsInYear, $simultaneousMaximumLoadUnits ]
      ) unless $model->{pcd};

    if ($pseudoLoadCoefficients) {

        $simultaneousMaximumLoadUnits = Stack(
            name => Label(
                    'Contributions of users on each tariff to '
                  . 'system simultaneous maximum load by network level (kW)'
            ),
            rows    => $allTariffsByEndUser,
            cols    => $drmExitLevels,
            sources => [
                (
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
                            grep { !$componentMap->{$_}{"Unit rate $r2 p/kWh"} }
                            @{ $pseudoLoadCoefficients->[$_]{rows}{list} }
                          )
                        {
                            my $relevantTariffs = Labelset(
                                name => "$rt-rate tariffs by end user",
                                $allTariffsByEndUser->{groups}
                                ? 'groups'
                                : 'list' => \@relevant
                            );

                            Arithmetic(
                                name => 'Contributions of users on '
                                  . $rt
                                  . '-rate multi tariffs to '
                                  . 'system simultaneous maximum load by network level (kW)',
                                arithmetic => '=('
                                  . join( '+', map { "A1$_*A3$_" } 0 .. $_ )
                                  . ')*A5/(24*A9)*1000',
                                rows      => $relevantTariffs,
                                cols      => $drmExitLevels,
                                arguments => {
                                    A9 => $daysInYear,
                                    A5 => $lineLossFactors,
                                    map {
                                        ;
                                        "A1$_" => $volumeData->{ 'Unit rate '
                                              . ( $_ + 1 )
                                              . ' p/kWh' },
                                          "A3$_" =>
                                          $pseudoLoadCoefficients->[$_];
                                    } 0 .. $_
                                },
                                defaultFormat => '0softnz'
                            );
                        }
                        else { (); }
                    } 0 .. $model->{maxUnitRates} - 1
                ),
                $simultaneousMaximumLoadUnits
            ],
            defaultFormat => '0copynz'
        );

    }

    my $forecastSml = GroupBy(
        name =>
          'Forecast system simultaneous maximum load (kW) from forecast units',
        rows          => 0,
        cols          => $drmExitLevels,
        source        => $simultaneousMaximumLoadUnits,
        defaultFormat => '0softnz',
    );

    my $simultaneousMaximumLoadCapacity;

    if ($fFactors) {

        push @{ $model->{forecastSml} }, $forecastSml;

        $simultaneousMaximumLoadCapacity = Arithmetic(
            name => Label(
                    'Contributions of users on generation capacity rates '
                  . 'to simultaneous maximum load by network level (kW)'
            ),
            arithmetic => '=-1*A1*A3*A5',
            cols       => $drmExitLevels,
            rows       => $generationCapacityTariffsByEndUser,
            arguments  => {
                A1 => $volumeData->{'Generation capacity rate p/kW/day'},
                A3 => $fFactors,
                A5 => $lineLossFactors
            }
        );

        $forecastSml = Arithmetic
          name       => 'Forecast system simultaneous maximum load (kW)',
          arithmetic => '=A1+A2',
          arguments  => {
            A1 => $forecastSml,
            A2 => GroupBy(
                name => 'Adjustment to simultaneous maximum load'
                  . ' from users on generation capacity rates (kW)',
                cols   => $drmExitLevels,
                source => $simultaneousMaximumLoadCapacity
            )
          };

    }

    push @{ $model->{forecastSml} }, $forecastSml;

    push @{ $model->{edcmTables} },
      Stack(
        name => 'EDCM input data ⇒1122. Forecast system '
          . 'simultaneous maximum load (kW) from CDCM users',
        singleRowName => 'Forecast system simultaneous maximum load',
        cols =>
          Labelset( list => [ @{ $forecastSml->{cols}{list} }[ 0 .. 5 ] ] ),
        sources => [$forecastSml]
      ) if $model->{edcmTables};

    $forecastSml, $simultaneousMaximumLoadUnits,
      $simultaneousMaximumLoadCapacity;

}

1;
