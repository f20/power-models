package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others.
Copyright 2011-2012 Franck Latrémolière, Reckon LLP and others.

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
        arithmetic => '=IV1*IV3*IV5/(24*IV9)*1000',
        cols       => $drmExitLevels,
        arguments  => {
            IV1 => $unitsInYear,
            IV3 => $loadCoefficients,
            IV9 => $daysInYear,
            IV5 => $lineLossFactors
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
                                  . join( '+', map { "IV1$_*IV3$_" } 0 .. $_ )
                                  . ')*IV5/(24*IV9)*1000',
                                rows      => $relevantTariffs,
                                cols      => $drmExitLevels,
                                arguments => {
                                    IV9 => $daysInYear,
                                    IV5 => $lineLossFactors,
                                    map {
                                        ;
                                        "IV1$_" => $volumeData->{ 'Unit rate '
                                              . ( $_ + 1 )
                                              . ' p/kWh' },
                                          "IV3$_" =>
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
            arithmetic => '=-1*IV1*IV3*IV5',
            cols       => $drmExitLevels,
            rows       => $generationCapacityTariffsByEndUser,
            arguments  => {
                IV1 => $volumeData->{'Generation capacity rate p/kW/day'},
                IV3 => $fFactors,
                IV5 => $lineLossFactors
            }
        );

        $forecastSml = Arithmetic
          name       => 'Forecast system simultaneous maximum load (kW)',
          arithmetic => '=IV1+IV2',
          arguments  => {
            IV1 => $forecastSml,
            IV2 => GroupBy(
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
        name => 'Forecast system simultaneous maximum load (kW)'
          . ' from CDCM users',
        defaultFormat => '0hard',
        number        => 1122,
        cols =>
          Labelset( list => [ @{ $forecastSml->{cols}{list} }[ 0 .. 5 ] ] ),
        sources => [$forecastSml]
      ) if $model->{edcmTables};

    $forecastSml, $simultaneousMaximumLoadUnits,
      $simultaneousMaximumLoadCapacity;

}

1;
