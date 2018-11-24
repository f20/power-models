package CDCM;

# Copyright 2009-2011 Energy Networks Association Limited and others.
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

sub yardsticks {

    my (
        $model,
        $drmExitLevels,
        $chargingDrmExitLevels,
        $unitTariffsByEndUser,
        $generationCapacityTariffsByEndUsers,
        $allTariffsByEndUser,
        $loadCoefficients,
        $pseudoLoadCoefficients,
        $costToSml,
        $fFactors,
        $lineLossFactors,
        $proportionCoveredByContributions,
        $daysInYear,
    ) = @_;

    # This unrestricted yardstick is used in reactive power calculations
    # even if all active power tariffs are on a multi-rate basis.
    # NB: possible inaccuracies here under some coincidenceAdj options.
    my $yardstickUnitsComponents = Arithmetic(
        name => Label(
            'Pay-as-you-go p/kWh',
            'Pay-as-you-go yardstick unit costs by charging level (p/kWh)'
        ),
        defaultFormat => '0.000softnz',
        rows          => $unitTariffsByEndUser,
        cols          => $chargingDrmExitLevels,
        arithmetic    => '=A1*A2*A3*(1-A4)/(24*A5)*100',
        arguments     => {
            A1 => $costToSml,
            A2 => $loadCoefficients,
            A3 => $lineLossFactors,
            A4 => $proportionCoveredByContributions,
            A5 => $daysInYear
        }
    );

    my $paygUnitYardstick = GroupBy(
        name => Label(
            'Pay-as-you-go yardstick p/kWh',
            'Pay-as-you-go yardstick unit rate (p/kWh)'
        ),
        rows   => $unitTariffsByEndUser,
        source => $yardstickUnitsComponents
    );

    push @{ $model->{yardsticks} },
      $model->{showSums}
      ? Columnset(
        name    => 'Pay-as-you-go yardstick unit rate (p/kWh)',
        columns => [ $yardstickUnitsComponents, $paygUnitYardstick ]
      )
      : $yardstickUnitsComponents;

    my $yardstickCapacityRates;

    if ($fFactors) {

        my $yardstickCapacityComponents = Arithmetic(
            name => Label(
                'Component p/kW/day',
                'Yardstick capacity costs by charging level (p/kW/day)'
            ),
            defaultFormat => '0.000softnz',
            rows          => $generationCapacityTariffsByEndUsers,
            cols          => $chargingDrmExitLevels,
            arithmetic    => '=-100*A1*A2*A3*(1-A4)/A5',
            arguments     => {
                A1 => $costToSml,
                A2 => $fFactors,
                A3 => $lineLossFactors,
                A4 => $proportionCoveredByContributions,
                A5 => $daysInYear,
            }
        );

        $yardstickCapacityRates = GroupBy(
            name => Label(
                'Yardstick p/kW/day',
                'Yardstick capacity costs (p/kW/day)'
            ),
            rows   => $generationCapacityTariffsByEndUsers,
            source => $yardstickCapacityComponents
        );

        push @{ $model->{yardsticks} },
          Columnset(
            name    => 'Yardstick capacity costs (p/kW/day)',
            columns => [ $yardstickCapacityComponents, $yardstickCapacityRates ]
          );

    }

    my $chargingDrmExitLevelsTimebandAware;

    my @paygUnitRates =
      map {
        $chargingDrmExitLevelsTimebandAware ||= Labelset(
            name    => 'Charging levels (DRM and exit) timeband aware',
            list    => $chargingDrmExitLevels->{list},
            accepts => [
                $drmExitLevels, $chargingDrmExitLevels,
                $pseudoLoadCoefficients->[$_]{cols}
            ]
        );

        my $rates = GroupBy(
            name => 'Pay-as-you-go unit rate ' . ( 1 + $_ ) . ' (p/kWh)',
            rows   => $pseudoLoadCoefficients->[$_]{tariffs},
            source => Arithmetic(
                name => 'Contributions to pay-as-you-go unit rate '
                  . ( 1 + $_ )
                  . ' (p/kWh)',
                arithmetic => '=A1*A2*A3*(1-A4)*100/(24*A6)',
                cols       => $chargingDrmExitLevelsTimebandAware,
                rows       => $pseudoLoadCoefficients->[$_]{tariffs},
                arguments  => {
                    A1 => $pseudoLoadCoefficients->[$_],
                    A2 => $costToSml,
                    A4 => $proportionCoveredByContributions,
                    A3 => $lineLossFactors,
                    A6 => $daysInYear,
                }
            )
        );

        push @{ $model->{yardsticks} },
          $model->{showSums}
          ? Columnset(
            name => 'Pay-as-you-go unit rate ' . ( 1 + $_ ) . ' p/kWh',
            columns => [ $rates->{source}, $rates ]
          )
          : $rates->{source};

        $rates;

      } 0 .. $model->{maxUnitRates} - 1 if $pseudoLoadCoefficients;

    $yardstickCapacityRates, $paygUnitYardstick, \@paygUnitRates;

}

1;
