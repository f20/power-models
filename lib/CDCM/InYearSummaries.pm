﻿package CDCM;

# Copyright 2009-2011 Energy Networks Association Limited and others.
# Copyright 2011-2012 Franck Latrémolière, Reckon LLP and others.
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

=head Development note

For some configurations, this uses $model->{inYear...} to get data from InYearAdjust behind the back of Master:
    $model->{inYearVolumes}  = [ $volumeDataBefore1, $volumeDataBefore2, ];
    $model->{inYearTariffs}  = [ $tariffsBefore1,    $tariffsBefore2, ];
    $model->{inYearRevenues} = [ $revenuesBefore1,   $revenuesBefore2, ];

=cut

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Shortcuts ':all';

sub summaryOfRevenuesHybrid {

    my (
        $model,        $allTariffs,       $nonExcludedComponents,
        $componentMap, $volumeDataAfter,  $daysBefore,
        $daysAfter,    $unitsInYearAfter, $tariffTable,
        $volumeData,   $daysInYear,       $unitsInYear,
    ) = @_;

    my @tariffs  = ( @{ $model->{inYearTariffs} }, $tariffTable );
    my $numParts = @tariffs;
    my @days     = (
        ref $daysBefore eq 'ARRAY' ? @$daysBefore : $daysBefore,
        $daysAfter, $daysInYear
    );
    my @volumes =
      ( @{ $model->{inYearVolumes} }, $volumeDataAfter, $volumeData );
    my @revenues =
      $model->{inYearRevenues} ? @{ $model->{inYearRevenues} } : ();
    my @units = (
        (
            map {
                my $part = $_;
                Arithmetic(
                    name => Label(
                        'All units (MWh)',
                        'All units aggregated by tariff (MWh)'
                    ),
                    noCopy     => 1,
                    arithmetic => '='
                      . join( '+', map { "A$_" } 1 .. $model->{maxUnitRates} ),
                    arguments => {
                        map {
                            ( "A$_" =>
                                  ${ $volumes[$part] }{"Unit rate $_ p/kWh"} )
                        } 1 .. $model->{maxUnitRates}
                    },
                    defaultFormat => '0softnz',
                  )
            } 0 .. ( $numParts - 2 )
        ),
        $unitsInYearAfter,
        $unitsInYear
    );
    my @revenuesFromTariffs;
    my @revenuesFromUnitRates;
    my @revenuesFromUnits2d;
    my @revenuesFromFixedCharges;
    my @revenuesFromCapacityCharges;
    my @revenuesFromUnauthDemandCharges;
    my @revenuesFromReactiveCharges;

    foreach my $part ( 0 .. $numParts ) {

        my $revenuesFromTariffs;
        my $revenuesFromUnitRates;
        my @revenuesFromUnits;
        my $revenuesFromFixedCharges;
        my $revenuesFromCapacityCharges;
        my $revenuesFromUnauthDemandCharges;
        my $revenuesFromReactiveCharges;

        my $volumeData  = $volumes[$part];
        my $daysInYear  = $days[$part];
        my $unitsInYear = $units[$part];

        if ( my $tariffTable = $tariffs[$part] ) {

            {
                my @termsNoDays;
                my @termsUnits;
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
                        push @termsUnits, "A2$pad*A3$pad" if /kWh/;
                    }
                    $args{"A2$pad"} = $tariffTable->{$_};
                    $args{"A3$pad"} = $volumeData->{$_};
                }
                $revenuesFromTariffs = $revenues[$part]
                  || Arithmetic(
                    name       => 'Net revenues (£)',
                    rows       => $allTariffs,
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
                  );
                $revenuesFromUnitRates = Arithmetic(
                    name          => 'Revenues from unit rates (£)',
                    defaultFormat => '0softnz',
                    rows          => $allTariffs,
                    arithmetic    => '=10*(' . join( '+', @termsUnits ) . ')',
                    arguments     => \%args,
                );

            }

            @revenuesFromUnits = map {
                Arithmetic(
                    name          => "Net revenues from unit rate $_ (£)",
                    defaultFormat => '0softnz',
                    arithmetic    => '=A1*A2*10',
                    arguments     => {
                        A2 => $volumeData->{"Unit rate $_ p/kWh"},
                        A1 => $tariffTable->{"Unit rate $_ p/kWh"},
                    }
                );
            } 1 .. $model->{maxUnitRates};

            $revenuesFromFixedCharges = Arithmetic(
                name          => 'Revenues from fixed charges (£)',
                defaultFormat => '0softnz',
                arithmetic    => '=A1*A4*A2/100',
                arguments     => {
                    A2 => $volumeData->{'Fixed charge p/MPAN/day'},
                    A4 => $daysInYear,
                    A1 => $tariffTable->{'Fixed charge p/MPAN/day'},
                }
            );

            $revenuesFromCapacityCharges = Arithmetic(
                name          => 'Revenues from capacity charges (£)',
                defaultFormat => '0softnz',
                arithmetic    => '=A1*A4*A2/100',
                arguments     => {
                    A2 => $volumeData->{'Capacity charge p/kVA/day'},
                    A4 => $daysInYear,
                    A1 => $tariffTable->{'Capacity charge p/kVA/day'},
                }
            ) if $volumeData->{'Capacity charge p/kVA/day'};

            $revenuesFromUnauthDemandCharges =
              $model->{unauth} && $model->{unauth} =~ /day/i
              ? Arithmetic(
                name          => 'Revenues from exceeded capacity charges (£)',
                defaultFormat => '0softnz',
                arithmetic    => '=A1*A4*A2/100',
                arguments     => {
                    A2 => $volumeData->{'Exceeded capacity charge p/kVA/day'},
                    A4 => $daysInYear,
                    A1 => $tariffTable->{'Exceeded capacity charge p/kVA/day'},
                }
              )
              : Arithmetic(
                name => 'Revenues from unauthorised demand charges (£)',
                defaultFormat => '0softnz',
                arithmetic    => '=A1*A2*10',
                arguments     => {
                    A2 => $volumeData->{'Unauthorised demand charge p/kVAh'},
                    A1 => $tariffTable->{'Unauthorised demand charge p/kVAh'},
                }
              ) if $model->{unauth};

            $revenuesFromReactiveCharges = Arithmetic(
                name          => 'Revenues from reactive power charges (£)',
                defaultFormat => '0softnz',
                arithmetic    => '=A1*A2*10',
                arguments     => {
                    A2 => $volumeData->{'Reactive power charge p/kVArh'},
                    A1 => $tariffTable->{'Reactive power charge p/kVArh'},
                }
              )
              if !$model->{reactiveExcluded}
              && $volumeData->{'Reactive power charge p/kVArh'};

        }

        else {    # aggregation of previous parts into tariff bits

            my $sum = '=' . join( '+', map { "A$_" } 1 .. $numParts );

            $revenuesFromTariffs = Arithmetic(
                name          => 'Aggregate revenues (£/year)',
                defaultFormat => '0softnz',
                arithmetic    => $sum,
                arguments     => {
                    map { ( "A$_" => $revenuesFromTariffs[ $_ - 1 ] ); }
                      1 .. $numParts
                },
            );

            $revenuesFromUnitRates = Arithmetic(
                name          => 'Aggregate revenues from unit rates (£/year)',
                defaultFormat => '0softnz',
                arithmetic    => $sum,
                arguments     => {
                    map { ( "A$_" => $revenuesFromUnitRates[ $_ - 1 ] ) }
                      1 .. $numParts
                },
            );

            @revenuesFromUnits = map {
                my $rate = $_;
                Arithmetic(
                    name => 'Aggregate revenues from units '
                      . $rate
                      . ' (£/year)',
                    defaultFormat => '0softnz',
                    arithmetic    => $sum,
                    arguments     => {
                        map {
                            ( "A$_" =>
                                  $revenuesFromUnits2d[ $_ - 1 ][ $rate - 1 ] )
                        } 1 .. $numParts
                    },
                  )
            } 1 .. $model->{maxUnitRates};

            $revenuesFromFixedCharges = Arithmetic(
                name => 'Aggregate revenues from fixed charges (£/year)',
                defaultFormat => '0softnz',
                arithmetic    => $sum,
                arguments     => {
                    map { ( "A$_" => $revenuesFromFixedCharges[ $_ - 1 ] ) }
                      1 .. $numParts
                },
            );

            $revenuesFromCapacityCharges = Arithmetic(
                name => 'Aggregate revenues from capacity charges (£/year)',
                defaultFormat => '0softnz',
                arithmetic    => $sum,
                arguments     => {
                    map { ( "A$_" => $revenuesFromCapacityCharges[ $_ - 1 ] ) }
                      1 .. $numParts
                },
            ) if $revenuesFromCapacityCharges[0];

            $revenuesFromUnauthDemandCharges = Arithmetic(
                name =>
                  'Aggregate revenues from excess demand charges (£/year)',
                defaultFormat => '0softnz',
                arithmetic    => $sum,
                arguments     => {
                    map {
                        ( "A$_" => $revenuesFromUnauthDemandCharges[ $_ - 1 ] )
                    } 1 .. $numParts
                },
            ) if $revenuesFromUnauthDemandCharges[0];

            $revenuesFromReactiveCharges = Arithmetic(
                name =>
                  'Aggregate revenues from excess reactive charges (£/year)',
                defaultFormat => '0softnz',
                arithmetic    => $sum,
                arguments     => {
                    map { ( "A$_" => $revenuesFromReactiveCharges[ $_ - 1 ] ) }
                      1 .. $numParts
                },
            ) if $revenuesFromReactiveCharges[0];

        }

        $revenuesFromTariffs[$part]         = $revenuesFromTariffs;
        $revenuesFromUnitRates[$part]       = $revenuesFromUnitRates;
        $revenuesFromUnits2d[$part]         = \@revenuesFromUnits;
        $revenuesFromFixedCharges[$part]    = $revenuesFromFixedCharges;
        $revenuesFromCapacityCharges[$part] = $revenuesFromCapacityCharges;
        $revenuesFromUnauthDemandCharges[$part] =
          $revenuesFromUnauthDemandCharges;
        $revenuesFromReactiveCharges[$part] = $revenuesFromReactiveCharges;

        my $myUnits =
            $unitsInYear->{noCopy} && !$unitsInYear->{location}
          ? $unitsInYear
          : Stack(
            rows    => $allTariffs,
            sources => [$unitsInYear]
          );

        my $averageUnitRate = Arithmetic(
            name       => Label('Average unit rate p/kWh'),
            rows       => $allTariffs,
            arithmetic => '=IF(A403<>0,0.1*A1/A402,0)',
            arguments  => {
                A1   => $revenuesFromUnitRates,
                A402 => $myUnits,
                A403 => $myUnits,
            },
        );

        my @unitProportion = $revenuesFromUnitRates
          ? map {
            Arithmetic(
                name          => "Rate $_ revenue proportion",
                defaultFormat => '%softnz',
                arithmetic    => '=IF(A3<>0,A5/A1,"")',
                arguments     => {
                    A1 => $revenuesFromUnitRates,
                    A3 => $revenuesFromUnitRates,
                    A5 => $revenuesFromUnits[ $_ - 1 ],
                }
            );
          } 1 .. $model->{maxUnitRates}
          : ();

        my $fixedProportion = Arithmetic(
            name          => 'Fixed charge proportion',
            defaultFormat => '%softnz',
            arithmetic    => '=IF(A3<>0,A5/A1,"")',
            arguments     => {
                A1 => $revenuesFromTariffs,
                A3 => $revenuesFromTariffs,
                A5 => $revenuesFromFixedCharges,
            }
        );

        my $capacityProportion = Arithmetic(
            name          => 'Capacity charge proportion',
            defaultFormat => '%softnz',
            arithmetic    => '=IF(A3<>0,A5/A1,"")',
            arguments     => {
                A1 => $revenuesFromTariffs,
                A3 => $revenuesFromTariffs,
                A5 => $revenuesFromCapacityCharges,
            }
        ) if $revenuesFromCapacityCharges;

        my $unauthProportion;
        $unauthProportion = Arithmetic(
            name => $model->{unauth} && $model->{unauth} =~ /day/i
            ? 'Exceeded capacity charge proportion'
            : 'Unauthorised demand charge proportion',
            defaultFormat => '%softnz',
            arithmetic    => '=IF(A3<>0,A5/A1,"")',
            arguments     => {
                A1 => $revenuesFromTariffs,
                A3 => $revenuesFromTariffs,
                A5 => $revenuesFromUnauthDemandCharges,
            }
        ) if $model->{unauth};

        my $reactiveProportion;
        $reactiveProportion = Arithmetic(
            name          => 'Reactive power charge proportion',
            defaultFormat => '%softnz',
            arithmetic    => '=IF(A3<>0,A5/A1,"")',
            arguments     => {
                A1 => $revenuesFromTariffs,
                A3 => $revenuesFromTariffs,
                A5 => $revenuesFromReactiveCharges,
            }
        ) if !$model->{reactiveExcluded} && $revenuesFromReactiveCharges;

        my $averageByUnit = Arithmetic(
            name       => 'Average p/kWh',
            arithmetic => '=IF(A3<>0,0.1*A1/A2,"")',
            arguments  => {
                A1 => $revenuesFromTariffs,
                A2 => $myUnits,
                A3 => $myUnits,
            }
        );

        my $myMpans = Stack(
            rows    => $allTariffs,
            sources => [ $volumeData->{'Fixed charge p/MPAN/day'} ],
        );

        my $averageByMpan = Arithmetic(
            name          => 'Average £/MPAN',
            defaultFormat => '0.00softnz',
            arithmetic    => '=IF(A3<>0,A1/A2,"")',
            arguments     => {
                A1 => $revenuesFromTariffs,
                A2 => $myMpans,
                A3 => $myMpans,
            }
        );

        my $partName =
            $part == $numParts     ? ' (aggregate)'
          : $part == $numParts - 1 ? ' (after tariff change)'
          :                          ( ' (period ' . ( $part + 1 ) . ')' );

        push @{ $model->{overallSummary} }, Columnset(
            name    => 'Revenue summary' . $partName,
            columns => [

                $myUnits,

                $myMpans,

                $revenuesFromTariffs,

                $revenuesFromUnitRates ? $revenuesFromUnitRates : (),

                $revenuesFromFixedCharges ? $revenuesFromFixedCharges : (),

                $revenuesFromCapacityCharges
                ? $revenuesFromCapacityCharges
                : (),

                $revenuesFromUnauthDemandCharges
                ? $revenuesFromUnauthDemandCharges
                : (),

                $revenuesFromReactiveCharges
                ? $revenuesFromReactiveCharges
                : (),

                $averageByUnit,

                $averageByMpan,

                1 ? () : Arithmetic(
                    name          => 'Average p/MPAN/day',
                    defaultFormat => '0.00softnz',
                    arithmetic    => '=IF(A3<>0,A1/A2*100/A4,"")',
                    arguments     => {
                        A1 => $revenuesFromTariffs,
                        A2 => $volumeData->{'Fixed charge p/MPAN/day'},
                        A3 => $volumeData->{'Fixed charge p/MPAN/day'},
                        A4 => $daysInYear,
                    }
                ),

                1 ? () : Arithmetic(
                    name          => 'Average p/kVA/day',
                    defaultFormat => '0.00softnz',
                    arithmetic    => '=IF(A3<>0,A1/A2*100/A4,"")',
                    arguments     => {
                        A1 => $revenuesFromTariffs,
                        A2 => $volumeData->{'Capacity charge p/kVA/day'},
                        A3 => $volumeData->{'Capacity charge p/kVA/day'},
                        A4 => $daysInYear,
                    }
                ),

                0 ? () : $averageUnitRate,

                1 ? () : (
                    map {
                        Stack(
                            sources       => [$_],
                            defaultFormat => !$_->{defaultFormat}
                              || $_->{defaultFormat} =~ /000/
                            ? '0.000copynz'
                            : '0.00copynz'
                        );
                    } @{ $model->{allTariffColumns} }
                ),

                @unitProportion ? ( @revenuesFromUnits, @unitProportion ) : (),

                $fixedProportion ? $fixedProportion : (),

                $capacityProportion ? $capacityProportion : (),

                $unauthProportion ? $unauthProportion : (),

                $reactiveProportion ? $reactiveProportion : (),

                1 ? () : Arithmetic(
                    name          => 'Fixed charge proportion',
                    defaultFormat => '%softnz',
                    arithmetic    => '=IF(A3<>0,A5/(A1/A2*100/A4),"")',
                    arguments     => {
                        A1 => $revenuesFromTariffs,
                        A2 => $volumeData->{'Fixed charge p/MPAN/day'},
                        A3 => $volumeData->{'Fixed charge p/MPAN/day'},
                        A4 => $daysInYear,
                        A5 => $tariffTable->{'Fixed charge p/MPAN/day'},
                    }
                ),

                1 ? () : Arithmetic(
                    name          => 'Capacity charge proportion',
                    defaultFormat => '%softnz',
                    arithmetic    => '=IF(A3<>0,A5/(A1/A2*100/A4),"")',
                    arguments     => {
                        A1 => $revenuesFromTariffs,
                        A2 => $volumeData->{'Capacity charge p/kVA/day'},
                        A3 => $volumeData->{'Capacity charge p/kVA/day'},
                        A4 => $daysInYear,
                        A5 => $tariffTable->{'Capacity charge p/kVA/day'},
                    }
                ),

              ]

        );

        my $totalUnits = GroupBy(
            name          => 'Total units (MWh)',
            rows          => 0,
            cols          => 0,
            defaultFormat => '0soft',
            source        => $myUnits,
        );

        my $totalMpans = GroupBy(
            name          => 'Total MPANs',
            rows          => 0,
            cols          => 0,
            source        => $myMpans,
            defaultFormat => '0soft'
        );

        my $totalRevenuesFromTariffs = GroupBy(
            name          => 'Total net revenues (£)',
            rows          => 0,
            cols          => 0,
            source        => $revenuesFromTariffs,
            defaultFormat => '0soft'
        );

        my $totalRevenuesFromUnitRates = GroupBy(
            name          => 'Total net revenues from unit rates (£)',
            rows          => 0,
            cols          => 0,
            source        => $revenuesFromUnitRates,
            defaultFormat => '0soft'
        );

        my $totalRevenuesFromFixed = GroupBy(
            name          => 'Total revenues from fixed charges (£)',
            defaultFormat => '0soft',
            source        => $revenuesFromFixedCharges,
        );

        my $totalFixedProportionSilly = Arithmetic(
            name          => 'Fixed charge proportion',
            defaultFormat => '%softnz',
            arithmetic    => '=IF(A3<>0,A5/A1,"")',
            arguments     => {
                A1 => $totalRevenuesFromTariffs,
                A3 => $totalRevenuesFromTariffs,
                A5 => $totalRevenuesFromFixed,
            }
        );

        my $totalRevenuesFromCapacity = GroupBy(
            name          => 'Total revenues from capacity charges (£)',
            defaultFormat => '0soft',
            source        => $revenuesFromCapacityCharges,
        ) if $revenuesFromCapacityCharges;

        my $totalRevenuesFromUnauth =
          $revenuesFromUnauthDemandCharges
          ? GroupBy(
            name => 'Total revenues from '
              . (
                $model->{unauth} && $model->{unauth} =~ /day/i
                ? 'exceeded capacity'
                : 'unauthorised demand'
              )
              . ' charges (£)',
            defaultFormat => '0soft',
            source        => $revenuesFromUnauthDemandCharges,
          )
          : undef;

        my $totalRevenuesFromReactive =
          $revenuesFromReactiveCharges
          ? GroupBy(
            name          => 'Total revenues from reactive power charges (£)',
            defaultFormat => '0soft',
            source        => $revenuesFromReactiveCharges,
          )
          : undef;

        push @{ $model->{overallSummary} },
          Columnset(
            name    => 'Revenue summary by tariff component' . $partName,
            columns => [
                $totalUnits,
                $totalMpans,
                $totalRevenuesFromTariffs,
                $totalRevenuesFromUnitRates,
                $totalRevenuesFromFixed,
                $totalRevenuesFromCapacity ? $totalRevenuesFromCapacity : (),
                $totalRevenuesFromUnauth   ? $totalRevenuesFromUnauth   : (),
                $totalRevenuesFromReactive ? $totalRevenuesFromReactive : (),
            ]
          );

    }

    @revenuesFromTariffs;

}

1;
