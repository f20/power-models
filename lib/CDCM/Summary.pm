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

sub summaryOfRevenues {

    my ( $model, $tariffTable, $volumeData, $daysInYear, $nonExcludedComponents,
        $componentMap, $allTariffs, $unitsInYear, )
      = @_;

    # Quick hack (independent of anything else) which puts a
    # user-editable (unlocked) table of formulas in a calculation sheet.
    if ( $model->{addVolumes} && $model->{addVolumes} =~ /summar/i ) {

        $volumeData = {
            map {
                $_ => Arithmetic(
                    name          => $volumeData->{$_}{name},
                    arithmetic    => '=A1',
                    defaultFormat => '0hard',
                    arguments     => { A1 => $volumeData->{$_} },
                  )
            } @$nonExcludedComponents
        };

        Columnset(
            name    => 'User-editable volume forecast for summary',
            columns => [ @{$volumeData}{@$nonExcludedComponents} ]
        );

        $unitsInYear = Arithmetic(
            name       => 'All units (MWh)',
            arithmetic => '='
              . join( '+', map { "A$_" } 1 .. $model->{maxUnitRates} ),
            arguments => {
                map { ( "A$_" => $volumeData->{"Unit rate $_ p/kWh"} ) }
                  1 .. $model->{maxUnitRates}
            },
            defaultFormat => '0softnz',
        );

    }

    my ( $revenuesFromTariffs, $revenuesFromUnitRates, $averageUnitRate,
        $myUnits, $myMpans );

    $myMpans = Stack(
        name    => $volumeData->{'Fixed charge p/MPAN/day'}->objectShortName,
        rows    => $allTariffs,
        sources => [ $volumeData->{'Fixed charge p/MPAN/day'} ],
    );

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
        $revenuesFromTariffs = Arithmetic(
            name       => 'Net revenues (£)',
            rows       => $allTariffs,
            arithmetic => '='
              . join( '+',
                @termsWithDays
                ? ( '0.01*A400*(' . join( '+', @termsWithDays ) . ')' )
                : ('0'),
                @termsNoDays ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
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
        $myUnits =
            $unitsInYear->{noCopy} && !$unitsInYear->{location}
          ? $unitsInYear
          : Stack(
            rows    => $allTariffs,
            sources => [$unitsInYear]
          );
        $averageUnitRate =
          $revenuesFromUnitRates
          ? Arithmetic(
            name       => Label('Average unit rate p/kWh'),
            rows       => $allTariffs,
            arithmetic => '=IF(A403<>0,0.1*A1/A402,0)',
            arguments  => {
                A1   => $revenuesFromUnitRates,
                A402 => $myUnits,
                A403 => $myUnits,
            },
          )
          : Arithmetic(
            name       => Label('Average revenue from unit rates (p/kWh)'),
            rows       => $allTariffs,
            arithmetic => '=IF(A403<>0,('
              . join( '+', @termsUnits )
              . ')/A402,0)',
            arguments => {
                %args,
                A402 => $myUnits,
                A403 => $myUnits,
            },
          );
    }

    my @revenuesFromUnits = map {
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

    my $revenuesFromFixedCharges = Arithmetic(
        name          => 'Revenues from fixed charges (£)',
        defaultFormat => '0softnz',
        arithmetic    => '=A1*A4*A2/100',
        arguments     => {
            A2 => $volumeData->{'Fixed charge p/MPAN/day'},
            A4 => $daysInYear,
            A1 => $tariffTable->{'Fixed charge p/MPAN/day'},
        }
    );

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

    my $revenuesFromCapacityCharges = Arithmetic(
        name          => 'Revenues from capacity charges (£)',
        defaultFormat => '0softnz',
        arithmetic    => '=A1*A4*A2/100',
        arguments     => {
            A2 => $volumeData->{'Capacity charge p/kVA/day'},
            A4 => $daysInYear,
            A1 => $tariffTable->{'Capacity charge p/kVA/day'},
        }
    ) if $volumeData->{'Capacity charge p/kVA/day'};

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

    my ( $revenuesFromUnauthDemandCharges, $unauthProportion );

    if ( $model->{unauth} ) {

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
            name          => 'Revenues from unauthorised demand charges (£)',
            defaultFormat => '0softnz',
            arithmetic    => '=A1*A2*10',
            arguments     => {
                A2 => $volumeData->{'Unauthorised demand charge p/kVAh'},
                A1 => $tariffTable->{'Unauthorised demand charge p/kVAh'},
            }
          );

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
        );

    }

    my ( $revenuesFromReactiveCharges, $reactiveProportion );

    unless ( $model->{reactiveExcluded} ) {

        $revenuesFromReactiveCharges = Arithmetic(
            name          => 'Revenues from reactive power charges (£)',
            defaultFormat => '0softnz',
            arithmetic    => '=A1*A2*10',
            arguments     => {
                A2 => $volumeData->{'Reactive power charge p/kVArh'},
                A1 => $tariffTable->{'Reactive power charge p/kVArh'},
            }
        ) if $volumeData->{'Reactive power charge p/kVArh'};

        $reactiveProportion = Arithmetic(
            name          => 'Reactive power charge proportion',
            defaultFormat => '%softnz',
            arithmetic    => '=IF(A3<>0,A5/A1,"")',
            arguments     => {
                A1 => $revenuesFromTariffs,
                A3 => $revenuesFromTariffs,
                A5 => $revenuesFromReactiveCharges,
            }
        ) if $revenuesFromReactiveCharges;

    }

    my $averageByUnit = Arithmetic(
        name       => 'Average p/kWh',
        arithmetic => '=IF(A3<>0,0.1*A1/A2,"")',
        arguments  => {
            A1 => $revenuesFromTariffs,
            A2 => $myUnits,
            A3 => $myUnits,
        }
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

    push @{ $model->{informationTables} }, Stack( sources => [$averageByMpan] )
      if $model->{summary} =~ /info/i;

    push @{ $model->{overallSummary} }, Columnset(
        name => 'Revenue summary'
          . (
            $model->{inYear}
            ? (
                $model->{inYear} =~ /partyear/
                ? ' (period to which the new tariffs apply)'
                : ' (annualised basis)'
              )
            : ''
          ),
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

    my $totalCapacityProportionSilly = Arithmetic(
        name          => 'Capacity charge proportion',
        defaultFormat => '%softnz',
        arithmetic    => '=IF(A3<>0,A5/A1,"")',
        arguments     => {
            A1 => $totalRevenuesFromTariffs,
            A3 => $totalRevenuesFromTariffs,
            A5 => $totalRevenuesFromCapacity,
        }
    ) if $totalRevenuesFromCapacity;

    my @revenueColumns = (
        $totalUnits,
        $totalMpans,
        $totalRevenuesFromTariffs,
        $totalRevenuesFromUnitRates,
        $totalRevenuesFromFixed,
        $totalRevenuesFromCapacity   ? $totalRevenuesFromCapacity        : (),
        $totalRevenuesFromUnauth     ? $totalRevenuesFromUnauth          : (),
        $totalRevenuesFromReactive   ? $totalRevenuesFromReactive        : (),
        $model->{otherTotalRevenues} ? @{ $model->{otherTotalRevenues} } : (),
    );

    if ( $model->{sharedData} ) {
        $model->{sharedData}->addStats( 'DNO-wide aggregates',
            $model, $totalUnits, $totalMpans, $totalRevenuesFromTariffs );
        $model->{sharedData}
          ->addStats( 'Average pence per unit', $model, $averageByUnit );
        $model->{sharedData}
          ->addStats( 'Average charge per MPAN', $model, $averageByMpan );
    }

    push @{ $model->{overallSummary} },
      Columnset(
        name => 'Revenue summary by tariff component'
          . (
            $model->{inYear}
            ? (
                $model->{inYear} =~ /partyear/
                ? ' (period to which the new tariffs apply)'
                : ' (annualised basis)'
              )
            : ''
          ),
        columns => \@revenueColumns
      );

    $revenuesFromTariffs;

}

sub comparisonSummary {

    my (
        $model,          $revenuesFromTariffs, $tariffTable,
        $volumeData,     $daysInYear,          $nonExcludedComponents,
        $componentMap,   $allTariffs,          $allTariffsByEndUser,
        $unitsInYear,    $unitsLossAdjustment, $currentTariffs,
        $revenuesBefore, $unitsWholeYear,
    ) = @_;

    my $atwTariffs = Labelset(
        name => 'All ATW tariffs',
        list => [
            map { Labelset( name => $_->{list}[0], list => $_->{list} ) } @{
                $allTariffsByEndUser->{groups} || $allTariffsByEndUser->{list}
            }
        ],
    );

    my $selectedTariffsForComparison = $model->{summary}
      && $model->{summary} =~ /gen/i ? $atwTariffs : Labelset(
        name => 'Demand ATW tariffs',
        list => [ grep { !/gener/ } @{ $atwTariffs->{list} } ]
      );

    my $currentRevenues = Dataset(
        name          => 'Comparison revenue if known (£)',
        defaultFormat => '0hardnz',
        rows          => $selectedTariffsForComparison,
        data => [ map { '' } @{ $selectedTariffsForComparison->{list} } ]
    );

    my %currentTariffs = map {
        my $component = $_;
        $_ => Dataset(
            name => "Current $_",
            m%p/k(W|VAr)h%
            ? ()
            : ( defaultFormat => '0.00hard' ),
            rows => $selectedTariffsForComparison,
            data => [
                map { $componentMap->{$_}{$component} ? 0 : undef }
                  @{ $selectedTariffsForComparison->{list} }
            ]
          )
    } @$nonExcludedComponents;

    {
        my @termsNoDays;
        my @termsWithDays;
        my %args = (
            A1   => $currentRevenues,
            A9   => $currentRevenues,
            A400 => $daysInYear
        );
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
            }
            $args{"A2$pad"} = $currentTariffs{$_};
            $args{"A3$pad"} = $volumeData->{$_};
        }
        Columnset(
            name     => 'Current tariff information',
            number   => 1201,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            columns =>
              [ $currentRevenues, @currentTariffs{@$nonExcludedComponents}, ]
        );
        push @{ $model->{overallSummary} },
          $currentRevenues = Arithmetic(
            name          => 'Revenues under current tariffs (£)',
            defaultFormat => '0softnz',
            arithmetic    => '=IF(A1,A9,'
              . join( '+',
                @termsWithDays
                ? ( '0.01*A400*(' . join( '+', @termsWithDays ) . ')' )
                : ('0'),
                @termsNoDays ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
                : ('0'),
              )
              . ')',
            arguments     => \%args,
            defaultFormat => '0soft'
          );
    }

    push @{ $model->{comparisonTables} },
      Columnset(
        name    => 'Comparison with selected current tariffs',
        columns => [
            Arithmetic(
                name          => 'Change',
                defaultFormat => '%softpm',
                arithmetic    => '=IF(A1,A2/A3-1,"")',
                arguments     => {
                    A1 => $currentRevenues,
                    A2 => $revenuesFromTariffs,
                    A3 => $currentRevenues
                }
            ),
            Arithmetic(
                name          => 'Absolute change (average p/kWh)',
                rows          => $selectedTariffsForComparison,
                defaultFormat => '0.000softpm',
                arithmetic    => '=(A1-A2)/IF(A7,A8,1)/10',
                arguments     => {
                    A1 => $revenuesFromTariffs,
                    A2 => $currentRevenues,
                    A7 => $unitsInYear,
                    A8 => $unitsInYear,
                }
            ),
            Stack(
                rows    => $selectedTariffsForComparison,
                sources => [ $volumeData->{'Fixed charge p/MPAN/day'}, ]
            ),
        ],
      );

}

1;
