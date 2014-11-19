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
                    arithmetic    => '=IV1',
                    defaultFormat => '0hard',
                    arguments     => { IV1 => $volumeData->{$_} },
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
              . join( '+', map { "IV$_" } 1 .. $model->{maxUnitRates} ),
            arguments => {
                map { ( "IV$_" => $volumeData->{"Unit rate $_ p/kWh"} ) }
                  1 .. $model->{maxUnitRates}
            },
            defaultFormat => '0softnz',
        );

    }

    my ( $revenuesFromTariffs, $revenuesFromUnitRates, $averageUnitRate,
        $myUnits, $myMpans );

    $myMpans = Stack(
        rows    => $allTariffs,
        sources => [ $volumeData->{'Fixed charge p/MPAN/day'} ],
    );

    {
        my @termsNoDays;
        my @termsUnits;
        my @termsWithDays;
        my %args = ( IV400 => $daysInYear );
        my $i = 1;
        foreach (@$nonExcludedComponents) {
            ++$i;
            my $pad = "$i";
            $pad = "0$pad" while length $pad < 3;
            if (m#/day#) {
                push @termsWithDays, "IV2$pad*IV3$pad";
            }
            else {
                push @termsNoDays, "IV2$pad*IV3$pad";
                push @termsUnits, "IV2$pad*IV3$pad" if /kWh/;
            }
            $args{"IV2$pad"} = $tariffTable->{$_};
            $args{"IV3$pad"} = $volumeData->{$_};
        }
        $revenuesFromTariffs = Arithmetic(
            name       => 'Net revenues (£)',
            rows       => $allTariffs,
            arithmetic => '='
              . join( '+',
                @termsWithDays
                ? ( '0.01*IV400*(' . join( '+', @termsWithDays ) . ')' )
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
            arithmetic => '=IF(IV403<>0,0.1*IV1/IV402,0)',
            arguments  => {
                IV1   => $revenuesFromUnitRates,
                IV402 => $myUnits,
                IV403 => $myUnits,
            },
          )
          : Arithmetic(
            name       => Label('Average revenue from unit rates (p/kWh)'),
            rows       => $allTariffs,
            arithmetic => '=IF(IV403<>0,('
              . join( '+', @termsUnits )
              . ')/IV402,0)',
            arguments => {
                %args,
                IV402 => $myUnits,
                IV403 => $myUnits,
            },
          );
    }

    my @revenuesFromUnits = map {
        Arithmetic(
            name          => "Net revenues from unit rate $_ (£)",
            defaultFormat => '0softnz',
            arithmetic    => '=IV1*IV2*10',
            arguments     => {
                IV2 => $volumeData->{"Unit rate $_ p/kWh"},
                IV1 => $tariffTable->{"Unit rate $_ p/kWh"},
            }
        );
    } 1 .. $model->{maxUnitRates};

    my @unitProportion = $revenuesFromUnitRates
      ? map {
        Arithmetic(
            name          => "Rate $_ revenue proportion",
            defaultFormat => '%softnz',
            arithmetic    => '=IF(IV3<>0,IV5/IV1,"")',
            arguments     => {
                IV1 => $revenuesFromUnitRates,
                IV3 => $revenuesFromUnitRates,
                IV5 => $revenuesFromUnits[ $_ - 1 ],
            }
        );
      } 1 .. $model->{maxUnitRates}
      : ();

    my $revenuesFromFixedCharges = Arithmetic(
        name          => 'Revenues from fixed charges (£)',
        defaultFormat => '0softnz',
        arithmetic    => '=IV1*IV4*IV2/100',
        arguments     => {
            IV2 => $volumeData->{'Fixed charge p/MPAN/day'},
            IV4 => $daysInYear,
            IV1 => $tariffTable->{'Fixed charge p/MPAN/day'},
        }
    );

    my $fixedProportion = Arithmetic(
        name          => 'Fixed charge proportion',
        defaultFormat => '%softnz',
        arithmetic    => '=IF(IV3<>0,IV5/IV1,"")',
        arguments     => {
            IV1 => $revenuesFromTariffs,
            IV3 => $revenuesFromTariffs,
            IV5 => $revenuesFromFixedCharges,
        }
    );

    my $revenuesFromCapacityCharges = Arithmetic(
        name          => 'Revenues from capacity charges (£)',
        defaultFormat => '0softnz',
        arithmetic    => '=IV1*IV4*IV2/100',
        arguments     => {
            IV2 => $volumeData->{'Capacity charge p/kVA/day'},
            IV4 => $daysInYear,
            IV1 => $tariffTable->{'Capacity charge p/kVA/day'},
        }
    ) if $volumeData->{'Capacity charge p/kVA/day'};

    my $capacityProportion = Arithmetic(
        name          => 'Capacity charge proportion',
        defaultFormat => '%softnz',
        arithmetic    => '=IF(IV3<>0,IV5/IV1,"")',
        arguments     => {
            IV1 => $revenuesFromTariffs,
            IV3 => $revenuesFromTariffs,
            IV5 => $revenuesFromCapacityCharges,
        }
    ) if $revenuesFromCapacityCharges;

    my ( $revenuesFromUnauthDemandCharges, $unauthProportion );

    if ( $model->{unauth} ) {

        $revenuesFromUnauthDemandCharges =
          $model->{unauth} && $model->{unauth} =~ /day/i
          ? Arithmetic(
            name          => 'Revenues from exceeded capacity charges (£)',
            defaultFormat => '0softnz',
            arithmetic    => '=IV1*IV4*IV2/100',
            arguments     => {
                IV2 => $volumeData->{'Exceeded capacity charge p/kVA/day'},
                IV4 => $daysInYear,
                IV1 => $tariffTable->{'Exceeded capacity charge p/kVA/day'},
            }
          )
          : Arithmetic(
            name          => 'Revenues from unauthorised demand charges (£)',
            defaultFormat => '0softnz',
            arithmetic    => '=IV1*IV2*10',
            arguments     => {
                IV2 => $volumeData->{'Unauthorised demand charge p/kVAh'},
                IV1 => $tariffTable->{'Unauthorised demand charge p/kVAh'},
            }
          );

        $unauthProportion = Arithmetic(
            name => $model->{unauth} && $model->{unauth} =~ /day/i
            ? 'Exceeded capacity charge proportion'
            : 'Unauthorised demand charge proportion',
            defaultFormat => '%softnz',
            arithmetic    => '=IF(IV3<>0,IV5/IV1,"")',
            arguments     => {
                IV1 => $revenuesFromTariffs,
                IV3 => $revenuesFromTariffs,
                IV5 => $revenuesFromUnauthDemandCharges,
            }
        );

    }

    my ( $revenuesFromReactiveCharges, $reactiveProportion );

    unless ( $model->{reactiveExcluded} ) {

        $revenuesFromReactiveCharges = Arithmetic(
            name          => 'Revenues from reactive power charges (£)',
            defaultFormat => '0softnz',
            arithmetic    => '=IV1*IV2*10',
            arguments     => {
                IV2 => $volumeData->{'Reactive power charge p/kVArh'},
                IV1 => $tariffTable->{'Reactive power charge p/kVArh'},
            }
        ) if $volumeData->{'Reactive power charge p/kVArh'};

        $reactiveProportion = Arithmetic(
            name          => 'Reactive power charge proportion',
            defaultFormat => '%softnz',
            arithmetic    => '=IF(IV3<>0,IV5/IV1,"")',
            arguments     => {
                IV1 => $revenuesFromTariffs,
                IV3 => $revenuesFromTariffs,
                IV5 => $revenuesFromReactiveCharges,
            }
        ) if $revenuesFromReactiveCharges;

    }

    my $averageByUnit = Arithmetic(
        name       => 'Average p/kWh',
        arithmetic => '=IF(IV3<>0,0.1*IV1/IV2,"")',
        arguments  => {
            IV1 => $revenuesFromTariffs,
            IV2 => $myUnits,
            IV3 => $myUnits,
        }
    );

    my $averageByMpan = Arithmetic(
        name          => 'Average £/MPAN',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(IV3<>0,IV1/IV2,"")',
        arguments     => {
            IV1 => $revenuesFromTariffs,
            IV2 => $myMpans,
            IV3 => $myMpans,
        }
    );

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
                arithmetic    => '=IF(IV3<>0,IV1/IV2*100/IV4,"")',
                arguments     => {
                    IV1 => $revenuesFromTariffs,
                    IV2 => $volumeData->{'Fixed charge p/MPAN/day'},
                    IV3 => $volumeData->{'Fixed charge p/MPAN/day'},
                    IV4 => $daysInYear,
                }
            ),

            1 ? () : Arithmetic(
                name          => 'Average p/kVA/day',
                defaultFormat => '0.00softnz',
                arithmetic    => '=IF(IV3<>0,IV1/IV2*100/IV4,"")',
                arguments     => {
                    IV1 => $revenuesFromTariffs,
                    IV2 => $volumeData->{'Capacity charge p/kVA/day'},
                    IV3 => $volumeData->{'Capacity charge p/kVA/day'},
                    IV4 => $daysInYear,
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
                arithmetic    => '=IF(IV3<>0,IV5/(IV1/IV2*100/IV4),"")',
                arguments     => {
                    IV1 => $revenuesFromTariffs,
                    IV2 => $volumeData->{'Fixed charge p/MPAN/day'},
                    IV3 => $volumeData->{'Fixed charge p/MPAN/day'},
                    IV4 => $daysInYear,
                    IV5 => $tariffTable->{'Fixed charge p/MPAN/day'},
                }
            ),

            1 ? () : Arithmetic(
                name          => 'Capacity charge proportion',
                defaultFormat => '%softnz',
                arithmetic    => '=IF(IV3<>0,IV5/(IV1/IV2*100/IV4),"")',
                arguments     => {
                    IV1 => $revenuesFromTariffs,
                    IV2 => $volumeData->{'Capacity charge p/kVA/day'},
                    IV3 => $volumeData->{'Capacity charge p/kVA/day'},
                    IV4 => $daysInYear,
                    IV5 => $tariffTable->{'Capacity charge p/kVA/day'},
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
        arithmetic    => '=IF(IV3<>0,IV5/IV1,"")',
        arguments     => {
            IV1 => $totalRevenuesFromTariffs,
            IV3 => $totalRevenuesFromTariffs,
            IV5 => $totalRevenuesFromFixed,
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
        arithmetic    => '=IF(IV3<>0,IV5/IV1,"")',
        arguments     => {
            IV1 => $totalRevenuesFromTariffs,
            IV3 => $totalRevenuesFromTariffs,
            IV5 => $totalRevenuesFromCapacity,
        }
    ) if $totalRevenuesFromCapacity;

    my @revenueColumns = (
        $totalUnits,
        $totalMpans,
        $totalRevenuesFromTariffs,
        $totalRevenuesFromUnitRates,
        $totalRevenuesFromFixed,
        $totalRevenuesFromCapacity ? $totalRevenuesFromCapacity : (),
        $totalRevenuesFromUnauth   ? $totalRevenuesFromUnauth   : (),
        $totalRevenuesFromReactive ? $totalRevenuesFromReactive : (),
    );

    if ( $model->{sharedData} ) {
        $model->{sharedData}->addStats( 'DNO-wide aggregates',
            $model, $totalUnits, $totalMpans, $totalRevenuesFromTariffs );
        $model->{sharedData}
          ->addStats( 'Average pence per unit', $model, $averageByUnit );
        if ( $model->{arp} && $model->{arp} =~ /permpan/i ) {
            $model->{sharedData}
              ->addStats( 'Average charge per MPAN', $model, $averageByMpan );
        }
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

sub consultationSummary {

    my (
        $model,          $revenuesFromTariffs, $tariffTable,
        $volumeData,     $daysInYear,          $nonExcludedComponents,
        $componentMap,   $allTariffs,          $allTariffsByEndUser,
        $unitsInYear,    $unitsLossAdjustment, $currentTariffs,
        $revenuesBefore, $unitsWholeYear,
    ) = @_;

    my $tariffsByAtw = Labelset(
        name   => 'All tariffs by ATW tariff',
        groups => [
            map { Labelset( name => $_->{list}[0], list => $_->{list} ) } @{
                $allTariffsByEndUser->{groups} || $allTariffsByEndUser->{list}
            }
        ]
    );

    my $atwTariffs = Labelset(
        name => 'All ATW tariffs',
        list => $tariffsByAtw->{groups}
    );

    my $selectedTariffsForComparison = $model->{summary}
      && $model->{summary} =~ /gen/i ? $atwTariffs : Labelset(
        name => 'Demand ATW tariffs',
        list => [ grep { !/gener/ } @{ $tariffsByAtw->{groups} } ]
      );

    my $hardRevenues = Dataset(
        name          => 'Current revenues if known (£)',
        defaultFormat => '0hardnz',
        rows          => $selectedTariffsForComparison,
        data => [ map { '' } @{ $selectedTariffsForComparison->{list} } ]
    );

    my $currentRevenues;

=head To do

Change something below to put something like table 1095 in full-year models instead of table 1201.

=cut

    if ( $currentTariffs && !$model->{force1201} ) {
        my @termsNoDays;
        my @termsWithDays;
        my %args = ( IV400 => $daysInYear );
        my $i = 1;
        foreach (@$nonExcludedComponents) {
            ++$i;
            my $pad = "$i";
            $pad = "0$pad" while length $pad < 3;
            if (m#/day#) {
                push @termsWithDays, "IV2$pad*IV3$pad";
            }
            else {
                push @termsNoDays, "IV2$pad*IV3$pad";
            }
            $args{"IV2$pad"} = $currentTariffs->{$_};
            $args{"IV3$pad"} = $volumeData->{$_};
        }
        $currentRevenues = Arithmetic(
            name          => 'Revenues under current tariffs (£)',
            defaultFormat => '0softnz',
            rows          => $selectedTariffsForComparison,
            arithmetic    => '='
              . join( '+',
                @termsWithDays
                ? ( '0.01*IV400*(' . join( '+', @termsWithDays ) . ')' )
                : ('0'),
                @termsNoDays ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
                : ('0'),
              ),
            arguments     => \%args,
            defaultFormat => '0soft'
        );
    }
    else {
        my %currentTariffs = map {
            my $component = $_;
            $_ => Dataset(
                name          => "Current $_",
                defaultFormat => '0.000hardnz',
                rows          => $selectedTariffsForComparison,
                data          => [
                    map { $componentMap->{$_}{$component} ? 0 : undef }
                      @{ $selectedTariffsForComparison->{list} }
                ]
              )
        } @$nonExcludedComponents;

        {
            my @termsNoDays;
            my @termsWithDays;
            my %args = (
                IV1   => $hardRevenues,
                IV9   => $hardRevenues,
                IV400 => $daysInYear
            );
            my $i = 1;
            foreach (@$nonExcludedComponents) {
                ++$i;
                my $pad = "$i";
                $pad = "0$pad" while length $pad < 3;
                if (m#/day#) {
                    push @termsWithDays, "IV2$pad*IV3$pad";
                }
                else {
                    push @termsNoDays, "IV2$pad*IV3$pad";
                }
                $args{"IV2$pad"} = $currentTariffs{$_};
                $args{"IV3$pad"} = $volumeData->{$_};
            }
            $currentRevenues = Arithmetic(
                name          => 'Revenues under current tariffs (£)',
                defaultFormat => '0softnz',
                arithmetic    => '=IF(IV1,IV9,'
                  . join( '+',
                    @termsWithDays
                    ? ( '0.01*IV400*(' . join( '+', @termsWithDays ) . ')' )
                    : ('0'),
                    @termsNoDays ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
                    : ('0'),
                  )
                  . ')',
                arguments     => \%args,
                defaultFormat => '0soft'
            );
        }

        push @{ $model->{consultationInput} },
          Columnset(
            name     => 'Current tariff information',
            number   => 1201,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            columns =>
              [ $hardRevenues, @currentTariffs{@$nonExcludedComponents}, ]
          );

    }

    push @{ $model->{consultationInput} }, $currentRevenues;

    my %atwVolumes = map {
        $_ => Stack( rows => $atwTariffs, sources => [ $volumeData->{$_} ] )
    } @$nonExcludedComponents;

    my $atwUnits = Stack( rows => $atwTariffs, sources => [$unitsInYear] );

    Columnset(
        name    => 'All-the-way volumes',
        columns => [ @atwVolumes{@$nonExcludedComponents}, $atwUnits ]
    );

    my $normalisedTo = Constant(
        name => 'Normalised to',
        rows => $atwTariffs,
        data => [
            map {
                    /gener|un-?met/                                  ? 'MWh'
                  : $componentMap->{$_}{'Capacity charge p/kVA/day'} ? 'kVA'
                  : 'MPAN'
            } @{ $atwTariffs->{list} }
        ]
    );

    push @{ $model->{consultationTables} },

      Columnset(
        name    => 'New tariffs',
        columns => [
            map {
                Stack(
                    sources       => [ $tariffTable->{$_} ],
                    defaultFormat => !$tariffTable->{$_}{defaultFormat}
                      || $tariffTable->{$_}{defaultFormat} =~ /000/
                    ? '0.000copynz'
                    : '0.00copynz'
                );
            } @$nonExcludedComponents
        ]
      ),

      Columnset(
        name    => 'Illustrative proposed reactive power charges',
        columns => [
            Stack(
                name => 'Reactive power charge p/kVArh',
                rows => Labelset(
                    list => [
                        grep {
                            !/LDNO/i
                              && $componentMap->{$_}
                              {'Reactive power charge p/kVArh'}
                          } map { $allTariffs->{list}[$_] }
                          $allTariffs->indices
                    ]
                ),
                sources => [ $tariffTable->{'Reactive power charge p/kVArh'} ]
            )
        ]
      )

      if $model->{summary} =~ /big/i;

    my %adjustedVolume = $atwVolumes{'Capacity charge p/kVA/day'}
      ? (
        map {
            $_ => Arithmetic(
                name => 'Normalised '
                  . (
                    ref $volumeData->{$_}{name}
                    ? $volumeData->{$_}{name}[0]
                    : $volumeData->{$_}{name}
                  ),
                rows       => $tariffsByAtw,
                arithmetic => '=IV1/IF(IV4="kVA",IF(IV51,IV52,1),'
                  . 'IF(IV3="MPAN",IF(IV6,IV9,1),IF(IV7,IV8,1)))'
                  . ( /Unit rate/i && $unitsLossAdjustment ? '/(1+IV2)' : '' ),
                arguments => {
                    IV1 => $atwVolumes{$_},
                    $unitsLossAdjustment ? ( IV2 => $unitsLossAdjustment ) : (),
                    IV3  => $normalisedTo,
                    IV4  => $normalisedTo,
                    IV51 => $atwVolumes{'Capacity charge p/kVA/day'},
                    IV52 => $atwVolumes{'Capacity charge p/kVA/day'},
                    IV6  => $atwVolumes{'Fixed charge p/MPAN/day'},
                    IV7  => $atwUnits,
                    IV8  => $atwUnits,
                    IV9  => $atwVolumes{'Fixed charge p/MPAN/day'}
                }
            );
        } @$nonExcludedComponents
      )
      : (
        map {
            $_ => Arithmetic(
                name => 'Normalised '
                  . (
                    ref $volumeData->{$_}{name} ? $volumeData->{$_}{name}[0]
                    : $volumeData->{$_}{name}
                  ),
                rows       => $tariffsByAtw,
                arithmetic => '=IV1/'
                  . 'IF(IV3="MPAN",IF(IV6,IV9,1),IF(IV7,IV8,1))'
                  . ( /Unit rate/i && $unitsLossAdjustment ? '/(1+IV2)' : '' ),
                arguments => {
                    IV1 => $atwVolumes{$_},
                    $unitsLossAdjustment ? ( IV2 => $unitsLossAdjustment ) : (),
                    IV3 => $normalisedTo,
                    IV4 => $normalisedTo,
                    IV6 => $atwVolumes{'Fixed charge p/MPAN/day'},
                    IV7 => $atwUnits,
                    IV8 => $atwUnits,
                    IV9 => $atwVolumes{'Fixed charge p/MPAN/day'}
                }
            );
        } @$nonExcludedComponents
      );

    my $rev;

    {
        my @termsNoDays;
        my @termsWithDays;
        my %args = ( IV400 => $daysInYear );
        my $i = 1;
        foreach (@$nonExcludedComponents) {
            ++$i;
            my $pad = "$i";
            $pad = "0$pad" while length $pad < 3;
            if (m#/day#) {
                push @termsWithDays, "IV2$pad*IV3$pad";
            }
            else {
                push @termsNoDays, "IV2$pad*IV3$pad";
            }
            $args{"IV2$pad"} = $tariffTable->{$_};
            $args{"IV3$pad"} = $adjustedVolume{$_};
        }
        $rev = Arithmetic(
            name          => 'Normalised revenues (£)',
            rows          => $tariffsByAtw,
            defaultFormat => '0softnz',
            arithmetic    => '='
              . join( '+',
                @termsWithDays
                ? ( '0.01*IV400*(' . join( '+', @termsWithDays ) . ')' )
                : ('0'),
                @termsNoDays ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
                : ('0'),
              ),
            arguments     => \%args,
            defaultFormat => '0.00soft'
        );
    }

    push @{ $model->{consultationInput} },
      Columnset(
        name    => 'Normalised volumes for comparisons',
        columns => [ @adjustedVolume{@$nonExcludedComponents}, $rev ]
      );

    my $atwMarginTariffs = Labelset(
        list => [ grep { $#{ $_->{list} } } @{ $atwTariffs->{list} } ] );

    my $atwRev = Stack(
        name          => 'All-the-way charges (normalised £)',
        defaultFormat => '0.00softnz',
        rows          => $atwMarginTariffs,
        sources       => [$rev]
    );

    my @rev = Stack(
        rows    => $atwMarginTariffs,
        sources => [$normalisedTo],
    );

    push @rev, $atwRev;

    foreach my $idnoType ( 'LV', 'HV', 'HV Sub', 'Any' ) {
        next unless grep { /^LDNO $idnoType:/ } @{ $allTariffs->{list} };
        push @{ $atwMarginTariffs->{accepts} }, my $tariffset = Labelset(
            list => [
                map {
                    ( grep { /^LDNO $idnoType:/ } @{ $_->{list} } )[0]
                      || 'N/A'
                } @{ $atwMarginTariffs->{list} }
            ],
        );
        push @{ $model->{consultationInput} },
          my $irev = Stack(
            name    => "LDNO $idnoType charges (normalised £)",
            rows    => $tariffset,
            sources => [$rev]
          );
        push @rev,
          Arithmetic(
            name          => "LDNO $idnoType margin (normalised £)",
            defaultFormat => '0.00soft',
            rowFormats    => [
                map { $_ eq 'N/A' ? 'unavailable' : undef }
                  @{ $tariffset->{list} }
            ],
            arithmetic => '=IF(IV3,IV1-IV2,"")',
            arguments  => { IV1 => $atwRev, IV2 => $irev, IV3 => $irev }
          );
    }

    push @{ $model->{consultationTables} }, Columnset(
        name    => 'Comparison with current all-the-way demand tariffs',
        columns => [

            $model->{summary} =~ /jun/i
            ? (
                Arithmetic(
                    name          => 'Revenues (£m)',
                    defaultFormat => '0.00softnz',
                    rows          => $selectedTariffsForComparison,
                    arithmetic    => '=IV1*1e-6',
                    arguments     => { IV1 => $revenuesFromTariffs }
                ),
                Stack(
                    sources => [$normalisedTo],
                    rows    => $selectedTariffsForComparison,

                    #                defaultFormat => 'right'
                ),
                Arithmetic(
                    name          => 'Normalised units (MWh)',
                    defaultFormat => '0.000softnz',
                    rows          => $selectedTariffsForComparison,
                    arithmetic    => '=IV1/IF(IV4="kVA",IF(IV51,IV52,1),'
                      . 'IF(IV3="MPAN",IF(IV6,IV9,1),IF(IV7,IV8,1)))',
                    arguments => {
                        IV1  => $unitsInYear,
                        IV3  => $normalisedTo,
                        IV4  => $normalisedTo,
                        IV51 => $volumeData->{'Capacity charge p/kVA/day'},
                        IV52 => $volumeData->{'Capacity charge p/kVA/day'},
                        IV6  => $volumeData->{'Fixed charge p/MPAN/day'},
                        IV7  => $unitsInYear,
                        IV8  => $unitsInYear,
                        IV9  => $volumeData->{'Fixed charge p/MPAN/day'}
                    }
                ),
                Arithmetic(
                    name          => 'Normalised revenue (£)',
                    rows          => $selectedTariffsForComparison,
                    defaultFormat => '0.00softnz',
                    arithmetic    => '=IV1/IF(IV4="kVA",IF(IV51,IV52,1),'
                      . 'IF(IV3="MPAN",IF(IV6,IV9,1),IF(IV7,IV8,1)))',
                    arguments => {
                        IV1  => $revenuesFromTariffs,
                        IV3  => $normalisedTo,
                        IV4  => $normalisedTo,
                        IV51 => $volumeData->{'Capacity charge p/kVA/day'},
                        IV52 => $volumeData->{'Capacity charge p/kVA/day'},
                        IV6  => $volumeData->{'Fixed charge p/MPAN/day'},
                        IV7  => $unitsInYear,
                        IV8  => $unitsInYear,
                        IV9  => $volumeData->{'Fixed charge p/MPAN/day'}
                    }
                ),
              )
            : $model->{summary} !~ /impact/i ? (    # August 2009

              )
            : (                                     # not used
                Arithmetic(
                    name          => 'Revenue impact (£m)',
                    defaultFormat => '0.00softpm',
                    arithmetic    => '=1e-6*(IV2-IV1)',
                    arguments     => {
                        IV1 => $currentRevenues,
                        IV2 => $revenuesFromTariffs,
                    }
                ),
                Arithmetic(
                    name          => 'Current p/kWh',
                    rows          => $selectedTariffsForComparison,
                    defaultFormat => '0.000softnz',
                    arithmetic    => '=IV1/IF(IV7,IV8,1)/10',
                    arguments     => {
                        IV1 => $currentRevenues,
                        IV7 => $unitsInYear,
                        IV8 => $unitsInYear,
                    }
                ),
                Arithmetic(
                    name          => 'Proposed p/kWh',
                    rows          => $selectedTariffsForComparison,
                    defaultFormat => '0.000softnz',
                    arithmetic    => '=IV1/IF(IV7,IV8,1)/10',
                    arguments     => {
                        IV1 => $revenuesFromTariffs,
                        IV7 => $unitsInYear,
                        IV8 => $unitsInYear,
                    }
                ),
            ),
            Arithmetic(
                name          => 'Change',
                defaultFormat => '%softpm',
                arithmetic    => '=IF(IV1,IV2/IV3-1,"")',
                arguments     => {
                    IV1 => $currentRevenues,
                    IV2 => $revenuesFromTariffs,
                    IV3 => $currentRevenues
                }
            ),
            $model->{summary} =~ /jun/i ? () : (    # August 2009
                Arithmetic(
                    name          => 'Absolute change (average p/kWh)',
                    rows          => $selectedTariffsForComparison,
                    defaultFormat => '0.000softpm',
                    arithmetic    => '=(IV1-IV2)/IF(IV7,IV8,1)/10',
                    arguments     => {
                        IV1 => $revenuesFromTariffs,
                        IV2 => $currentRevenues,
                        IV7 => $unitsInYear,
                        IV8 => $unitsInYear,
                    }
                ),
                Stack(
                    rows    => $selectedTariffsForComparison,
                    sources => [ $volumeData->{'Fixed charge p/MPAN/day'}, ]
                ),
            ),
          ]

    );

    push @{ $model->{consultationTables} },
      Columnset(
        name    => 'LDNO margins in use of system charges',
        columns => \@rev
      );

}

1;
