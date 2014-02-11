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
use SpreadsheetModel::SegmentRoot;

sub matchingdcp123 {

    my (
        $model,                  $adderAmount,
        $componentMap,           $allTariffsByEndUser,
        $demandTariffsByEndUser, $allEndUsers,
        $chargingLevels,         $nonExcludedComponents,
        $allComponents,          $daysAfter,
        $volumeAfter,            $loadCoefficients,
        $tariffsExMatching,      $daysFullYear,
        $volumeFullYear
    ) = @_;

    my $fixedAdderPot = Labelset( list => ['Adder'] );

    my $totalRevenueFromUnits = GroupBy(
        name => 'Total revenues from demand '
          . ( $model->{dcp123doublecounting} ? '' : 'active power ' )
          . 'unit rates (£/year)',
        defaultFormat => '0soft',
        source        => Arithmetic(
            name => Label(
                'From unit rates',
                'Revenues from demand '
                  . ( $model->{dcp123doublecounting} ? '' : 'active power ' )
                  . 'unit rates before matching (£/year)'
            ),
            rows       => $allTariffsByEndUser,
            arithmetic => '=IF(IV5<0,0,10*('
              . join( '+',
                $model->{dcp123doublecounting} ? 'IV701*IV702' : (),
                map { "IV$_*IV90$_" } 1 .. $model->{maxUnitRates} )
              . '))',
            arguments => {
                IV5 => $loadCoefficients,
                $model->{dcp123doublecounting}
                ? (
                    IV701 =>
                      $tariffsExMatching->{'Reactive power charge p/kVArh'},
                    IV702 => $volumeFullYear->{'Reactive power charge p/kVArh'}
                  )
                : (),
                map {
                    my $name = "Unit rate $_ p/kWh";
                    "IV$_"     => $tariffsExMatching->{$name},
                      "IV90$_" => $volumeFullYear->{$name};
                } 1 .. $model->{maxUnitRates}
            },
            defaultFormat => '0soft'
        ),
    );

    # replace with a version that includes generation
    $totalRevenueFromUnits = GroupBy(
        name => 'Total net revenues from '
          . ( $model->{dcp123doublecounting} ? '' : 'active power ' )
          . 'unit rates (£/year)',
        defaultFormat => '0soft',
        source        => Arithmetic(
            name => Label(
                'From unit rates',
                'Net revenues from '
                  . ( $model->{dcp123doublecounting} ? '' : 'active power ' )
                  . 'unit rates before matching (£/year)'
            ),
            rows       => $allTariffsByEndUser,
            arithmetic => '=10*('
              . join( '+',
                $model->{dcp123doublecounting} ? 'IV701*IV702' : (),
                map { "IV$_*IV90$_" } 1 .. $model->{maxUnitRates} )
              . ')',
            arguments => {
                IV5 => $loadCoefficients,
                $model->{dcp123doublecounting}
                ? (
                    IV701 =>
                      $tariffsExMatching->{'Reactive power charge p/kVArh'},
                    IV702 => $volumeFullYear->{'Reactive power charge p/kVArh'}
                  )
                : (),
                map {
                    my $name = "Unit rate $_ p/kWh";
                    "IV$_"     => $tariffsExMatching->{$name},
                      "IV90$_" => $volumeFullYear->{$name};
                } 1 .. $model->{maxUnitRates}
            },
            defaultFormat => '0soft'
        ),
    );

    my $totalRevenueFromFixed = GroupBy(
        name          => 'Total revenues from fixed charges (£/year)',
        defaultFormat => '0soft',
        source        => Arithmetic(
            name => Label(
                'From fixed charges',
                'Revenues from demand fixed charges before matching (£/year)'
            ),
            rows       => $allTariffsByEndUser,
            arithmetic => '=0.01*IV6*IV1*IV2', # '=IF(IV5<0,0,0.01*IV6*IV1*IV2)'
            arguments  => {
                IV5 => $loadCoefficients,
                IV6 => $daysFullYear,
                IV1 => $tariffsExMatching->{'Fixed charge p/MPAN/day'},
                IV2 => $volumeFullYear->{'Fixed charge p/MPAN/day'},
            },
            defaultFormat => '0soft'
        ),
    );

    my $totalRevenueFromCapacity = GroupBy(
        name          => 'Total revenues from demand capacity charges (£/year)',
        defaultFormat => '0soft',
        source        => Arithmetic(
            name => Label(
                'From capacity charges',
                'Revenues from demand capacity charges before matching (£/year)'
            ),
            rows       => $allTariffsByEndUser,
            arithmetic => '=IF(IV5<0,0,0.01*IV6*IV1*IV2)',
            arguments  => {
                IV5 => $loadCoefficients,
                IV6 => $daysFullYear,
                IV1 => $tariffsExMatching->{'Capacity charge p/kVA/day'},
                IV2 => $volumeFullYear->{'Capacity charge p/kVA/day'},
            },
            defaultFormat => '0soft'
        ),
    );

    my $totalRevenueFromReactive = GroupBy(
        name          => 'Total revenues from reactive power charges (£/year)',
        defaultFormat => '0soft',
        source        => Arithmetic(
            name => Label(
                'From reactive power charges',
                'Revenues from reactive power charges before matching (£/year)'
            ),
            rows       => $allTariffsByEndUser,
            arithmetic => '=10*IV1*IV2',
            arguments  => {
                IV5 => $loadCoefficients,
                IV1 => $tariffsExMatching->{'Reactive power charge p/kVArh'},
                IV2 => $volumeFullYear->{'Reactive power charge p/kVArh'},
            },
            defaultFormat => '0soft'
        ),
    );

    Columnset(
        name => 'Analysis of annual revenue by tariff before matching (£/year)',
        columns => [
            map { $_->{source} } $totalRevenueFromUnits,
            $totalRevenueFromFixed,
            $totalRevenueFromCapacity,
            $totalRevenueFromReactive,
        ]
    );

    Columnset(
        name    => 'Analysis of total annual revenue before matching (£/year)',
        columns => [
            $totalRevenueFromUnits,    $totalRevenueFromFixed,
            $totalRevenueFromCapacity, $totalRevenueFromReactive,
        ]
    );

    my @hybridTargets;

    if ( $model->{scaler} =~ /opt3|hybrid/i ) {

        @hybridTargets = (
            Arithmetic(
                name => 'Revenue matching target from unit rates (£/year)',
                defaultFormat => '0soft',
                arithmetic    => '=IV1*IV2/(IV3+IV4+IV5+IV6)',
                arguments     => {
                    IV1 => $adderAmount,
                    IV2 => $totalRevenueFromUnits,
                    IV3 => $totalRevenueFromUnits,
                    IV4 => $totalRevenueFromFixed,
                    IV5 => $totalRevenueFromCapacity,
                    IV6 => $totalRevenueFromReactive,
                }
            ),
            Arithmetic(
                name => 'Revenue matching target from fixed charges (£/year)',
                defaultFormat => '0soft',
                arithmetic    => '=IV1*IV2/(IV3+IV4+IV5+IV6)',
                arguments     => {
                    IV1 => $adderAmount,
                    IV2 => $totalRevenueFromFixed,
                    IV3 => $totalRevenueFromUnits,
                    IV4 => $totalRevenueFromFixed,
                    IV5 => $totalRevenueFromCapacity,
                    IV6 => $totalRevenueFromReactive,
                }
            ),
            Arithmetic(
                name =>
                  'Revenue matching target from capacity charges (£/year)',
                defaultFormat => '0soft',
                arithmetic    => '=IV1*IV2/(IV3+IV4+IV5+IV6)',
                arguments     => {
                    IV1 => $adderAmount,
                    IV2 => $totalRevenueFromCapacity,
                    IV3 => $totalRevenueFromUnits,
                    IV4 => $totalRevenueFromFixed,
                    IV5 => $totalRevenueFromCapacity,
                    IV6 => $totalRevenueFromReactive,
                }
            ),
            Arithmetic(
                name =>
'Revenue matching target from reactive power charges (£/year)',
                defaultFormat => '0soft',
                arithmetic    => '=IV1*IV2/(IV3+IV4+IV5+IV6)',
                arguments     => {
                    IV1 => $adderAmount,
                    IV2 => $totalRevenueFromReactive,
                    IV3 => $totalRevenueFromUnits,
                    IV4 => $totalRevenueFromFixed,
                    IV5 => $totalRevenueFromCapacity,
                    IV6 => $totalRevenueFromReactive,
                }
            ),
        );

        push @{ $model->{adderResults} },
          Columnset(
            name    => 'Allocation of matching revenue target (£/year)',
            columns => \@hybridTargets
          );

    }

    my $adderTable = {};

    {    # units

        my @columns = grep { /kWh/ } @$nonExcludedComponents;

        my @slope = map {
            Arithmetic(
                name       => "Effect through $_",
                arithmetic => '=IF(IV3<0,0,IV1*10)',
                arguments  => {
                    IV3 => $loadCoefficients,
                    IV2 => $daysAfter,
                    IV1 => $volumeAfter->{$_},
                }
            );
        } @columns;

        my $slopeSet = Columnset(
            name    => 'Marginal revenue effect of adder',
            columns => \@slope,
        );

        my $minValue;
        $minValue = Dataset(
            name     => 'Mininum unit rate p/kWh',
            data     => [ [ $1 || '' ] ],
            number   => 1077,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
        ) if $model->{scaler} =~ /min([0-9.]*)/;

        my %minAdder = map {
            my $tariffComponent = $_;
            $_ => Arithmetic(
                name       => "Adder threshold for $_",
                arithmetic => $minValue ? '=IF(IV4<0,0,IV9-IV1)'
                : '=IF(IV4<0,0,0-IV1)',
                arguments => {
                    IV4 => $loadCoefficients,
                    IV1 => $tariffsExMatching->{$_},
                    $minValue ? ( IV9 => $minValue ) : (),
                },
                rowFormats => [
                    map {
                        $componentMap->{$_}{$tariffComponent} ? undef
                          : 'unavailable';
                    } @{ $allTariffsByEndUser->{list} }
                ]
            );
        } @columns;

        my $minAdderSet = Columnset(
            name    => 'Adder value at which the minimum is breached',
            columns => [ @minAdder{@columns} ]
        );

        my $adderRate = new SpreadsheetModel::SegmentRoot(
            name   => 'General adder rate (p/kWh)',
            slopes => $slopeSet,
            target => @hybridTargets ? $hybridTargets[0] : $adderAmount,
            min    => $minAdderSet,
        );

        foreach (@columns) {
            my $tariffComponent = $_;
            $adderTable->{$_} = Arithmetic(
                name       => "Adder on $_",
                rows       => $allTariffsByEndUser,
                cols       => Labelset( list => ['Adder'] ),
                arithmetic => '=IF(IV3<0,0,MAX(IV6,IV1))',
                arguments  => {
                    IV1 => $adderRate,
                    IV3 => $loadCoefficients,
                    IV6 => $minAdder{$_},
                },
                rowFormats => [
                    map {
                        $componentMap->{$_}{$tariffComponent}
                          ? undef
                          : 'unavailable';
                    } @{ $allTariffsByEndUser->{list} }
                ]
            );
        }

    }

    if (@hybridTargets) {

        {    # fixed

            my @columns = grep { /fixed/i } @$nonExcludedComponents;

            my @slope = map {
                Arithmetic(
                    name       => "Effect through $_",
                    arithmetic => '=IF(IV9>0,IV1*IV2*0.01,0)'
                    ,    # '=IF(IV3>0,IV1*IV2*0.01,0)',
                    arguments => {
                        IV3 => $loadCoefficients,
                        IV2 => $daysAfter,
                        IV1 => $volumeAfter->{$_},
                        IV9 => $tariffsExMatching->{$_},
                    }
                );
            } @columns;

            my $slopeSet = Columnset(
                name    => 'Marginal revenue effect of adder',
                columns => \@slope,
            );

            my %minAdder = map {
                my $tariffComponent = $_;
                $_ => Arithmetic(
                    name => "Adder threshold for $_",
                    arithmetic => '=0-IV1',    # '=IF(IV4<0,0,0-IV1)',
                    arguments  => {
                        IV4 => $loadCoefficients,
                        IV1 => $tariffsExMatching->{$_},
                    },
                    rowFormats => [
                        map {
                            $componentMap->{$_}{$tariffComponent}
                              ? undef
                              : 'unavailable';
                        } @{ $allTariffsByEndUser->{list} }
                    ]
                );
            } @columns;

            my $minAdderSet = Columnset(
                name    => 'Adder value at which the minimum is breached',
                columns => [ @minAdder{@columns} ]
            );

            my $adderRate = new SpreadsheetModel::SegmentRoot(
                name   => 'General adder rate (p/MPAN/day)',
                slopes => $slopeSet,
                target => $hybridTargets[1],
                min    => $minAdderSet,
            );

            foreach (@columns) {
                my $tariffComponent = $_;
                $adderTable->{$_} = Arithmetic(
                    name       => "Adder on $_",
                    rows       => $allTariffsByEndUser,
                    cols       => Labelset( list => ['Adder'] ),
                    arithmetic => '=IF(IV9>0,MAX(IV6,IV1),0)'
                    ,    # '=IF(IV3<0,0,MAX(IV6,IV1))',
                    arguments => {
                        IV1 => $adderRate,
                        IV3 => $loadCoefficients,
                        IV6 => $minAdder{$_},
                        IV9 => $tariffsExMatching->{$_},
                    },
                    rowFormats => [
                        map {
                            $componentMap->{$_}{$tariffComponent}
                              ? undef
                              : 'unavailable';
                        } @{ $allTariffsByEndUser->{list} }
                    ]
                );
            }

        }

        {    # capacity

            my @columns = grep { /kVA\/day/ } @$nonExcludedComponents;

            my @slope = map {
                Arithmetic(
                    name       => "Effect through $_",
                    arithmetic => '=IF(IV3<0,0,IV1*IV2*0.01)',
                    arguments  => {
                        IV3 => $loadCoefficients,
                        IV2 => $daysAfter,
                        IV1 => $volumeAfter->{$_},
                    }
                );
            } @columns;

            my $slopeSet = Columnset(
                name    => 'Marginal revenue effect of adder',
                columns => \@slope,
            );

            my %minAdder = map {
                my $tariffComponent = $_;
                $_ => Arithmetic(
                    name       => "Adder threshold for $_",
                    arithmetic => '=IF(IV4<0,0,0-IV1)',
                    arguments  => {
                        IV4 => $loadCoefficients,
                        IV1 => $tariffsExMatching->{$_},
                    },
                    rowFormats => [
                        map {
                            $componentMap->{$_}{$tariffComponent}
                              ? undef
                              : 'unavailable';
                        } @{ $allTariffsByEndUser->{list} }
                    ]
                );
            } @columns;

            my $minAdderSet = Columnset(
                name    => 'Adder value at which the minimum is breached',
                columns => [ @minAdder{@columns} ]
            );

            my $adderRate = new SpreadsheetModel::SegmentRoot(
                name   => 'General adder rate (p/kVA/day)',
                slopes => $slopeSet,
                target => $hybridTargets[2],
                min    => $minAdderSet,
            );

            foreach (@columns) {
                my $tariffComponent = $_;
                $adderTable->{$_} = Arithmetic(
                    name       => "Adder on $_",
                    rows       => $allTariffsByEndUser,
                    cols       => Labelset( list => ['Adder'] ),
                    arithmetic => '=IF(IV3<0,0,MAX(IV6,IV1))',
                    arguments  => {
                        IV1 => $adderRate,
                        IV3 => $loadCoefficients,
                        IV6 => $minAdder{$_},
                    },
                    rowFormats => [
                        map {
                            $componentMap->{$_}{$tariffComponent}
                              ? undef
                              : 'unavailable';
                        } @{ $allTariffsByEndUser->{list} }
                    ]
                );
            }

        }

    }

    {    # reactive

        my @columns = grep { /kVArh/ } @$nonExcludedComponents;

        my @slope = map {
            Arithmetic(
                name       => "Effect through $_",
                arithmetic => '=IV1*10',             # '=IF(IV3<0,0,IV1*10)',
                arguments  => {
                    IV3 => $loadCoefficients,
                    IV1 => $volumeAfter->{$_},
                }
            );
        } @columns;

        my $slopeSet = Columnset(
            name    => 'Marginal revenue effect of adder',
            columns => \@slope,
        );

        my %minAdder = map {
            my $tariffComponent = $_;
            $_ => Arithmetic(
                name       => "Adder threshold for $_",
                arithmetic => '=0-IV1',                  # '=IF(IV4<0,0,0-IV1)',
                arguments  => {
                    IV4 => $loadCoefficients,
                    IV1 => $tariffsExMatching->{$_},
                },
                rowFormats => [
                    map {
                        $componentMap->{$_}{$tariffComponent}
                          ? undef
                          : 'unavailable';
                    } @{ $allTariffsByEndUser->{list} }
                ]
            );
        } @columns;

        my $minAdderSet = Columnset(
            name    => 'Adder value at which the minimum is breached',
            columns => [ @minAdder{@columns} ]
        );

        my $adderRate = new SpreadsheetModel::SegmentRoot(
            name   => 'General adder rate (p/kVArh)',
            slopes => $slopeSet,
            target => $hybridTargets[3],
            min    => $minAdderSet,
        );

        foreach (@columns) {
            my $tariffComponent = $_;
            $adderTable->{$_} = Arithmetic(
                name => "Adder on $_",
                rows => $allTariffsByEndUser,
                cols => Labelset( list => ['Adder'] ),
                arithmetic => '=MAX(IV6,IV1)',    # '=IF(IV3<0,0,MAX(IV6,IV1))',
                arguments  => {
                    IV1 => $adderRate,
                    IV3 => $loadCoefficients,
                    IV6 => $minAdder{$_},
                },
                rowFormats => [
                    map {
                        $componentMap->{$_}{$tariffComponent}
                          ? undef
                          : 'unavailable';
                    } @{ $allTariffsByEndUser->{list} }
                ]
            );
        }

    }

    # end of option 3-only SegmentRoots

    my $revenuesFromAdder;

    {
        my @termsNoDays;
        my @termsWithDays;
        my %args = ( IV400 => $daysAfter );
        my $i = 1;

        foreach ( grep { $adderTable->{$_} } @$nonExcludedComponents ) {
            ++$i;
            my $pad = "$i";
            $pad = "0$pad" while length $pad < 3;
            if (m#/day#) {
                push @termsWithDays, "IV2$pad*IV3$pad";
            }
            else {
                push @termsNoDays, "IV2$pad*IV3$pad";
            }
            $args{"IV2$pad"} = $adderTable->{$_};
            $args{"IV3$pad"} = $volumeAfter->{$_};
        }

        $revenuesFromAdder = Arithmetic(
            name       => 'Net revenues by tariff from adder',
            rows       => $allTariffsByEndUser,
            arithmetic => '='
              . join( '+',
                @termsWithDays
                ? ( '0.01*IV400*(' . join( '+', @termsWithDays ) . ')' )
                : ('0'),
                @termsNoDays ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
                : ('0'),
              ),
            arguments     => \%args,
            defaultFormat => '0softnz'
        );

        push @{ $model->{adderResults} },
          Columnset(
            name    => 'Adder',
            columns => [
                ( grep { $_ } @{$adderTable}{@$nonExcludedComponents} ),
                $revenuesFromAdder
            ]
          );

    }

    my $totalRevenuesFromAdder = GroupBy(
        name          => 'Total net revenues from adder (£/year)',
        rows          => 0,
        cols          => 0,
        source        => $revenuesFromAdder,
        defaultFormat => '0softnz'
    );

    my $siteSpecificCharges;    # not implemented in this option

    $totalRevenuesFromAdder, $siteSpecificCharges, $adderTable;

}

1;
