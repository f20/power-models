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
            arithmetic => '=IF(A5<0,0,10*('
              . join( '+',
                $model->{dcp123doublecounting} ? 'A701*A702' : (),
                map { "A$_*A90$_" } 1 .. $model->{maxUnitRates} )
              . '))',
            arguments => {
                A5 => $loadCoefficients,
                $model->{dcp123doublecounting}
                ? (
                    A701 =>
                      $tariffsExMatching->{'Reactive power charge p/kVArh'},
                    A702 => $volumeFullYear->{'Reactive power charge p/kVArh'}
                  )
                : (),
                map {
                    my $name = "Unit rate $_ p/kWh";
                    "A$_"     => $tariffsExMatching->{$name},
                      "A90$_" => $volumeFullYear->{$name};
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
                $model->{dcp123doublecounting} ? 'A701*A702' : (),
                map { "A$_*A90$_" } 1 .. $model->{maxUnitRates} )
              . ')',
            arguments => {
                A5 => $loadCoefficients,
                $model->{dcp123doublecounting}
                ? (
                    A701 =>
                      $tariffsExMatching->{'Reactive power charge p/kVArh'},
                    A702 => $volumeFullYear->{'Reactive power charge p/kVArh'}
                  )
                : (),
                map {
                    my $name = "Unit rate $_ p/kWh";
                    "A$_"     => $tariffsExMatching->{$name},
                      "A90$_" => $volumeFullYear->{$name};
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
            arithmetic => '=0.01*A6*A1*A2',       # '=IF(A5<0,0,0.01*A6*A1*A2)'
            arguments  => {
                A5 => $loadCoefficients,
                A6 => $daysFullYear,
                A1 => $tariffsExMatching->{'Fixed charge p/MPAN/day'},
                A2 => $volumeFullYear->{'Fixed charge p/MPAN/day'},
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
            arithmetic => '=IF(A5<0,0,0.01*A6*A1*A2)',
            arguments  => {
                A5 => $loadCoefficients,
                A6 => $daysFullYear,
                A1 => $tariffsExMatching->{'Capacity charge p/kVA/day'},
                A2 => $volumeFullYear->{'Capacity charge p/kVA/day'},
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
            arithmetic => '=10*A1*A2',
            arguments  => {
                A5 => $loadCoefficients,
                A1 => $tariffsExMatching->{'Reactive power charge p/kVArh'},
                A2 => $volumeFullYear->{'Reactive power charge p/kVArh'},
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
                arithmetic    => '=A1*A2/(A3+A4+A5+A6)',
                arguments     => {
                    A1 => $adderAmount,
                    A2 => $totalRevenueFromUnits,
                    A3 => $totalRevenueFromUnits,
                    A4 => $totalRevenueFromFixed,
                    A5 => $totalRevenueFromCapacity,
                    A6 => $totalRevenueFromReactive,
                }
            ),
            Arithmetic(
                name => 'Revenue matching target from fixed charges (£/year)',
                defaultFormat => '0soft',
                arithmetic    => '=A1*A2/(A3+A4+A5+A6)',
                arguments     => {
                    A1 => $adderAmount,
                    A2 => $totalRevenueFromFixed,
                    A3 => $totalRevenueFromUnits,
                    A4 => $totalRevenueFromFixed,
                    A5 => $totalRevenueFromCapacity,
                    A6 => $totalRevenueFromReactive,
                }
            ),
            Arithmetic(
                name =>
                  'Revenue matching target from capacity charges (£/year)',
                defaultFormat => '0soft',
                arithmetic    => '=A1*A2/(A3+A4+A5+A6)',
                arguments     => {
                    A1 => $adderAmount,
                    A2 => $totalRevenueFromCapacity,
                    A3 => $totalRevenueFromUnits,
                    A4 => $totalRevenueFromFixed,
                    A5 => $totalRevenueFromCapacity,
                    A6 => $totalRevenueFromReactive,
                }
            ),
            Arithmetic(
                name =>
'Revenue matching target from reactive power charges (£/year)',
                defaultFormat => '0soft',
                arithmetic    => '=A1*A2/(A3+A4+A5+A6)',
                arguments     => {
                    A1 => $adderAmount,
                    A2 => $totalRevenueFromReactive,
                    A3 => $totalRevenueFromUnits,
                    A4 => $totalRevenueFromFixed,
                    A5 => $totalRevenueFromCapacity,
                    A6 => $totalRevenueFromReactive,
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
                arithmetic => '=IF(A3<0,0,A1*10)',
                arguments  => {
                    A3 => $loadCoefficients,
                    A2 => $daysAfter,
                    A1 => $volumeAfter->{$_},
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
                arithmetic => $minValue ? '=IF(A4<0,0,A9-A1)'
                : '=IF(A4<0,0,0-A1)',
                arguments => {
                    A4 => $loadCoefficients,
                    A1 => $tariffsExMatching->{$_},
                    $minValue ? ( A9 => $minValue ) : (),
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
                arithmetic => '=IF(A3<0,0,MAX(A6,A1))',
                arguments  => {
                    A1 => $adderRate,
                    A3 => $loadCoefficients,
                    A6 => $minAdder{$_},
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
                    name => "Effect through $_",
                    arithmetic =>
                      '=IF(A9>0,A1*A2*0.01,0)',    # '=IF(A3>0,A1*A2*0.01,0)',
                    arguments => {
                        A3 => $loadCoefficients,
                        A2 => $daysAfter,
                        A1 => $volumeAfter->{$_},
                        A9 => $tariffsExMatching->{$_},
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
                    arithmetic => '=0-A1',                 # '=IF(A4<0,0,0-A1)',
                    arguments  => {
                        A4 => $loadCoefficients,
                        A1 => $tariffsExMatching->{$_},
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
                    name => "Adder on $_",
                    rows => $allTariffsByEndUser,
                    cols => Labelset( list => ['Adder'] ),
                    arithmetic =>
                      '=IF(A9>0,MAX(A6,A1),0)',    # '=IF(A3<0,0,MAX(A6,A1))',
                    arguments => {
                        A1 => $adderRate,
                        A3 => $loadCoefficients,
                        A6 => $minAdder{$_},
                        A9 => $tariffsExMatching->{$_},
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
                    arithmetic => '=IF(A3<0,0,A1*A2*0.01)',
                    arguments  => {
                        A3 => $loadCoefficients,
                        A2 => $daysAfter,
                        A1 => $volumeAfter->{$_},
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
                    arithmetic => '=IF(A4<0,0,0-A1)',
                    arguments  => {
                        A4 => $loadCoefficients,
                        A1 => $tariffsExMatching->{$_},
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
                    arithmetic => '=IF(A3<0,0,MAX(A6,A1))',
                    arguments  => {
                        A1 => $adderRate,
                        A3 => $loadCoefficients,
                        A6 => $minAdder{$_},
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

    if ( $hybridTargets[3] ) {    # reactive

        my @columns = grep { /kVArh/ } @$nonExcludedComponents;

        my @slope = map {
            Arithmetic(
                name       => "Effect through $_",
                arithmetic => '=A1*10',              # '=IF(A3<0,0,A1*10)',
                arguments  => {
                    A3 => $loadCoefficients,
                    A1 => $volumeAfter->{$_},
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
                arithmetic => '=0-A1',                    # '=IF(A4<0,0,0-A1)',
                arguments  => {
                    A4 => $loadCoefficients,
                    A1 => $tariffsExMatching->{$_},
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
                arithmetic => '=MAX(A6,A1)',    # '=IF(A3<0,0,MAX(A6,A1))',
                arguments  => {
                    A1 => $adderRate,
                    A3 => $loadCoefficients,
                    A6 => $minAdder{$_},
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
        my %args = ( A400 => $daysAfter );
        my $i = 1;

        foreach ( grep { $adderTable->{$_} } @$nonExcludedComponents ) {
            ++$i;
            my $pad = "$i";
            $pad = "0$pad" while length $pad < 3;
            if (m#/day#) {
                push @termsWithDays, "A2$pad*A3$pad";
            }
            else {
                push @termsNoDays, "A2$pad*A3$pad";
            }
            $args{"A2$pad"} = $adderTable->{$_};
            $args{"A3$pad"} = $volumeAfter->{$_};
        }

        $revenuesFromAdder = Arithmetic(
            name       => 'Net revenues by tariff from adder',
            rows       => $allTariffsByEndUser,
            arithmetic => '='
              . join( '+',
                @termsWithDays
                ? ( '0.01*A400*(' . join( '+', @termsWithDays ) . ')' )
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
