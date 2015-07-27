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

=head Development note

For some configurations, this uses $model->{inYear...} to communicate with InYearSummaries behind the back of Master:
    $model->{inYearVolumes}  = [ $volumeDataBefore1, $volumeDataBefore2, ];
    $model->{inYearTariffs}  = [ $tariffsBefore1,    $tariffsBefore2, ];
    $model->{inYearRevenues} = [ $revenuesBefore1,   $revenuesBefore2, ];

=cut

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Shortcuts ':all';

sub inYearAdjustUsingBefore {

    my (
        $model,           $nonExcludedComponents, $volumeData,
        $allEndUsers,     $componentMap,          $daysAfter,
        $daysBefore,      $daysInYear,            $revenueBefore,
        $revenuesBefore,  $tariffsBefore,         $unitsInYearAfter,
        $volumeDataAfter, $volumesAdjustedAfter,
    ) = @_;

    if ( $model->{inYear} =~ /twice/ ) {

        my $volumeDataBefore1 = {
            map {
                $_ => Dataset(
                    name => SpreadsheetModel::Object::_shortName(
                        $volumeData->{$_}{name}
                    ),
                    validation => {
                        validate    => 'decimal',
                        criteria    => '>=',
                        value       => 0,
                        error_title => 'Volume data error',
                        error_message =>
                          'The volume must be a non-negative number.'
                    },
                    defaultFormat => $volumeData->{$_}{defaultFormat},
                    rows          => $volumeData->{$_}{rows},
                    data          => [
                        map { defined $_ ? 0 : undef }
                          @{ $volumeData->{$_}{data} }
                    ],
                  )
            } @$nonExcludedComponents
        };

        Columnset(
            name => 'If modelling a second in-year tariff change, '
              . 'volumes within the charging year to which tariffs in table 1097 apply (if any)',
            lines => [
                    'Leave this table blank unless '
                  . 'you are modelling a second in-year tariff change.'
            ],
            number   => 1098,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            columns  => [ @{$volumeDataBefore1}{@$nonExcludedComponents} ]
        );

        my $volumeDataBefore2 = {
            map {
                $_ => Dataset(
                    name => SpreadsheetModel::Object::_shortName(
                        $volumeData->{$_}{name}
                    ),
                    validation => {
                        validate    => 'decimal',
                        criteria    => '>=',
                        value       => 0,
                        error_title => 'Volume data error',
                        error_message =>
                          'The volume must be a non-negative number.'
                    },
                    defaultFormat => $volumeData->{$_}{defaultFormat},
                    rows          => $volumeData->{$_}{rows},
                    data          => [
                        map { defined $_ ? 0 : undef }
                          @{ $volumeData->{$_}{data} }
                    ],
                  )
            } @$nonExcludedComponents
        };

        Columnset(
            name => 'If modelling an in-year tariff change, '
              . 'volumes within the charging year to which tariffs in table 1095 apply (if any)',
            lines => [
'Leave this table blank if setting tariffs for a full financial year.'
            ],
            number   => 1096,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            columns  => [ @{$volumeDataBefore2}{@$nonExcludedComponents} ]
        );

        $$volumeDataAfter = {
            map {
                $_ => /\/day/
                  ? Arithmetic(
                    name => SpreadsheetModel::Object::_shortName(
                        $volumeData->{$_}{name}
                    ),
                    arithmetic => '=(A7*A1-A83*A21-A84*A22)/A9',
                    arguments  => {
                        A1  => $volumeData->{$_},
                        A21 => $volumeDataBefore1->{$_},
                        A22 => $volumeDataBefore2->{$_},
                        A7  => $daysInYear,
                        A83 => $daysBefore->[0],
                        A84 => $daysBefore->[1],
                        A9  => $daysAfter,
                    },
                    defaultFormat => '0soft',
                  )
                  : Arithmetic(
                    name => SpreadsheetModel::Object::_shortName(
                        $volumeData->{$_}{name}
                    ),
                    arithmetic => '=A1-A21-A22',
                    arguments  => {
                        A1  => $volumeData->{$_},
                        A21 => $volumeDataBefore1->{$_},
                        A22 => $volumeDataBefore2->{$_},
                    },
                  );
            } @$nonExcludedComponents
        };

        my $tariffsBefore1 = {
            map {
                my $tariffComponent = $_;
                $_ => Dataset(
                    name       => $tariffComponent,
                    validation => {
                        validate => 'decimal',
                        criteria => 'between',
                        minimum  => -999_999.999,
                        maximum  => 999_999.999,
                    },
                    rows          => $volumeData->{$_}{rows},
                    defaultFormat => /k(?:W|VAr)h/ ? '0.000hard'
                    : '0.00hard',
                    rowFormats => [
                        map {
                            $componentMap->{$_}{$tariffComponent} ? undef
                              : 'unavailable';
                        } @{ $volumeData->{$_}{rows}{list} }
                    ],
                    data => [
                        [
                            map {
                                $componentMap->{$_}{$tariffComponent}
                                  ? 0
                                  : undef;
                            } @{ $volumeData->{$_}{rows}{list} }
                        ]
                    ]
                );
            } @$nonExcludedComponents
        };

        Columnset(
            name =>
'If modelling a second in-year tariff change, tariffs that applied before the first in-year tariff change',
            lines => [
'This table is only used when modelling a second in-year change.'
            ],
            number   => 1097,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            columns  => [ @{$tariffsBefore1}{@$nonExcludedComponents} ]
        );

        my $tariffsBefore2 = $$tariffsBefore = {
            map {
                my $tariffComponent = $_;
                $_ => Dataset(
                    name       => $tariffComponent,
                    validation => {
                        validate => 'decimal',
                        criteria => 'between',
                        minimum  => -999_999.999,
                        maximum  => 999_999.999,
                    },
                    rows          => $volumeData->{$_}{rows},
                    defaultFormat => /k(?:W|VAr)h/ ? '0.000hard'
                    : '0.00hard',
                    rowFormats => [
                        map {
                            $componentMap->{$_}{$tariffComponent} ? undef
                              : 'unavailable';
                        } @{ $volumeData->{$_}{rows}{list} }
                    ],
                    data => [
                        [
                            map {
                                $componentMap->{$_}{$tariffComponent}
                                  ? 0
                                  : undef;
                            } @{ $volumeData->{$_}{rows}{list} }
                        ]
                    ]
                );
            } @$nonExcludedComponents
        };

        Columnset(
            name =>
'Current tariffs (those in force immediately before the tariffs calculated by this model would come into effect)',
            number   => 1095,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            columns  => [ @{$tariffsBefore2}{@$nonExcludedComponents} ]
        );

        $model->{inYearVolumes} = [ $volumeDataBefore1, $volumeDataBefore2, ];
        $model->{inYearTariffs} = [ $tariffsBefore1,    $tariffsBefore2, ];

        # NB: "override" is dead
        if ( !$model->{targetRevenue} && $model->{inYear} =~ /adjust/ ) {
            my @termsNoDays1;
            my @termsNoDays2;
            my @termsWithDays1;
            my @termsWithDays2;
            my %args = ( A400 => $daysBefore->[0], A700 => $daysBefore->[1] );
            my $i = 1;
            my $rows;
            foreach (@$nonExcludedComponents) {
                ++$i;
                my $pad = "$i";
                $pad = "0$pad" while length $pad < 3;
                if (m#/day#) {
                    push @termsWithDays1, "A2$pad*A3$pad";
                    push @termsWithDays2, "A5$pad*A6$pad";
                }
                else {
                    push @termsNoDays1, "A2$pad*A3$pad";
                    push @termsNoDays2, "A5$pad*A6$pad";
                }
                $args{"A2$pad"} = $tariffsBefore1->{$_};
                $args{"A3$pad"} = $volumeDataBefore1->{$_};
                $args{"A5$pad"} = $tariffsBefore2->{$_};
                $args{"A6$pad"} = $volumeDataBefore2->{$_};
                $rows ||= $volumeDataBefore1->{$_}{rows}
                  || $volumeDataBefore2->{$_}{rows};
            }
            $model->{inYearRevenues} = [
                Arithmetic(
                    name =>
                      "Revenues in the period covered by tables 1097/1098 (£)",
                    defaultFormat => '0soft',
                    rows          => $rows,
                    arithmetic    => '='
                      . join(
                        '+',
                        @termsWithDays1
                        ? ( '0.01*A400*('
                              . join( '+', @termsWithDays1 )
                              . ')' )
                        : ('0'),
                        @termsNoDays1
                        ? ( '10*(' . join( '+', @termsNoDays1 ) . ')' )
                        : ('0'),
                      ),
                    arguments => \%args,
                ),
                Arithmetic(
                    name =>
                      "Revenues in the period covered by tables 1095/1096 (£)",
                    defaultFormat => '0soft',
                    rows          => $rows,
                    arithmetic    => '='
                      . join(
                        '+',
                        @termsWithDays2
                        ? ( '0.01*A700*('
                              . join( '+', @termsWithDays2 )
                              . ')' )
                        : ('0'),
                        @termsNoDays2
                        ? ( '10*(' . join( '+', @termsNoDays2 ) . ')' )
                        : ('0'),
                      ),
                    arguments => \%args,
                )
            ];
            push @{ $model->{revenueMatching} },
              my @revBef = (
                GroupBy(
                    name =>
"Total net revenue in the period covered by tables 1097/1098 (£)",
                    defaultFormat => '0soft',
                    source        => $model->{inYearRevenues}[0],
                ),
                GroupBy(
                    name =>
"Total net revenue in the period covered by tables 1095/1096 (£)",
                    defaultFormat => '0soft',
                    source        => $model->{inYearRevenues}[1],
                )
              );
            push @{ $model->{adjustRevenuesBefore} },
              Dataset(
                name =>
'Adjustment to net revenue in the period covered by tables 1095/1096 (£)',
                defaultFormat => '0hard',
                data          => [ [0] ],
              ),
              Dataset(
                name =>
'Adjustment to net revenue in the period covered by tables 1097/1098 (£)',
                defaultFormat => '0hard',
                data          => [ [0] ],
              );
            $$revenueBefore = Arithmetic(
                name =>
'Total net revenue in the charging year before the tariff change (£)',
                defaultFormat => '0soft',
                arithmetic    => '=A1+A3+A4+A6',
                arguments     => {
                    A1 => $model->{adjustRevenuesBefore}[0],
                    A3 => $revBef[0],
                    A4 => $model->{adjustRevenuesBefore}[1],
                    A6 => $revBef[1],
                },
            );

        }

        else

        {
            my @termsNoDays;
            my @termsWithDays1;
            my @termsWithDays2;
            my %args = ( A400 => $daysBefore->[0], A700 => $daysBefore->[1] );
            my $i = 1;
            my $rows;
            foreach (@$nonExcludedComponents) {
                ++$i;
                my $pad = "$i";
                $pad = "0$pad" while length $pad < 3;
                if (m#/day#) {
                    push @termsWithDays1, "A2$pad*A3$pad";
                    push @termsWithDays2, "A5$pad*A6$pad";
                }
                else {
                    push @termsNoDays, "A2$pad*A3$pad";
                    push @termsNoDays, "A5$pad*A6$pad";
                }
                $args{"A2$pad"} = $tariffsBefore1->{$_};
                $args{"A3$pad"} = $volumeDataBefore1->{$_};
                $args{"A5$pad"} = $tariffsBefore2->{$_};
                $args{"A6$pad"} = $volumeDataBefore2->{$_};
                $rows ||= $volumeDataBefore1->{$_}{rows}
                  || $volumeDataBefore2->{$_}{rows};
            }
            $$revenuesBefore = Arithmetic(
                name =>
'Revenues in charging year before tariffs calculated in this model would come into effect (£)',
                defaultFormat => '0soft',
                rows          => $rows,
                arithmetic    => '='
                  . join( '+',
                    @termsWithDays1
                    ? ( '0.01*A400*(' . join( '+', @termsWithDays1 ) . ')' )
                    : ('0'),
                    @termsWithDays2
                    ? ( '0.01*A700*(' . join( '+', @termsWithDays2 ) . ')' )
                    : ('0'),
                    @termsNoDays ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
                    : ('0'),
                  ),
                arguments => \%args,
            );
            $$revenueBefore = GroupBy(
                name =>
'Total net revenue in charging year before tariffs calculated in this model would come into effect (£)',
                defaultFormat => '0soft',
                source        => $$revenuesBefore
            );
        }

    }

    else {

        my $volumeDataBefore = {
            map {
                $_ => Dataset(
                    name       => $volumeData->{$_}{name},
                    validation => {
                        validate    => 'decimal',
                        criteria    => '>=',
                        value       => 0,
                        error_title => 'Volume data error',
                        error_message =>
                          'The volume must be a non-negative number.'
                    },
                    defaultFormat => $volumeData->{$_}{defaultFormat},
                    rows          => $volumeData->{$_}{rows},
                    data          => [
                        map { defined $_ ? 0 : undef }
                          @{ $volumeData->{$_}{data} }
                    ],
                  )
            } @$nonExcludedComponents
        };

        Columnset(
            name =>
'Volume forecasts for the part of the charging year before the tariff change (if any)',
            lines => [
'Leave this table blank if setting tariffs for a full financial year.'
            ],
            number   => 1054,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            columns  => [ @{$volumeDataBefore}{@$nonExcludedComponents} ]
        );

        if ( $model->{inYear} =~ /multiple/ ) {

            $$revenueBefore = Dataset(
                name =>
'Total net revenue in charging year before tariffs calculated in this model would come into effect (£)',
                defaultFormat => '0hard',
                data          => [ [0] ],
                number        => 1078,
                appendTo      => $model->{inputTables},
                dataset       => $model->{dataset},
            );

        }

        else {

            $$tariffsBefore = {
                map {
                    my $tariffComponent = $_;
                    $_ => Dataset(
                        name       => $tariffComponent,
                        validation => {
                            validate => 'decimal',
                            criteria => 'between',
                            minimum  => -999_999.999,
                            maximum  => 999_999.999,
                        },
                        rows          => $volumeData->{$_}{rows},
                        defaultFormat => /k(?:W|VAr)h/ ? '0.000hard'
                        : '0.00hard',
                        rowFormats => [
                            map {
                                $componentMap->{$_}{$tariffComponent} ? undef
                                  : 'unavailable';
                            } @{ $volumeData->{$_}{rows}{list} }
                        ],
                        data => [
                            [
                                map {
                                    $componentMap->{$_}{$tariffComponent}
                                      ? 0
                                      : undef;
                                } @{ $volumeData->{$_}{rows}{list} }
                            ]
                        ]
                    );
                } @$nonExcludedComponents
            };

            Columnset(
                name =>
'Tariffs for the part of the charging year before the tariff change (if any)',
                lines => [
'Leave this table blank if setting tariffs for a full financial year.'
                ],
                number   => 1095,
                appendTo => $model->{inputTables},
                dataset  => $model->{dataset},
                columns  => [ @{$$tariffsBefore}{@$nonExcludedComponents} ]
            );

            my @termsNoDays;
            my @termsWithDays;
            my %args = ( A400 => $daysBefore );
            my $i = 1;
            my $rows;
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
                $args{"A2$pad"} = $$tariffsBefore->{$_};
                $args{"A3$pad"} = $volumeDataBefore->{$_};
                $rows ||= $volumeDataBefore->{$_}{rows};
            }
            $$revenuesBefore = Arithmetic(
                name =>
'Revenues in charging year before tariffs calculated in this model would come into effect (£)',
                defaultFormat => '0soft',
                rows          => $rows,
                arithmetic    => '='
                  . join( '+',
                    @termsWithDays
                    ? ( '0.01*A400*(' . join( '+', @termsWithDays ) . ')' )
                    : ('0'),
                    @termsNoDays ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
                    : ('0'),
                  ),
                arguments => \%args,
            );
            $$revenueBefore = GroupBy(
                name =>
'Total net revenue in charging year before tariffs calculated in this model would come into effect (£)',
                defaultFormat => '0soft',
                source        => $$revenuesBefore
            );

            $model->{inYearVolumes} = [ $volumeDataBefore, ];
            $model->{inYearTariffs} = [ $$tariffsBefore, ];

        }

        $$volumeDataAfter = {
            map {
                $_ => /\/day/
                  ? Arithmetic(
                    name => SpreadsheetModel::Object::_shortName(
                        $volumeData->{$_}{name}
                    ),
                    arithmetic => '=(A7*A1-A8*A2)/A9',
                    arguments  => {
                        A1 => $volumeData->{$_},
                        A2 => $volumeDataBefore->{$_},
                        A7 => $daysInYear,
                        A8 => $daysBefore,
                        A9 => $daysAfter,
                    },
                    defaultFormat => '0soft',
                  )
                  : Arithmetic(
                    name => SpreadsheetModel::Object::_shortName(
                        $volumeData->{$_}{name}
                    ),
                    arithmetic => '=A1-A2',
                    arguments  => {
                        A1 => $volumeData->{$_},
                        A2 => $volumeDataBefore->{$_},
                    },
                  );
            } @$nonExcludedComponents
        };

    }

    push @{ $model->{volumeData} },
      $$unitsInYearAfter = Arithmetic(
        defaultFormat => '0softnz',
        name          => 'All units after tariff change (MWh)',
        arithmetic    => '='
          . join( '+', map { "A$_" } 1 .. $model->{maxUnitRates} ),
        arguments => {
            map { ( "A$_" => $$volumeDataAfter->{"Unit rate $_ p/kWh"} ) }
              1 .. $model->{maxUnitRates}
        },
        defaultFormat => '0softnz'
      );

    Columnset(
        name =>
'Volumes in period to which tariffs calculated in this model would apply',
        columns => [ @{$$volumeDataAfter}{@$nonExcludedComponents} ]
    );

    if ($volumesAdjustedAfter) {

        my %intermediateAfter = map {
            $_ => Arithmetic(
                name => SpreadsheetModel::Object::_shortName(
                    $$volumeDataAfter->{$_}{name}
                ),
                arithmetic => '=A1*(1-A2)',
                arguments  => {
                    A1 => $$volumeDataAfter->{$_},
                    A2 => /fix/i
                    ? $model->{pcd}{discountFixed}
                    : $model->{pcd}{discount}
                }
            );
        } @$nonExcludedComponents;

        Columnset(
            name =>
'Volumes in period to which tariffs calculated in this model would apply, adjusted for IDNO discounts',
            columns => [ @intermediateAfter{@$nonExcludedComponents} ]
        );

        $$volumesAdjustedAfter = {
            map {
                $_ => GroupBy(
                    name => SpreadsheetModel::Object::_shortName(
                        $intermediateAfter{$_}{name}
                    ),
                    rows   => $allEndUsers,
                    source => $intermediateAfter{$_}
                );
            } @$nonExcludedComponents
        };

        push @{ $model->{volumeData} },
          Columnset(
            name =>
'Equivalent volume for each end user, in period to which new tariffs are to apply',
            columns => [ @{$$volumesAdjustedAfter}{@$nonExcludedComponents} ]
          );

    }

}

sub inYearAdjustUsingAfter {

    my (
        $model,            $nonExcludedComponents, $volumeData,
        $allEndUsers,      $componentMap,          $revenueBefore,
        $unitsInYearAfter, $volumeDataAfter,       $volumesAdjustedAfter,
    ) = @_;

    $$volumeDataAfter = {
        map {
            $_ => Dataset(
                name       => $volumeData->{$_}{name},
                validation => {
                    validate      => 'decimal',
                    criteria      => '>=',
                    value         => 0,
                    error_title   => 'Volume data error',
                    error_message => 'The volume must be a non-negative number.'
                },
                defaultFormat => $volumeData->{$_}{defaultFormat},
                rows          => $volumeData->{$_}{rows},
                data          => [
                    map { defined $_ ? 0 : undef } @{ $volumeData->{$_}{data} }
                ],
              )
        } @$nonExcludedComponents
    };

    Columnset(
        name     => 'Volumes to which the new tariffs would apply',
        number   => 1054,
        appendTo => $model->{inputTables},
        dataset  => $model->{dataset},
        columns  => [ @{$$volumeDataAfter}{@$nonExcludedComponents} ]
    );

    $$revenueBefore = Dataset(
        name =>
'Total net revenue in the charging year before the tariffs calculated in this model would come into effect (£)',
        defaultFormat => '0hard',
        data          => [ [0] ],
        number        => 1079,
        appendTo      => $model->{inputTables},
        dataset       => $model->{dataset},
    ) if $revenueBefore;

    push @{ $model->{volumeData} },
      $$unitsInYearAfter = Arithmetic(
        noCopy        => 1,
        defaultFormat => '0softnz',
        name          => 'All units after tariff change (MWh)',
        arithmetic    => '='
          . join( '+', map { "A$_" } 1 .. $model->{maxUnitRates} ),
        arguments => {
            map { ( "A$_" => $$volumeDataAfter->{"Unit rate $_ p/kWh"} ) }
              1 .. $model->{maxUnitRates}
        },
        defaultFormat => '0softnz'
      );

    if ( $model->{pcd} ) {
        my %intermediateAfter = map {
            $_ => Arithmetic(
                name => SpreadsheetModel::Object::_shortName(
                    $$volumeDataAfter->{$_}{name}
                ),
                arithmetic => '=A1*(1-A2)',
                arguments  => {
                    A1 => $$volumeDataAfter->{$_},
                    A2 => /fix/i
                    ? $model->{pcd}{discountFixed}
                    : $model->{pcd}{discount}
                }
            );
        } @$nonExcludedComponents;

        Columnset(
            name =>
'Volumes in period to which new tariffs apply, adjusted for IDNO discounts',
            columns => [ @intermediateAfter{@$nonExcludedComponents} ]
        );

        $$volumesAdjustedAfter = {
            map {
                $_ => GroupBy(
                    name => SpreadsheetModel::Object::_shortName(
                        $intermediateAfter{$_}{name}
                    ),
                    rows   => $allEndUsers,
                    source => $intermediateAfter{$_}
                );
            } @$nonExcludedComponents
        };

        push @{ $model->{volumeData} },
          Columnset(
            name =>
'Equivalent volume for each end user, in period to which new tariffs are to apply',
            columns => [ @{$$volumesAdjustedAfter}{@$nonExcludedComponents} ]
          );
    }

}

sub displayWholeYearTarget {

    my (
        $model,                 $nonExcludedComponents,
        $daysInYear,            $volumeData,
        $tariffsBeforeRounding, $allowedRevenue,
        $revenueFromElsewhere,  $siteSpecificCharges,
    ) = @_;

    my @termsNoDays;
    my @termsWithDays;
    my %args = ( A400 => $daysInYear );
    my $i = 1;
    my $rows;
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
        $args{"A2$pad"} = $tariffsBeforeRounding->{$_};
        $args{"A3$pad"} = $volumeData->{$_};
        $rows ||= $volumeData->{$_}{rows};
    }
    push @{ $model->{revenueSummary} },
      my $annualisedRevenue = GroupBy(
        name => 'Total net revenue if applied to the whole year (£/year)',
        defaultFormat => '0soft',
        source        => Arithmetic(
            name          => 'Revenue if applied to the whole year (£/year)',
            defaultFormat => '0soft',
            rows          => $rows,
            arithmetic    => '='
              . join( '+',
                @termsWithDays
                ? ( '0.01*A400*(' . join( '+', @termsWithDays ) . ')' )
                : ('0'),
                @termsNoDays ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
                : ('0'),
              ),
            arguments => \%args,
        )
      );
    push @{ $model->{revenueSummary} },
      Arithmetic(
        name => 'Suggested adjustment to model 100 target revenue (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=A1-A2' . ( $revenueFromElsewhere ? '+A3' : '' ),
        arguments     => {
            A1 => $annualisedRevenue,
            A2 => $allowedRevenue,
            $revenueFromElsewhere ? ( A3 => $revenueFromElsewhere ) : (),
        }
      ) unless $siteSpecificCharges;

}

1;
