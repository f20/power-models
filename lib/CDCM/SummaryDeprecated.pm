package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others.
Copyright 2011-2017 Franck Latrémolière, Reckon LLP and others.

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

sub consultationSummaryDeprecated {

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

    if ( $currentTariffs && !$model->{force1201} ) {
        my @termsNoDays;
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
            }
            $args{"A2$pad"} = $currentTariffs->{$_};
            $args{"A3$pad"} = $volumeData->{$_};
        }
        $currentRevenues = Arithmetic(
            name          => 'Revenues under current tariffs (£)',
            defaultFormat => '0softnz',
            rows          => $selectedTariffsForComparison,
            arithmetic    => '='
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
    }
    else {
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
                A1   => $hardRevenues,
                A9   => $hardRevenues,
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

    push @{ $model->{consultationTables} }, Columnset(
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
                            !/(?:LD|Q)NO/i
                              && $componentMap->{$_}
                              {'Reactive power charge p/kVArh'}
                          } map { $allTariffs->{list}[$_] }
                          $allTariffs->indices
                    ]
                ),
                sources => [ $tariffTable->{'Reactive power charge p/kVArh'} ]
            )
        ]
      ) if $model->{summary} =~ /big/i;

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
                arithmetic => '=A1/IF(A4="kVA",IF(A51,A52,1),'
                  . 'IF(A3="MPAN",IF(A6,A9,1),IF(A7,A8,1)))'
                  . ( /Unit rate/i && $unitsLossAdjustment ? '/(1+A2)' : '' ),
                arguments => {
                    A1 => $atwVolumes{$_},
                    $unitsLossAdjustment ? ( A2 => $unitsLossAdjustment ) : (),
                    A3  => $normalisedTo,
                    A4  => $normalisedTo,
                    A51 => $atwVolumes{'Capacity charge p/kVA/day'},
                    A52 => $atwVolumes{'Capacity charge p/kVA/day'},
                    A6  => $atwVolumes{'Fixed charge p/MPAN/day'},
                    A7  => $atwUnits,
                    A8  => $atwUnits,
                    A9  => $atwVolumes{'Fixed charge p/MPAN/day'}
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
                arithmetic => '=A1/'
                  . 'IF(A3="MPAN",IF(A6,A9,1),IF(A7,A8,1))'
                  . ( /Unit rate/i && $unitsLossAdjustment ? '/(1+A2)' : '' ),
                arguments => {
                    A1 => $atwVolumes{$_},
                    $unitsLossAdjustment ? ( A2 => $unitsLossAdjustment ) : (),
                    A3 => $normalisedTo,
                    A4 => $normalisedTo,
                    A6 => $atwVolumes{'Fixed charge p/MPAN/day'},
                    A7 => $atwUnits,
                    A8 => $atwUnits,
                    A9 => $atwVolumes{'Fixed charge p/MPAN/day'}
                }
            );
        } @$nonExcludedComponents
      );

    my $rev;

    {
        my @termsNoDays;
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
            }
            $args{"A2$pad"} = $tariffTable->{$_};
            $args{"A3$pad"} = $adjustedVolume{$_};
        }
        $rev = Arithmetic(
            name          => 'Normalised revenues (£)',
            rows          => $tariffsByAtw,
            defaultFormat => '0softnz',
            arithmetic    => '='
              . join( '+',
                @termsWithDays
                ? ( '0.01*A400*(' . join( '+', @termsWithDays ) . ')' )
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
        next unless grep { /^(?:LD|Q)NO $idnoType:/ } @{ $allTariffs->{list} };
        push @{ $atwMarginTariffs->{accepts} }, my $tariffset = Labelset(
            list => [
                map {
                    ( grep { /^(?:LD|Q)NO $idnoType:/ } @{ $_->{list} } )[0]
                      || 'N/A'
                } @{ $atwMarginTariffs->{list} }
            ],
        );
        push @{ $model->{consultationInput} },
          my $irev = Stack(
            name    => "$model->{ldnoWord} $idnoType charges (normalised £)",
            rows    => $tariffset,
            sources => [$rev]
          );
        push @rev,
          Arithmetic(
            name => "$model->{ldnoWord} $idnoType margin (normalised £)",
            defaultFormat => '0.00soft',
            rowFormats    => [
                map { $_ eq 'N/A' ? 'unavailable' : undef }
                  @{ $tariffset->{list} }
            ],
            arithmetic => '=IF(A3,A1-A2,"")',
            arguments  => { A1 => $atwRev, A2 => $irev, A3 => $irev }
          );
    }

    push @{ $model->{consultationTables} }, Columnset(
        name    => 'Comparison with current all-the-way demand tariffs',
        columns => [

            $model->{summary} =~ /jun/i
            ? (    # June 2009
                Arithmetic(
                    name          => 'Revenues (£m)',
                    defaultFormat => '0.00softnz',
                    rows          => $selectedTariffsForComparison,
                    arithmetic    => '=A1*1e-6',
                    arguments     => { A1 => $revenuesFromTariffs }
                ),
                Stack(
                    sources => [$normalisedTo],
                    rows    => $selectedTariffsForComparison,
                ),
                Arithmetic(
                    name          => 'Normalised units (MWh)',
                    defaultFormat => '0.000softnz',
                    rows          => $selectedTariffsForComparison,
                    arithmetic    => '=A1/IF(A4="kVA",IF(A51,A52,1),'
                      . 'IF(A3="MPAN",IF(A6,A9,1),IF(A7,A8,1)))',
                    arguments => {
                        A1  => $unitsInYear,
                        A3  => $normalisedTo,
                        A4  => $normalisedTo,
                        A51 => $volumeData->{'Capacity charge p/kVA/day'},
                        A52 => $volumeData->{'Capacity charge p/kVA/day'},
                        A6  => $volumeData->{'Fixed charge p/MPAN/day'},
                        A7  => $unitsInYear,
                        A8  => $unitsInYear,
                        A9  => $volumeData->{'Fixed charge p/MPAN/day'}
                    }
                ),
                Arithmetic(
                    name          => 'Normalised revenue (£)',
                    rows          => $selectedTariffsForComparison,
                    defaultFormat => '0.00softnz',
                    arithmetic    => '=A1/IF(A4="kVA",IF(A51,A52,1),'
                      . 'IF(A3="MPAN",IF(A6,A9,1),IF(A7,A8,1)))',
                    arguments => {
                        A1  => $revenuesFromTariffs,
                        A3  => $normalisedTo,
                        A4  => $normalisedTo,
                        A51 => $volumeData->{'Capacity charge p/kVA/day'},
                        A52 => $volumeData->{'Capacity charge p/kVA/day'},
                        A6  => $volumeData->{'Fixed charge p/MPAN/day'},
                        A7  => $unitsInYear,
                        A8  => $unitsInYear,
                        A9  => $volumeData->{'Fixed charge p/MPAN/day'}
                    }
                ),
              )
            : $model->{summary} !~ /impact/i ? (    # default
              )
            : (                                     # impact option
                Arithmetic(
                    name          => 'Revenue impact (£m)',
                    defaultFormat => '0.00softpm',
                    arithmetic    => '=1e-6*(A2-A1)',
                    arguments     => {
                        A1 => $currentRevenues,
                        A2 => $revenuesFromTariffs,
                    }
                ),
                Arithmetic(
                    name          => 'Current p/kWh',
                    rows          => $selectedTariffsForComparison,
                    defaultFormat => '0.000softnz',
                    arithmetic    => '=A1/IF(A7,A8,1)/10',
                    arguments     => {
                        A1 => $currentRevenues,
                        A7 => $unitsInYear,
                        A8 => $unitsInYear,
                    }
                ),
                Arithmetic(
                    name          => 'Proposed p/kWh',
                    rows          => $selectedTariffsForComparison,
                    defaultFormat => '0.000softnz',
                    arithmetic    => '=A1/IF(A7,A8,1)/10',
                    arguments     => {
                        A1 => $revenuesFromTariffs,
                        A7 => $unitsInYear,
                        A8 => $unitsInYear,
                    }
                ),
            ),
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
            $model->{summary} =~ /jun/i
            ? (    # June 2009
              )
            : (    # default
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
            ),
          ]

    );

    push @{ $model->{consultationTables} },
      Columnset(
        name    => "$model->{ldnoWord} margins in use of system charges",
        columns => \@rev
      );

}

1;
