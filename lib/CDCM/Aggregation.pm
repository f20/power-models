package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others.

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

# Aggregation includes rounding, finishing and matrices (including revenue matrices)

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Shortcuts ':all';

sub aggregation {

    my ( $model, $componentMap, $allTariffsByEndUser, $chargingLevels,
        $nonExcludedComponents, $allComponents, $sourceMap, )
      = @_;

    my %tariffComponentSources;

    my $aggregationTable = sub {
        my ( $component, $sourceName, @sourceTables ) = @_;
        my @relevantGroups = grep {
                 $componentMap->{$_}{$component}
              && $componentMap->{$_}{$component} eq $sourceName;
        } @{ $allTariffsByEndUser->{groups} || $allTariffsByEndUser->{list} };
        return unless @relevantGroups;
        my $relevantTariffs = Labelset(
            name   => "Tariffs with $component from $sourceName",
            groups => \@relevantGroups
        );
        push @{ $tariffComponentSources{$component} }, map {
            View(
                name =>
                  Label( $_->{name}, "$_->{name} — for $relevantTariffs" ),
                rows => $relevantTariffs,
                cols => ref $_ eq 'SpreadsheetModel::GroupBy'
                ? $_->{source}{cols}
                : $_->{cols},
                sources =>
                  [ ref $_ eq 'SpreadsheetModel::GroupBy' ? $_->{source} : $_ ]
            );
        } grep { $_ } @sourceTables;
    };

    foreach my $c (@$nonExcludedComponents) {
        foreach my $s ( sort keys %{ $sourceMap->{$c} } ) {
            $aggregationTable->( $c, $s, @{ $sourceMap->{$c}{$s} } );
        }
    }

    sub _sortRule {
        local $_ = SpreadsheetModel::Object::_shortName( $_[0]{name} );
        ( /cust|serv/i  ? 2 : 1 )
          . ( /stand/i  ? 1 : 2 )
          . ( /yard/i   ? 2 : 1 )
          . ( /operat/i ? 2 : 1 )
          . $_;
    }

    my $tariffsExMatching = {
        map {
            my $tariffComponent = $_;
            my $stack           = Stack(
                name    => "$tariffComponent (elements)",
                rows    => $allTariffsByEndUser,
                cols    => $chargingLevels,
                sources => [
                    sort { _sortRule($a) cmp _sortRule($b) }
                      @{ $tariffComponentSources{$tariffComponent} }
                ],
            );
            push @{ $model->{preliminaryAggregation} }, $stack
              if @{ $tariffComponentSources{$tariffComponent} };
            my $agg = GroupBy(
                name       => "$tariffComponent (total)",
                rows       => $allTariffsByEndUser,
                cols       => 0,
                source     => $stack,
                rowFormats => [
                    map {
                        $componentMap->{$_}{$tariffComponent}
                          ? undef
                          : 'unavailable';
                    } @{ $allTariffsByEndUser->{list} }
                ]
            );
            0 and Columnset(
                name    => $tariffComponent . ' aggregation',
                columns => [ $stack, $agg ]
            );
            $_ => $agg;
        } @$nonExcludedComponents
    };

    $tariffsExMatching;

}

sub roundingAndFinishing {

    my (
        $model,                 $allComponents,
        $tariffsExMatching,     $componentLabelset,
        $allTariffs,            $componentMap,
        $sourceMap,             $daysInYear,
        $nonExcludedComponents, $unitsLossAdjustment,
        $volumeData,            $allTariffsByEndUser,
        $totalRevenuesSoFar,    $totalRevenuesFromMatching,
        $allowedRevenue,        $revenueBefore,
        $revenueFromElsewhere,  $siteSpecificCharges,
        $chargingLevels,        @matchingTables,
    ) = @_;

    my %adjTable;

    if ($unitsLossAdjustment) {

        my $lossAdjCol = Labelset( list => ['Losses adjustment'] );

        foreach my $tariffComponent ( grep { /Unit rate/i } @$allComponents ) {

            my @tables = $tariffsExMatching->{$tariffComponent};

            push @tables, $_ foreach grep { $_ }
              map { $_->{$tariffComponent} } @matchingTables;

            my $cols = $componentLabelset->{$tariffComponent} || $lossAdjCol;

            $adjTable{$tariffComponent} = Arithmetic(
                name =>
                  Label( $tariffComponent, "$tariffComponent loss adjustment" ),
                rows       => $allTariffsByEndUser,
                cols       => $cols,
                arithmetic => '=IV9*('
                  . join( '+', map { "IV$_" } 1 .. @tables ) . ')',
                arguments => {
                    IV9 => $unitsLossAdjustment,
                    map { 'IV' . ( 1 + $_ ) => $tables[$_] } 0 .. $#tables
                },
                defaultFormat => '0.000softnz',
            );

        }

        foreach my $comp ( keys %adjTable ) {
            foreach my $src ( keys %{ $sourceMap->{$comp} } ) {
                push @{ $sourceMap->{$comp}{$src} }, $adjTable{$comp};
            }
        }

        Columnset(
            name => 'Adjust unit rates for losses'
              . ' (for embedded network tariffs)',
            columns => [ @adjTable{ grep { /Unit rate/i } @$allComponents } ]
        );

    }

    my $theRoundingCol = Labelset( list => ['Rounding'] );

    my ( %tariffsBeforeRounding, %roundingRule );

    foreach my $tariffComponent (@$allComponents) {

        my @tables = $tariffsExMatching->{$tariffComponent};

        push @tables, $_
          foreach grep { $_ } map { $_->{$tariffComponent} } @matchingTables,
          \%adjTable;

        my $cols = $componentLabelset->{$tariffComponent} || $theRoundingCol;

        $tariffsBeforeRounding{$tariffComponent} = Arithmetic(
            name =>
              Label( $tariffComponent, "$tariffComponent before rounding" ),
            rows       => $allTariffsByEndUser,
            cols       => $cols,
            arithmetic => '=' . join( '+', map { "IV$_" } 1 .. @tables ),
            arguments =>
              { map { 'IV' . ( 1 + $_ ) => $tables[$_] } 0 .. $#tables },
            defaultFormat => '0.00000soft',
            rowFormats    => [
                map {
                    $componentMap->{$_}{$tariffComponent}
                      ? undef
                      : 'unavailable';
                } @{ $allTariffsByEndUser->{list} }
            ]
        );

        my $digits =
          $tariffComponent =~ m%p/k(W|VAr)h%
          ? 3
          : 2;

        $roundingRule{$tariffComponent} = Constant(
            name =>
              Label( $tariffComponent, "$tariffComponent decimal places" ),
            cols          => $cols,
            data          => [ map { $digits } @{ $cols->{list} } ],
            defaultFormat => '0connz',
            1 ? () : (
                rowFormats => [
                    map {
                        $componentMap->{$_}{$tariffComponent}
                          ? undef
                          : 'unavailable';
                    } @{ $allTariffsByEndUser->{list} }
                ]
            )
        );

    }

    my %roundingTable = map {
        my $tariffComponent = $_;
        $tariffComponent => Arithmetic(
            name => Label( $_, "$tariffComponent rounding" ),
            rows => $allTariffsByEndUser,
            cols       => $roundingRule{$tariffComponent}{cols},
            arithmetic => $model->{monthly} && $tariffComponent =~ /day/i
            ? '=IF(IV99="",ROUND(IV1,IV2)-IV3,ROUND(IV91*IV98/12,0)*12/IV97-IV93)'
            : $model->{capping} ? '=IF(IV81<0,0,ROUND(IV1,IV2))-IV3'
            : '=ROUND(IV1,IV2)-IV3',
            arguments => {
                IV1  => $tariffsBeforeRounding{$tariffComponent},
                IV81 => $tariffsBeforeRounding{$tariffComponent},
                IV2  => $roundingRule{$tariffComponent},
                IV3  => $tariffsBeforeRounding{$tariffComponent},
                $model->{monthly}
                ? (
                    IV91 => $tariffsBeforeRounding{$tariffComponent},
                    IV93 => $tariffsBeforeRounding{$tariffComponent},
                    IV97 => $daysInYear,
                    IV98 => $daysInYear,
                    IV99 => $tariffsBeforeRounding{
                        $componentLabelset->{'Reactive power charge p/kVArh'}
                        ? 'Capacity charge p/kVA/day'
                        : 'Reactive power charge p/kVArh'
                    },
                  )
                : ()
            },
            defaultFormat => '0.00000soft',
            rowFormats    => [
                map {
                    $componentMap->{$_}{$tariffComponent}
                      ? undef
                      : 'unavailable';
                } @{ $allTariffsByEndUser->{list} }
            ]
        );
    } @$allComponents;

    foreach my $comp ( keys %roundingTable ) {
        foreach my $src ( keys %{ $sourceMap->{$comp} } ) {
            push @{ $sourceMap->{$comp}{$src} }, $roundingTable{$comp};
        }
    }

    my $revenuesFromRounding;

    {
        my @termsNoDays;
        my @termsWithDays;
        my %args = ( IV400 => $daysInYear );
        my $i = 1;
        foreach ( grep { $roundingTable{$_} } @$nonExcludedComponents ) {
            ++$i;
            my $pad = "$i";
            $pad = "0$pad" while length $pad < 3;
            if (m#/day#) {
                push @termsWithDays, "IV2$pad*IV3$pad";
            }
            else {
                push @termsNoDays, "IV2$pad*IV3$pad";
            }
            $args{"IV2$pad"} = $roundingTable{$_};
            $args{"IV3$pad"} = $volumeData->{$_};
        }
        $revenuesFromRounding = Arithmetic(
            name       => 'Net revenues by tariff from rounding',
            rows       => $allTariffsByEndUser,
            arithmetic => '='
              . join( '+',
                @termsWithDays
                ? ( '0.01*IV400*(' . join( '+', @termsWithDays ) . ')' )
                : (),
                @termsNoDays ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
                : (),
              ),
            arguments     => \%args,
            defaultFormat => '0soft'
          )
    };

    my $totalRevenuesFromRounding = GroupBy(
        name          => 'Total net revenues from rounding (£)',
        rows          => 0,
        cols          => 0,
        source        => $revenuesFromRounding,
        defaultFormat => '0soft'
    );

    push @{ $model->{roundingResults} },
      Columnset(
        name    => 'Tariffs before rounding',
        columns => [ grep { $_ } @tariffsBeforeRounding{@$allComponents} ]
      ),
      Columnset(
        name    => 'Decimal places',
        columns => [ grep { $_ } @roundingRule{@$allComponents} ]
      ),
      Columnset(
        name    => 'Tariff rounding',
        columns => [ grep { $_ } @roundingTable{@$allComponents} ]
      );

    {
        my @columns = map {
                 !$_->{location}
              && !$_->{mustCopy}
              && !$model->{copy} ? $_ : Arithmetic(
                defaultFormat => '0soft',
                arithmetic    => '=IV1',
                name => ref $_->{name} ? $_->{name}->shortName : $_->{name},
                arguments => { IV1 => $_ }
              )
          } (
            $totalRevenuesSoFar,
            $revenueFromElsewhere      ? $revenueFromElsewhere      : (),
            $totalRevenuesFromMatching ? $totalRevenuesFromMatching : (),
          );

        splice @columns, 1, 0,
          GroupBy(
            name          => 'Total site-specific sole use asset charges (£)',
            defaultFormat => '0softnz',
            source        => $siteSpecificCharges
          ) if $siteSpecificCharges;
        push @columns, $totalRevenuesFromRounding;
        my $totalNet = Arithmetic(
            name          => 'Total net revenues (£)',
            defaultFormat => '0soft',
            arithmetic    => '=' . join( '+', map { "IV1$_" } 0 .. $#columns ),
            arguments => { map { ; "IV1$_" => $columns[$_]; } 0 .. $#columns }
        );
        my $revenueError = Arithmetic(
            name          => 'Deviation from target revenue (£)',
            defaultFormat => '0soft',
            arithmetic    => '=IV1-IV2' . ( $revenueBefore ? '+IV3' : '' ),
            arguments     => {
                IV1 => $totalNet,
                IV2 => $allowedRevenue,
                $revenueBefore ? ( IV3 => $revenueBefore ) : (),
            }
        );
        push @{ $model->{revenueSummaryTables} },
          Columnset(
            name    => 'Revenue forecast summary',
            columns => [ @columns, $totalNet, $revenueError, ],
          );
        my $rerr = Stack( sources => [$revenueError] );
        splice @{ $model->{summaryColumns} }, 1, 0,
          $totalRevenuesFromMatching
          ? Stack( sources => [$totalRevenuesFromMatching] )
          : Constant(
            name => 'No revenue matching',
            data => [''],
          ),
          $rerr,
          Arithmetic(
            name          => 'Over/under recovery',
            defaultFormat => '%soft',
            arithmetic    => '=IV1/IV2',
            arguments     => { IV1 => $rerr, IV2 => $allowedRevenue }
          );
    }

    my %tariffTable =
      map {
        my $tariffComponent = $_;
        my @tables;
        if ( $model->{repeatCalculation} ) {
            @tables = $tariffsExMatching->{$tariffComponent};
            push @tables, $_ foreach grep { $_ }
              map { $_->{$tariffComponent} } @matchingTables, \%adjTable,
              \%roundingTable;
        }
        else {
            @tables = (
                $tariffsBeforeRounding{$tariffComponent},
                $roundingTable{$tariffComponent}
            );
        }
        $tariffComponent => Arithmetic(
            name       => $tariffComponent,
            rows       => $allTariffs,
            cols       => $componentLabelset->{$_},
            arithmetic => '=' . join( '+', map { "IV$_" } 1 .. @tables ),
            arguments =>
              { map { 'IV' . ( 1 + $_ ) => $tables[$_] } 0 .. $#tables },
            $tariffComponent =~ m%p/k(W|VAr)h% ? ()
            : ( defaultFormat => '0.00soft' ),
            rowFormats => [
                map {
                    $componentMap->{$_}{$tariffComponent} ? undef
                      : 'unavailable';
                } @{ $allTariffs->{list} }
            ]
        );
      } @$allComponents;

    $model->{allTariffColumns} = [ @tariffTable{@$allComponents} ];

    push @{ $model->{tariffSummary} }, Columnset(
        name => ( $model->{pcd} ? 'All the way tariffs' : 'Tariffs' )
          . ( $model->{monthly} ? ' with daily charges' : '' ),
        $model->{pcd} || $model->{noLLFCs}
        ? ()
        : (
            dataset               => $model->{dataset},
            doNotCopyInputColumns => 1,
            number                => 3701,
        ),    # hacks to get the LLFCs to copy
        columns => $model->{pcd}
          || $model->{noLLFCs} ? $model->{allTariffColumns} : [
            Dataset(
                rows          => $allTariffs,
                defaultFormat => 'texthard',
                data          => [ map { '' } @{ $allTariffs->{list} } ],
                name          => 'Open LLFCs',
            ),
            Constant(
                rows          => $allTariffs,
                defaultFormat => 'textcon',
                data          => [
                    map {
                        my ($pc) = map { /^PC(.*)/ ? $1 : () }
                          keys %{ $componentMap->{$_} };
                        $pc || '';
                    } @{ $allTariffs->{list} }
                ],
                name => 'PCs',
            ),
            @{ $model->{allTariffColumns} },
            Dataset(
                rows          => $allTariffs,
                defaultFormat => 'texthard',
                data          => [ map { '' } @{ $allTariffs->{list} } ],
                name          => 'Closed LLFCs',
            ),
          ]
    );

    push @{ $model->{tariffSummary} }, Columnset(
        name    => 'Tariffs with monthly charges',
        columns => (
            $model->{allTariffColumns} = [
                map {
                    !/day/i ? Stack( sources => [ $tariffTable{$_} ] )
                      : /kVA/i ? Arithmetic(
                        name          => 'Capacity charge p/kVA/month',
                        defaultFormat => '0softnz',
                        arithmetic    => '=IV1*IV2/12',
                        arguments =>
                          { IV1 => $tariffTable{$_}, IV2 => $daysInYear }
                      )
                      : (
                        Arithmetic(
                            name          => 'Fixed charge p/MPAN/day',
                            defaultFormat => '0.00soft',
                            arithmetic    => '=IF(IV3="",IV1,"")',
                            arguments     => {
                                IV1 => $tariffTable{$_},
                                IV2 => $daysInYear,
                                IV3 => $tariffTable{
                                    $componentLabelset->{
                                        'Reactive power charge p/kVArh'}
                                    ? 'Capacity charge p/kVA/day'
                                    : 'Reactive power charge p/kVArh'
                                }
                            }
                        ),
                        Arithmetic(
                            name          => 'Fixed charge p/MPAN/month',
                            defaultFormat => '0soft',
                            arithmetic    => '=IF(IV3="","",IV1*IV2/12)',
                            arguments     => {
                                IV1 => $tariffTable{$_},
                                IV2 => $daysInYear,
                                IV3 => $tariffTable{
                                    $componentLabelset->{
                                        'Reactive power charge p/kVArh'}
                                    ? 'Capacity charge p/kVA/day'
                                    : 'Reactive power charge p/kVArh'
                                }
                            }
                        )
                      )
                } @$allComponents
            ]
        )
    ) if $model->{monthly};

    \%tariffTable, \%tariffsBeforeRounding, \%adjTable;

}

sub makeMatrixClosure {

    my (
        $model,             $adjTableRef,    $allComponents,
        $componentLabelset, $allTariffs,     $componentMap,
        $sourceMap,         $daysInYear,     $volumeData,
        $unitsInYear,       $chargingLevels, @matchingTables,
    ) = @_;

    $model->{niceTariffMatrices} = sub {

        my ($filteringClosure) = @_;

        map {

            /([^\n]*)$/s;
            my $tariffShort = $1 || $_;
            my $tariff      = $_;
            my $theCol      = Labelset( list => [$tariff] );
            my @components =
              grep { $componentMap->{$tariff}{$_} && !$componentLabelset->{$_} }
              @$allComponents;

            my %source =
              map {
                $_ => [
                    map {
                        ref $_ eq 'SpreadsheetModel::GroupBy'
                          ? $_->{source}
                          : $_
                      }
                      grep { $_ }
                      @{ $sourceMap->{$_}{ $componentMap->{$tariff}{$_} } }
                  ]
              } @components;

            my $revenuePots = Labelset(
                name => 'Revenue pots',
                list => [
                    grep {
                        my $pot = $_;
                        !/^GSP/i && grep {
                            grep { $pot eq $_ }
                              @{ $_->{cols}{list} }
                          } map { @{ $source{$_} } } @components
                      } (
                        @{ $chargingLevels->{list} },
                        (
                            map {
                                my ($sampleDataset) = grep { $_ } values %$_;
                                $sampleDataset
                                  ? @{ $sampleDataset->{cols}{list} }
                                  : ();
                              } @matchingTables,
                            $adjTableRef
                        ),
                        'Rounding'
                      )
                ]
            );

            my %columns;

            my @columns = map {
                $columns{$_} = Stack(
                    name => Label( $_, "$tariffShort $_" ),
                    rows => $revenuePots,
                    cols => $componentLabelset->{$_} || $theCol,
                    sources => $source{$_} || [],
                    defaultFormat => m%month% ? '0copy'
                    : m%p/k(W|VAr)h% ? '0.000copy'
                    :                  '0.00copy'
                );
            } $model->{matrices} =~ /bigger/i ? @$allComponents : @components;

            my @columnsets;

            if ( $model->{matrices} =~ /big/i ) {

                my %volumes;

                my @volumes = map {
                    $volumes{$_} =
                      !$volumeData->{$_} ? Stack(
                        name    => '',
                        rows    => $theCol,
                        sources => [],
                      )
                      : $model->{matrices} =~ /cust/i ? Arithmetic(
                        name => Label(
                            $volumeData->{$_}{name},
                            "$tariffShort $volumeData->{$_}{name}"
                        ),
                        rows          => $theCol,
                        arithmetic    => '=IF(IV1<>0,IV2/IV3,IV4)',
                        defaultFormat => '0softnz',
                        arguments     => {
                            IV1 => $volumeData->{'Fixed charge p/MPAN/day'},
                            IV2 => $volumeData->{$_},
                            IV3 => $volumeData->{'Fixed charge p/MPAN/day'},
                            IV4 => $volumeData->{$_},
                        }
                      )
                      : Stack(
                        name => Label(
                            $volumeData->{$_}{name},
                            "$tariffShort $volumeData->{$_}{name}"
                        ),
                        defaultFormat => '0copynz',
                        rows          => $theCol,
                        sources       => [ $volumeData->{$_} ],
                      );
                  } $model->{matrices} =~ /bigger/i
                  ? @$allComponents
                  : @components;

                my $units =
                  $model->{matrices} =~ /cust/i
                  ? Arithmetic(
                    name => $model->{matrices} =~ /partyear/
                    ? Label( 'MWh',      "$tariffShort MWh" )
                    : Label( 'MWh/year', "$tariffShort MWh/year" ),
                    rows       => $theCol,
                    arithmetic => '=IF(IV1<>0,IV2/IV3,IV4)',
                    arguments  => {
                        IV1 => $volumeData->{'Fixed charge p/MPAN/day'},
                        IV2 => $unitsInYear,
                        IV3 => $volumeData->{'Fixed charge p/MPAN/day'},
                        IV4 => $unitsInYear,
                    }
                  )
                  : Stack(
                    name => $model->{matrices} =~ /partyear/
                    ? Label( 'MWh',      "$tariffShort MWh" )
                    : Label( 'MWh/year', "$tariffShort MWh/year" ),
                    rows    => $theCol,
                    sources => [$unitsInYear],
                  );

                push @volumes, $units;

                push @volumes,
                  Arithmetic(
                    name => $model->{matrices} =~ /partyear/
                    ? Label( 'MWh/MPAN',      "$tariffShort MWh/MPAN" )
                    : Label( 'MWh/MPAN/year', "$tariffShort MWh/MPAN/year" ),
                    arithmetic => '=IF(IV3,IV1/IV2,"")',
                    arguments  => {
                        IV1 => $units,
                        IV2 => $volumes{'Fixed charge p/MPAN/day'},
                        IV3 => $volumes{'Fixed charge p/MPAN/day'}
                    }
                  ) if $units && $volumes{'Fixed charge p/MPAN/day'};

                push @columnsets,
                  Columnset( name => $tariffShort, columns => \@volumes );

                my ( $revenues, $averageUnitRate );

                {
                    my @termsNoDays;
                    my @termsWithDays;
                    my %args = ( IV400 => $daysInYear );
                    my $i = 1;
                    foreach (@components) {
                        ++$i;
                        my $pad = "$i";
                        $pad = "0$pad" while length $pad < 3;
                        if (m#/day#) {
                            push @termsWithDays, "IV2$pad*IV3$pad";
                        }
                        else {
                            push @termsNoDays, "IV2$pad*IV3$pad";
                        }
                        $args{"IV2$pad"} = $columns{$_};
                        $args{"IV3$pad"} = $volumes{$_};
                    }

                    push @{ $model->{revenueTables} },
                      $revenues = Arithmetic(
                        name => $model->{matrices} =~ /partyear/
                        ? 'Revenue (£)'
                        : 'Revenue (£/year)',
                        rows       => $revenuePots,
                        cols       => $theCol,
                        arithmetic => '='
                          . join(
                            '+',
                            @termsWithDays
                            ? ( '0.01*IV400*('
                                  . join( '+', @termsWithDays )
                                  . ')' )
                            : ('0'),
                            @termsNoDays
                            ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
                            : ('0'),
                          ),
                        arguments     => \%args,
                        defaultFormat => $model->{matrices} =~ /cust/i
                        ? '0.00softnz'
                        : '0softnz'
                      );

                    $averageUnitRate = Arithmetic(
                        name       => Label('Average unit rate (p/kWh)'),
                        rows       => $revenuePots,
                        arithmetic => '=IF(IV903<>0,('
                          . ( '(' . join( '+', @termsNoDays ) . ')' )
                          . ')/IV902,0)',
                        arguments => {
                            %args,
                            IV902 => $units,
                            IV903 => $units,
                        },
                    );
                };

                push @columns,
                  $volumes{'Unit rate 2 p/kWh'} ? $averageUnitRate : (),
                  $revenues,
                  Arithmetic(
                    name       => 'Average p/kWh',
                    arithmetic => '=IF(IV3<>0,0.1*IV1/IV2,"")',
                    arguments  => {
                        IV1 => $revenues,
                        IV2 => $units,
                        IV3 => $units,
                    }
                  ),
                  $volumes{'Fixed charge p/MPAN/day'}
                  ? Arithmetic(
                    name => $model->{matrices} =~ /partyear/
                    ? 'Average £/MPAN'
                    : 'Average £/MPAN/year',
                    defaultFormat => '0.00softnz',
                    arithmetic    => '=IF(IV3<>0,IV1/IV2,"")',
                    arguments     => {
                        IV1 => $revenues,
                        IV2 => $volumes{'Fixed charge p/MPAN/day'},
                        IV3 => $volumes{'Fixed charge p/MPAN/day'},
                    }
                  )
                  : (),
                  $volumes{'Capacity charge p/kVA/day'}
                  ? Arithmetic(
                    name          => 'Average p/kVA/day',
                    defaultFormat => '0.00softnz',
                    arithmetic    => '=IF(IV3<>0,IV1/IV2*100/IV4,"")',
                    arguments     => {
                        IV1 => $revenues,
                        IV2 => $volumes{'Capacity charge p/kVA/day'},
                        IV3 => $volumes{'Capacity charge p/kVA/day'},
                        IV4 => $daysInYear,
                    }
                  )
                  : (),;

            }

            my $tariffComponentColumnset = Columnset(
                name => $model->{matrices} =~ /big/i ? '' : $tariffShort,
                columns => [@columns]
            );

            push @columnsets, $tariffComponentColumnset, Columnset(
                rows          => 0,
                name          => '',
                noHeaders     => 1,
                singleRowName => 'Total',
                columns       => [
                    map {
                        GroupBy(
                            name   => $_->{name},
                            rows   => 0,
                            source => $_,
                            $_->{defaultFormat}
                            ? (
                                defaultFormat => $_->{defaultFormat} =~ /000/
                                ? '0.000soft'
                                : $_->{defaultFormat} =~ /00/ ? '0.00soft'
                                :                               '0soft'
                              )
                            : ()
                          )
                    } @columns
                ],
            );

            @columnsets;

          }
          grep {
                 $componentMap->{$_}{'Unit rate 1 p/kWh'}
              && $filteringClosure->($_)
          }
          map { $allTariffs->{list}[$_] } $allTariffs->indices;

    };

}

sub revenueMatrices {
    my ($model) = @_;
    my @levels;
    my %levels;
    foreach my $l ( map { @{ $_->{rows}{list} } } @{ $model->{revenueTables} } )
    {
        unless ( exists $levels{$l} ) {
            push @levels, $l;
            undef $levels{$l};
        }
    }
    my $levels = Labelset( list => \@levels );

    my @tariffs;
    my %tariffs;
    foreach my $l ( map { @{ $_->{cols}{list} } } @{ $model->{revenueTables} } )
    {
        unless ( exists $tariffs{$l} ) {
            push @tariffs, $l;
            undef $tariffs{$l};
        }
    }
    my $tariffs = Labelset( list => \@tariffs );

    my $revenues = Stack(
        name => 'Revenue matrix by tariff, '
          . 'charging element and network level',
        rows          => $tariffs,
        cols          => $levels,
        sources       => $model->{revenueTables},
        defaultFormat => '0copynz'
    );

    Columnset(
        name    => 'Revenue matrix by tariff',
        columns => [
            $revenues,
            GroupBy(
                name => 'Total net revenue by tariff (£/'
                  . (
                    $model->{matrices} =~ /partyear/
                    ? 'period)'
                    : 'year)'
                  ),
                source        => $revenues,
                rows          => $tariffs,
                defaultFormat => '0softnz'
              )

        ]
      ),

      Columnset(
        name    => 'Revenues by charging element and network level',
        columns => [
            GroupBy(
                name => 'Total net revenue by charging element'
                  . ' and network level (£/'
                  . (
                    $model->{matrices} =~ /partyear/ ? 'period)'
                    : 'year)'
                  ),
                source        => $revenues,
                cols          => $levels,
                defaultFormat => '0softnz'
            ),
            GroupBy(
                name => 'Total net revenue (£/'
                  . (
                    $model->{matrices} =~ /partyear/ ? 'period)'
                    : 'year)'
                  ),
                source        => $revenues,
                defaultFormat => '0softnz'
            )
        ]
      );

}

1;
