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

sub pcdPreprocessVolumes {

    my (
        $model,                   $allEndUsers,
        $componentMap,            $daysAfter,
        $daysBefore,              $daysInYear,
        $nonExcludedComponents,   $revenueBeforeRef,
        $revenuesBeforeRef,       $tariffsBeforeRef,
        $unitsByEndUserRef,       $unitsInYearRef,
        $unitsInYearAfterRef,     $volumeDataRef,
        $volumeDataAfterRef,      $volumesAdjustedRef,
        $volumesAdjustedAfterRef, $volumesByEndUserRef,
    ) = @_;

    my @combinations;
    my @data;

    foreach ( @{ $model->{pcd}{allTariffsByEndUser}{list} } ) {
        my $combi =
            /gener/i                              ? 'No discount'
          : /^(LDNO (?:.*?): (?:\S+V(?: Sub)?))/i ? "$1 user"
          :                                         'No discount';
        push @combinations, $combi
          unless grep { $_ eq $combi } @combinations;
        push @data,
          [
            ( map { $_ eq $combi ? 1 : 0 } @combinations ),
            ( map { 0 } 1 .. 20 )
          ];
    }

    my $combinations =
      Labelset( name => 'Discount combinations', list => \@combinations );

    my $rawDiscount = Dataset(
        name          => 'Embedded network (LDNO) discounts',
        singleRowName => 'LDNO discount',
        number        => 1037,
        appendTo      => $model->{inputTables},
        dataset       => $model->{dataset},
        validation    => {
            validate      => 'decimal',
            criteria      => 'between',
            minimum       => 0,
            maximum       => 1,
            input_title   => 'LDNO discount:',
            input_message => 'Between 0% and 100%',
            error_message => 'The LDNO discount'
              . ' must be between 0% and 100%.'
        },
        lines => [ 'Source: separate price control disaggregation model.', ],
        defaultFormat => '%hardnz',
        cols          => $combinations,
        data          => [
            map {
                    /^no/i         ? undef
                  : /LDNO LV: LV/i ? 0.3
                  : /LDNO HV: LV/i ? 0.4
                  :                  0;
            } @combinations
        ]
    );

    push @{ $model->{volumeData} },
      $model->{pcd}{discount} = SumProduct(
        name          => 'Discount for each tariff (except for fixed charges)',
        defaultFormat => '%softnz',
        matrix        => Constant(
            name          => 'Discount map',
            defaultFormat => '0connz',
            rows          => $model->{pcd}{allTariffsByEndUser},
            cols          => $combinations,
            byrow         => 1,
            data          => \@data
        ),
        vector => $rawDiscount
      );

    my $ldnoGenerators = Labelset(
        name => 'Generators on LDNO networks',
        list => [
            grep { /ldno/i && /gener/i }
              @{ $model->{pcd}{allTariffsByEndUser}{list} }
        ]
    );

    $model->{pcd}{discountFixed} = Stack(
        name          => 'Discount for each tariff for fixed charges only',
        defaultFormat => '%copynz',
        rows          => $model->{pcd}{discount}{rows},
        cols          => $model->{pcd}{discount}{cols},
        sources       => [
            Constant(
                name => '100 per cent discount for generators on LDNO networks',
                defaultFormat => '%connz',
                rows          => $ldnoGenerators,
                data          => [ [ map { 1 } @{ $ldnoGenerators->{list} } ] ]
            ),
            $model->{pcd}{discount}
        ]
    );

    if ( $model->{portfolio} && $model->{portfolio} =~ /ehv/i ) {

        # Supplement table 1037 with table 1181 or
        # replace everything with a totally new table 1038
        # A new table 1038 would show, for each level pair,
        # separate discounts for demand unit changes, demand standing charges,
        # generation credits and generation fixed charges.

        die 'EDCM discounted LDNO tariffs are not implemented here (yet)';

    }

    ( $model->{pcd}{volumeData} ) =
      $model->volumes( $model->{pcd}{allTariffsByEndUser},
        $allEndUsers, $nonExcludedComponents, $componentMap, 'no aggregation' );

    if ( $model->{inYear} ) {
        if ( $model->{inYear} =~ /after/i ) {
            $model->inYearAdjustUsingAfter(
                $nonExcludedComponents, $model->{pcd}{volumeData},
                $allEndUsers,           $componentMap,
                $revenueBeforeRef,      $unitsInYearAfterRef,
                $volumeDataAfterRef,    $volumesAdjustedAfterRef,
            );
        }
        else {
            $model->inYearAdjustUsingBefore(
                $nonExcludedComponents, $model->{pcd}{volumeData},
                $allEndUsers,           $componentMap,
                $daysAfter,             $daysBefore,
                $daysInYear,            $revenueBeforeRef,
                $revenuesBeforeRef,     $tariffsBeforeRef,
                $unitsInYearAfterRef,   $volumeDataAfterRef,
                $volumesAdjustedAfterRef,
            );
        }
    }
    elsif ( $model->{addVolumes} && $model->{addVolumes} =~ /matching/i ) {
        $model->inYearAdjustUsingAfter(
            $nonExcludedComponents, $model->{pcd}{volumeData},
            $allEndUsers,           $componentMap,
            undef,                  $unitsInYearAfterRef,
            $volumeDataAfterRef,    $volumesAdjustedAfterRef,
        );
    }

    my %intermediate = map {
        $_ => Arithmetic(
            name => SpreadsheetModel::Object::_shortName(
                $model->{pcd}{volumeData}{$_}{name}
            ),
            arithmetic => '=IV1*(1-IV2)',
            arguments  => {
                IV1 => $model->{pcd}{volumeData}{$_},
                IV2 => /fix/i
                ? $model->{pcd}{discountFixed}
                : $model->{pcd}{discount}
            }
        );
    } @$nonExcludedComponents;

    Columnset(
        name    => 'LDNO discounts and volumes adjusted for discount',
        columns => [
            $model->{pcd}{discount},
            $model->{pcd}{discountFixed},
            @intermediate{@$nonExcludedComponents}
        ]
    );

    $$volumesAdjustedRef = $$volumesByEndUserRef = $$volumeDataRef = {
        map {
            $_ => GroupBy(
                name => SpreadsheetModel::Object::_shortName(
                    $intermediate{$_}{name}
                ),
                rows          => $allEndUsers,
                source        => $intermediate{$_},
                defaultFormat => '0soft',
            );
        } @$nonExcludedComponents
    };

    push @{ $model->{volumeData} },
      Columnset(
        name    => 'Equivalent volume for each end user',
        columns => [ @{$$volumesAdjustedRef}{@$nonExcludedComponents} ]
      );

    $$unitsInYearRef = $$unitsByEndUserRef = Arithmetic(
        noCopy     => 1,
        name       => 'All units (MWh)',
        arithmetic => '='
          . join( '+', map { "IV$_" } 1 .. $model->{maxUnitRates} ),
        arguments => {
            map {
                ( "IV$_" => ${$volumesByEndUserRef}->{"Unit rate $_ p/kWh"} )
            } 1 .. $model->{maxUnitRates}
        },
        defaultFormat => '0softnz'
    );

}

sub pcdApplyDiscounts {

    my (
        $model,     $allComponents, $allTariffs, $componentMap,
        $daysAfter, $tariffTable,   $volumeData, $volumeDataAfter,
    ) = @_;

    my $unitsInYear = Arithmetic(
        noCopy     => 1,
        name       => 'All units (MWh)',
        arithmetic => '='
          . join( '+', map { "IV$_" } 1 .. $model->{maxUnitRates} ),
        arguments => {
            map { ( "IV$_" => $volumeData->{"Unit rate $_ p/kWh"} ) }
              1 .. $model->{maxUnitRates}
        },
        defaultFormat => '0softnz'
    );

    my $electionBung;
    if ( $model->{electionBung} ) {
        $electionBung = Dataset(
            name     => 'Election bung (p/MPAN/day)',
            rows     => $allTariffs,
            number   => 1098,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            , data => [ map { '' } @{ $allTariffs->{list} } ]
        );
        push @{ $model->{summaryColumns} },
          my $totalImpactOfElectionBung = Arithmetic(
            name => 'Revenue impact of election bung (£, not accounted for)',
            defaultFormat => '0soft',
            arithmetic    => '=0.01*IV1*SUMPRODUCT(IV2_IV3,IV4_IV5)',
            arguments     => {
                IV1     => $daysAfter,
                IV2_IV3 => $electionBung,
                IV4_IV5 => ( $volumeDataAfter || $volumeData )
                  ->{'Fixed charge p/MPAN/day'}
            }
          );
        $model->{arpSharedData}->addStats( $model, $totalImpactOfElectionBung )
          if $model->{arpSharedData};
    }

    $tariffTable = {
        map {
            $_ => Arithmetic(
                name => SpreadsheetModel::Object::_shortName(
                    $tariffTable->{$_}{name}
                ),
                defaultFormat => (
                    map { local $_ = $_; s/soft/copy/ if $_; $_; }
                      $tariffTable->{$_}{defaultFormat}
                ),
                arithmetic => $model->{model100} ? '=IV2*(1-IV1)'
                : (   ( $electionBung && /MPAN/ ? '=IV3+' : '=' )
                    . 'ROUND('
                      . 'IV2*(1-IV1),'
                      . ( /kWh|kVArh/ ? 3 : 2 )
                      . ')' ),
                rows      => $allTariffs,
                cols      => $tariffTable->{$_}{cols},
                arguments => {
                    IV2 => $tariffTable->{$_},
                    IV1 => /fix/i ? $model->{pcd}{discountFixed}
                    : $model->{pcd}{discount},
                    $electionBung && /MPAN/ ? ( IV3 => $electionBung ) : (),
                },
            );
        } @$allComponents
    };

    my $allTariffsReordered = $allTariffs;
    my @allTariffColumns    = @{$tariffTable}{@$allComponents};
    $model->{allTariffColumns} = \@allTariffColumns;

    unless ( $model->{tariffOrder} ) {

        push @{ $model->{postPcdApplicationResults} },
          Columnset(
            name    => 'Tariffs',
            columns => [ @{$tariffTable}{@$allComponents} ],
          );

        my @allT = map { $allTariffs->{list}[$_] } $allTariffs->indices;

        $allTariffsReordered = Labelset(
            name => 'All tariffs',
            list => [
                ( grep { !/LDNO/i } @allT ),
                ( grep { /LDNO lv/i } @allT ),
                ( grep { /LDNO hv/i && !/LDNO hv sub/i } @allT ),
                ( grep { /LDNO hv sub/i } @allT ),
                ( grep { /LDNO 33/i && !/LDNO 33kV sub/i } @allT ),
                ( grep { /LDNO 33kV sub/i } @allT ),
                ( grep { /LDNO 132/i } @allT ),
                (
                    grep {
                             /LDNO/i
                          && !/LDNO lv/i
                          && !/LDNO hv/i
                          && !/LDNO 33/i
                          && !/LDNO 132/i
                    } @allT
                )
            ]
        );

        @allTariffColumns =
          map {
            Stack(
                name          => $tariffTable->{$_}{name},
                defaultFormat => $tariffTable->{$_}{defaultFormat},
                rows          => $allTariffsReordered,
                cols          => $tariffTable->{$_}{cols},
                sources       => [ $tariffTable->{$_} ]
              )
          } @$allComponents;

    }

    push @{ $model->{roundingResults} }, pop @{ $model->{tariffSummary} };

    unshift @{ $model->{tariffSummary} }, Columnset(
        $model->{noLLFCs}
        ? ( name => '' )
        : (
            name                  => 'Tariffs',
            dataset               => $model->{dataset},
            doNotCopyInputColumns => 1,
            number                => 3701,
        ),    # hacks to get the LLFCs to copy
        columns => [
            $model->{noLLFCs} ? () : (
                Dataset(
                    rows          => $allTariffsReordered,
                    defaultFormat => 'texthard',
                    data          => [ map { '' } @{ $allTariffs->{list} } ],
                    name          => 'Open LLFCs',
                ),
                Constant(
                    rows          => $allTariffsReordered,
                    defaultFormat => 'textcon',
                    data          => [
                        map {
                            my ($pc) = map { /^PC(.*)/ ? $1 : () }
                              keys %{ $componentMap->{$_} };
                            $pc || '';
                        } @{ $allTariffsReordered->{list} }
                    ],
                    name => 'PCs',
                )
            ),
            @allTariffColumns,
            $model->{noLLFCs} ? () : Dataset(
                rows          => $allTariffsReordered,
                defaultFormat => 'texthard',
                data          => [ map { '' } @{ $allTariffs->{list} } ],
                name          => 'Closed LLFCs',
            ),
            $model->{checksums}
            ? (
                map {
                    SpreadsheetModel::Checksum->new(
                        name => $_,
                        /recursive|model/i ? ( recursive => 1 ) : (),
                        digits => /([0-9])/ ? $1 : 6,
                        columns => \@allTariffColumns,
                        factors => [
                            map {
                                     $_->{defaultFormat}
                                  && $_->{defaultFormat} !~ /000/
                                  ? 100
                                  : 1000;
                            } @allTariffColumns
                        ]
                    );
                  } split /;\s*/,
                $model->{checksums}
              )
            : (),
        ]
    );

    $tariffTable, $unitsInYear;

}

1;
