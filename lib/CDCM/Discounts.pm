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

sub pcdSetUp {
    my ( $model, $allEndUsers, $allTariffsByEndUser, $allTariffs ) = @_;
    delete $model->{portfolio};
    delete $model->{boundary};
    $model->{pcd} = {
        allTariffsByEndUser => $allTariffsByEndUser,
        allTariffs          => $allTariffs
    };
    map { $_->{name} =~ s/> //g; } @{ $allEndUsers->{list} };
    $allEndUsers, $allEndUsers, $allEndUsers;
}

sub pcdPreprocessedVolumes {

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
            /gener/i                                    ? 'No discount'
          : /^((?:LD|Q)NO Any): .*(?:ums|unmeter)/i     ? "$1: Unmetered"
          : /^((?:LD|Q)NO (?:.*?): (?:\S+V(?: Sub)?))/i ? "$1 user"
          :                                               'No discount';
        $combi =~ s/\bsub/Sub/;
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

    $model->{embeddedModelM} = ModelM->new(
        dataset => $model->{dataset},
        objects => {
            inputTables => $model->{noSingleInputSheet}
            ? []
            : $model->{inputTables},
            table1037sources => [],
            table1039sources => [],
        },
        %{ $model->{embeddedModelM} },
    ) if $model->{embeddedModelM};

    my $rawDiscount =
      $model->{pcdByTariff} ? Dataset(
        name          => "Embedded network ($model->{ldnoWord}) discounts",
        defaultFormat => '%hard',
        number        => 1038,
        appendTo      => $model->{inputTables},
        dataset       => $model->{dataset},
        validation    => {
            validate      => 'decimal',
            criteria      => '>=',
            value         => 0,
            input_title   => "$model->{ldnoWord} discount:",
            input_message => 'At least 0%',
            error_title   => "Invalid $model->{ldnoWord} discount",
            error_message => "Invalid $model->{ldnoWord} discount"
              . ' (negative number or unused cell).'
        },
        rows => $model->{pcd}{allTariffsByEndUser},
        data => [ map { 0 } @{ $model->{pcd}{allTariffsByEndUser}{list} } ],
      )
      : $model->{embeddedModelM} ? Stack(
        name          => "Embedded network ($model->{ldnoWord}) discounts",
        singleRowName => "$model->{ldnoWord} discount",
        defaultFormat => '%copy',
        cols          => $combinations,
        sources       => $model->{embeddedModelM}{objects}{table1037sources},
      )
      : Dataset(
        name          => "Embedded network ($model->{ldnoWord}) discounts",
        singleRowName => "$model->{ldnoWord} discount",
        number        => 1037,
        appendTo      => $model->{inputTables},
        dataset       => $model->{dataset},
        validation    => {
            validate      => 'decimal',
            criteria      => 'between',
            minimum       => 0,
            maximum       => 1,
            input_title   => "$model->{ldnoWord} discount:",
            input_message => 'Between 0% and 100%',
            error_title   => "Invalid $model->{ldnoWord} discount",
            error_message => "Invalid $model->{ldnoWord} discount"
              . ' (negative number or unused cell).'
        },
        lines => [ 'Source: separate price control disaggregation model.', ],
        defaultFormat => '%hardnz',
        cols          => $combinations,
        data          => [
            map {
                    /^no/i               ? undef
                  : /(?:LD|Q)NO LV: LV/i ? 0.3
                  : /(?:LD|Q)NO HV: LV/i ? 0.5
                  :                        0.4;
            } @combinations
        ]
      );

    push @{ $model->{volumeData} },
      $model->{pcd}{discount} =
      $rawDiscount->{rows} ? $rawDiscount : SumProduct(
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
        name => "Generators on $model->{ldnoWord} networks",
        list => [
            grep { /(?:LD|Q)NO/i && /gener/i }
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
                name =>
"100 per cent discount for generators on $model->{ldnoWord} networks",
                defaultFormat => '%connz',
                rows          => $ldnoGenerators,
                data          => [ [ map { 1 } @{ $ldnoGenerators->{list} } ] ]
            ),
            $model->{pcd}{discount}
        ]
    );

    if ( $model->{portfolio} && $model->{portfolio} =~ /ehv/i ) {

        # To do: supplement table 1037 with table 1181.
        die
          "EDCM discounted $model->{ldnoWord} tariffs are not implemented here";
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
            defaultFormat => '0soft',
            arithmetic    => '=A1*(1-A2)',
            arguments     => {
                A1 => $model->{pcd}{volumeData}{$_},
                A2 => /fix/i
                ? $model->{pcd}{discountFixed}
                : $model->{pcd}{discount}
            }
        );
    } @$nonExcludedComponents;

    Columnset(
        name =>
          "$model->{ldnoWord} discounts and volumes adjusted for discount",
        columns => [
            ref $model->{pcd}{discount} eq 'SpreadsheetModel::Dataset'
            ? ()
            : $model->{pcd}{discount},
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
          . join( '+', map { "A$_" } 1 .. $model->{maxUnitRates} ),
        arguments => {
            map { ( "A$_" => ${$volumesByEndUserRef}->{"Unit rate $_ p/kWh"} ) }
              1 .. $model->{maxUnitRates}
        },
        defaultFormat => '0softnz'
    );

}

sub pcdApplyDiscounts {

    my ( $model, $allComponents, $tariffTable, $daysInYear, ) = @_;

    my $allTariffs = $model->{pcd}{allTariffsByEndUser};
    my $volumeData = $model->{pcd}{volumeData};

    push @{ $model->{roundingResults} },
      Columnset(
        name    => 'All the way tariffs',
        columns => $model->{allTariffColumns},
      );

    push @{ $model->{edcmTables} },
      Columnset(
        name          => 'EDCM input data ⇒1182. CDCM end user tariffs',
        singleRowName => 'CDCM end user tariffs',
        columns =>
          [ map { Stack( sources => [$_] ); } @{ $model->{allTariffColumns} } ]
      ) if $model->{edcmTables};

    my $unitsInYear = Arithmetic(
        noCopy     => 1,
        name       => 'All units (MWh)',
        arithmetic => '='
          . join( '+', map { "A$_" } 1 .. $model->{maxUnitRates} ),
        arguments => {
            map { ( "A$_" => $volumeData->{"Unit rate $_ p/kWh"} ) }
              1 .. $model->{maxUnitRates}
        },
        defaultFormat => '0softnz'
    );

    my $electionBung;
    if ( $model->{electionBung} ) {
        $electionBung = Dataset(
            name               => 'Election bung (p/MPAN/day)',
            rows               => $allTariffs,
            number             => 1098,
            appendTo           => $model->{inputTables},
            dataset            => $model->{dataset},
            data               => [ map { '' } @{ $allTariffs->{list} } ],
            usePlaceholderData => 1,
            validation =>
              {    # some validation is needed to enable lenient locking
                validate => 'decimal',
                criteria => 'between',
                minimum  => -999_999.999,
                maximum  => 999_999.999,
              },
        );
        push @{ $model->{otherTotalRevenues} },
          my $totalImpactOfElectionBung = Arithmetic(
            name => 'Revenue impact of election bung (£, not accounted for)',
            defaultFormat => '0soft',
            arithmetic    => '=0.01*A1*SUMPRODUCT(A2_A3,A4_A5)',
            arguments     => {
                A1    => $daysInYear,
                A2_A3 => $electionBung,
                A4_A5 => $volumeData->{'Fixed charge p/MPAN/day'}
            }
          );
        $model->{sharedData}->addStats( 'DNO-wide aggregates',
            $model, $totalImpactOfElectionBung )
          if $model->{sharedData};
    }

    my $newTariffTable = {
        map {
            $_ => Arithmetic(
                name          => $_,
                defaultFormat => $tariffTable->{$_}{defaultFormat},
                arithmetic    => $model->{model100} ? '=A2*(1-A1)'
                : (   ( $electionBung && /MPAN/ ? '=A3+' : '=' )
                    . 'ROUND('
                      . 'A2*(1-A1),'
                      . ( /kWh|kVArh/ ? 3 : 2 )
                      . ')' ),
                rows      => $allTariffs,
                cols      => $tariffTable->{$_}{cols},
                arguments => {
                    A2 => $tariffTable->{$_},
                    A1 => /fix/i ? $model->{pcd}{discountFixed}
                    : $model->{pcd}{discount},
                    $electionBung && /MPAN/ ? ( A3 => $electionBung ) : (),
                },
            );
        } @$allComponents
    };

    $allTariffs, $allTariffs, $volumeData, $unitsInYear, $newTariffTable;

}

1;
