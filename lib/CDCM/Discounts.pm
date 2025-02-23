package CDCM;

# Copyright 2009-2011 Energy Networks Association Limited and others.
# Copyright 2011-2025 Franck Latrémolière and others.
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

sub pcdSetUp {
    my ( $model, $allEndUsers, $allTariffsByEndUser, $allTariffs ) = @_;
    $model->{pcd} = {
        allTariffsByEndUser => $allTariffsByEndUser,
        allTariffs          => $allTariffs,
        portfolio           => delete $model->{portfolio},
        boundary            => delete $model->{boundary},
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
    my @discountMapData;

    foreach ( @{ $model->{pcd}{allTariffsByEndUser}{list} } ) {
        my $combi =
            /gener/i                                     ? 'No discount'
          : /^((?:ID|LD|Q)NO Any): .*(?:ums|unmeter)/i   ? "$1: Unmetered"
          : /^((?:ID|LD|Q)NO [HL]V: (?:\S+V(?: Sub)?))/i ? "$1 user"
          :                                                'No discount';
        $combi =~ s/\bsub/Sub/;
        push @combinations, $combi
          unless grep { $_ eq $combi } @combinations;
        push @discountMapData,
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
            inputTables => $model->{embeddedModelM}{noSingleInputSheet}
            ? []
            : $model->{inputTables},
        },
        %{ $model->{embeddedModelM} },
    ) if $model->{embeddedModelM};

    $model->{embeddedModelG} = CDCM->new(
        dataset     => $model->{dataset},
        inputTables => $model->{embeddedModelG}{noSingleInputSheet}
        ? []
        : $model->{inputTables},
        table1038sources => [],
        %{ $model->{embeddedModelG} },
    ) if $model->{embeddedModelG};

    my $rawDiscount =
      $model->{pcdByTariff}
      ? (
        $model->{embeddedModelG}
        ? Stack(
            name =>
              "Embedded network ($model->{ldnoWord}) discount percentages",
            defaultFormat => '%copy',
            rows          => $model->{pcd}{allTariffsByEndUser},
            sources       => $model->{embeddedModelG}{table1038sources},
          )
        : Dataset(
            name =>
              "Embedded network ($model->{ldnoWord}) discount percentages",
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
      )
      : $model->{embeddedModelM} ? Stack(
        name => "Embedded network ($model->{ldnoWord}) discount percentages",
        singleRowName => "$model->{ldnoWord} discount",
        defaultFormat => '%copy',
        cols          => $combinations,
        sources       => $model->{embeddedModelM}{objects}{table1037sources},
      )
      : Dataset(
        name => "Embedded network ($model->{ldnoWord}) discount percentages",
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

    if ( $model->{pcd}{portfolio} && $model->{pcd}{portfolio} =~ /5|7/i ) {

        my $cdcmLevels = Labelset( list => [ split /\n/, <<EOL] );
LV demand
LV Sub demand or LV generation
HV demand or LV Sub generation
HV generation
EOL

        my $addBoundaryLevels = Labelset( list => [ split /\n/, <<EOL] );
Boundary 0000
Boundary 132kV
Boundary 132kV/EHV
Boundary EHV
Boundary HVplus
EOL

        my $addDiscounts = !$model->{embeddedModelM}
          ? Dataset(
            name =>
"$model->{ldnoWord} discount percentage for higher-boundary cases",
            cols          => $cdcmLevels,
            rows          => $addBoundaryLevels,
            defaultFormat => '%hard',
            data          => [
                map {
                    [ map { '' } @{ $addBoundaryLevels->{list} } ]
                } @{ $cdcmLevels->{list} }
            ],
            number     => 1181,
            dataset    => $model->{dataset},
            appendTo   => $model->{inputTables},
            validation => {
                validate      => 'decimal',
                criteria      => '>=',
                value         => 0,
                input_title   => "$model->{ldnoWord} discount:",
                input_message => 'At least zero',
                error_title   => "Invalid $model->{ldnoWord} discount",
                error_message => "Invalid $model->{ldnoWord} discount"
                  . ' (negative number or unused cell).',
            },
          )
          : $model->{embeddedModelM}{objects}{table1181sources};

        $model->{pcd}{discount} = new SpreadsheetModel::Custom(
            name => 'Discount for each tariff (except for some fixed charges)',
            rows => $model->{pcd}{allTariffsByEndUser},
            defaultFormat => '%copy',
            arithmetic    => join(
                ' or ', 0,
                (
                    map { "A$_"; }
                      1 .. ( ref $addDiscounts eq 'ARRAY' ? @$addDiscounts : 1 )
                ),
                $rawDiscount ? 'A9' : ()
            ),
            custom     => [ '=A1', $rawDiscount ? '=A9' : (), ],
            objectType => 'Special copy',
            arguments  => {
                ref $addDiscounts eq 'ARRAY'
                ? ( map { ( "A$_" => $addDiscounts->[ $_ - 1 ] ); }
                      1 .. @$addDiscounts )
                : ( A1 => $addDiscounts ),
                $rawDiscount ? ( A9 => $rawDiscount ) : (),
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    local $_ = $model->{pcd}{allTariffsByEndUser}{list}[$y];
                    return 0, $format
                      if /^$model->{ldnoWord} [HL]V: / && /gener/i;
                    return '', $format, $formula->[1],
                      qr/\bA9\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{A9},
                        $colh->{A9} + ( /^HV/i ? 4 : /^LV Sub/i ? 3 : 2 ),
                        1, 1, )
                      if s/^$model->{ldnoWord} HV: //;
                    return '', $format, $formula->[1],
                      qr/\bA9\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{A9}, $colh->{A9} + 1,
                        1,           1, )
                      if s/^$model->{ldnoWord} LV: //;
                    my $yt;
                    $yt = 0 if s/^$model->{ldnoWord} 0000: //;
                    $yt = 2 if s/^$model->{ldnoWord} 132kV\/EHV: //;
                    $yt = 1 if s/^$model->{ldnoWord} 132kV: //;
                    $yt = 3 if s/^$model->{ldnoWord} EHV: //;
                    $yt = 4 if s/^$model->{ldnoWord} HVplus: //;
                    return 0, $format unless defined $yt;
                    $x =
                        /^HV/i && /gener/i     ? 3
                      : /^HV/i                 ? 2
                      : /^LV Sub/i && /gener/i ? 2
                      : /^LV Sub/i             ? 1
                      : /gener/i               ? 1
                      :                          0;
                    return '#VALUE!', $format if $x > 3;
                    '', $format, $formula->[0],
                      qr/\bA1\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{A1} + $yt,
                        $colh->{A1} + $x,
                        1, 1,
                      );
                };
            },
        );

    }

    else {

        push @{ $model->{volumeData} },
          $model->{pcd}{discount} =
          $rawDiscount->{rows} ? $rawDiscount : SumProduct(
            name => 'Discount for each tariff (except for some fixed charges)',
            defaultFormat => '%softnz',
            matrix        => Constant(
                name          => 'Discount map',
                defaultFormat => '0con',
                rows          => $model->{pcd}{allTariffsByEndUser},
                cols          => $combinations,
                byrow         => 1,
                data          => \@discountMapData,
            ),
            vector => $rawDiscount
          );

    }

    my $ldnoGenerators = Labelset(
        name => "Generators on $model->{ldnoWord} LV "
          . "and $model->{ldnoWord} HV networks",
        list => [
            grep { /(?:ID|LD|Q)NO [HL]V: /i && /gener/i }
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
                name => '100 per cent discount for generators'
                  . " on $model->{ldnoWord} networks",
                defaultFormat => '%con',
                rows          => $ldnoGenerators,
                data          => [ [ map { 1 } @{ $ldnoGenerators->{list} } ] ]
            ),
            $model->{pcd}{discount}
        ]
    );

    ( $model->{pcd}{volumeData} ) =
      $model->volumes( $model->{pcd}{allTariffsByEndUser},
        $allEndUsers, $nonExcludedComponents, $componentMap, 'no aggregation' );

    if ( $model->{fixedChargeAdders} ) {
        my $targetRevenueMethod;
        if ( $model->{targetRevenue} =~ /dcp421|2024/i ) {
            $targetRevenueMethod = 'table1001_2024';
        }
        elsif ( $model->{targetRevenue} =~ /dcp334|2019/i ) {
            $targetRevenueMethod = 'table1001_2019';
        }
        my @adderNames = ( 'Domestic demand', 'Metered demand' );
        ( undef, my @adderTargetRevenue ) = $model->$targetRevenueMethod();
        $model->{revenueFromElsewhere} = Arithmetic(
            name          => 'Total revenue from fixed charge adders (£/year)',
            defaultFormat => '0soft',
            mustCopy      => 1,
            arithmetic    => '=A10+A11',
            arguments     => {
                map { ( "A1$_" => $adderTargetRevenue[$_] ); }
                  0 .. $#adderTargetRevenue,
            },
        );
        Columnset(
            name    => 'Revenues from fixed charge adders',
            columns => [ @adderTargetRevenue, $model->{revenueFromElsewhere}, ],
        );
        my $rowset = $model->{pcd}{volumeData}{'Fixed charge p/MPAN/day'}{rows};
        my @applicationRules = (
            Constant(
                name          => 'Domestic demand',
                defaultFormat => '0con',
                rows          => $rowset,
                data          => [
                    map {
                            /non.domestic|related|additional/i ? 0
                          : /domestic/i                        ? 1
                          :                                      0;
                    } @{ $rowset->{list} }
                ],
            ),
            Constant(
                name          => 'Metered demand',
                defaultFormat => '0con',
                rows          => $rowset,
                data          => [
                    map { /unmetered|ums|gener|related|additional/i ? 0 : 1; }
                      @{ $rowset->{list} }
                ],
            )
        );
        Columnset(
            name    => 'Fixed charge adders application matrix',
            columns => \@applicationRules,
        );
        my @adderRates = map {
            Arithmetic(
                name          => $adderNames[$_],
                defaultFormat => '0.00soft',
                arithmetic    => '=A1/SUMPRODUCT(A2_A3,A4_A5)',
                arguments     => {
                    A1    => $adderTargetRevenue[$_],
                    A2_A3 => $applicationRules[$_],
                    A4_A5 =>
                      $model->{pcd}{volumeData}{'Fixed charge p/MPAN/day'},
                },
            );
        } 0 .. $#adderNames;
        Columnset(
            name    => 'Fixed charge adder rates (£/MPAN/year)',
            columns => \@adderRates,
        );
        $model->{pcd}{preRoundingFiddles}{'Fixed charge p/MPAN/day'} =
          Arithmetic(
            name          => 'Fixed charge adder (p/MPAN/day)',
            defaultFormat => '0.00soft',
            arithmetic    => '=(A1*A2+A3*A4)/A5*100',
            arguments     => {
                A1 => $applicationRules[0],
                A2 => $adderRates[0],
                A3 => $applicationRules[1],
                A4 => $adderRates[1],
                A5 => $daysAfter,
            },
          );
        push @{ $model->{edcmTables} },
          Stack(
            name    => 'EDCM input data ⇒1183. Fixed charge adder (p/MPAN/day)',
            rows    => $applicationRules[0]{rows},
            sources =>
              [ $model->{pcd}{preRoundingFiddles}{'Fixed charge p/MPAN/day'} ]
          ) if $model->{edcmTables};
    }

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

    my ( $model, $allComponents, $tariffTable, $daysInYear, $componentMap ) =
      @_;

    my $allTariffs = $model->{pcd}{allTariffsByEndUser};
    my $volumeData = $model->{pcd}{volumeData};

    my %rowFormats;
    foreach my $component (@$allComponents) {
        $rowFormats{$component} =
          [ map { $componentMap->{$_}{$component} ? undef : 'unavailable'; }
              @{ $allTariffs->{list} } ];
    }

    push @{ $model->{roundingResults} },
      Columnset(
        name    => 'All the way tariffs',
        columns => $model->{allTariffColumns},
      );

    push @{ $model->{edcmTables} },
      Columnset(
        name => 'EDCM input data ⇒1182. CDCM end user'
          . ' tariffs excluding undiscountable elements',
        singleRowName => 'CDCM end user tariffs',
        columns       =>
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

    my ( %postRoundingFiddles, %postRoundingOverrides );
    if ( $model->{bung} || $model->{electionBung} ) {
        my $bung = Dataset(
            name               => 'Bung (p/MPAN/day)',
            defaultFormat      => '0.00hardpm',
            rows               => $allTariffs,
            rowFormats         => $rowFormats{'Fixed charge p/MPAN/day'},
            number             => 1098,
            appendTo           => $model->{inputTables},
            dataset            => $model->{dataset},
            data               => [ map { '' } @{ $allTariffs->{list} } ],
            usePlaceholderData => 1,
            validation         =>
              {    # some validation is needed to enable lenient locking
                validate => 'decimal',
                criteria => 'between',
                minimum  => -666_666_666.666,
                maximum  => 666_666_666.666,
              },
        );
        push @{ $model->{otherTotalRevenues} },
          my $totalImpactOfBung = Arithmetic(
            name          => 'Revenue impact of bung (£, not accounted for)',
            defaultFormat => '0soft',
            arithmetic    => '=0.01*A1*SUMPRODUCT(A2_A3,A4_A5)',
            arguments     => {
                A1    => $daysInYear,
                A2_A3 => $bung,
                A4_A5 => $volumeData->{'Fixed charge p/MPAN/day'}
            }
          );
        $model->{sharedData}
          ->addStats( 'DNO-wide aggregates', $model, $totalImpactOfBung )
          if $model->{sharedData};
        $postRoundingFiddles{'Fixed charge p/MPAN/day'} = $bung;
    }
    elsif ( $model->{fiddle} ) {
        my @fiddles = map {
            Dataset(
                name               => $_,
                rows               => $allTariffs,
                rowFormats         => $rowFormats{$_},
                cols               => $tariffTable->{$_}{cols},
                defaultFormat      => /day/ ? '0.00hardpm' : '0.000hardpm',
                data               => [ map { 0; } @{ $allTariffs->{list} } ],
                usePlaceholderData => 1,
                validation         =>
                  {    # some validation is needed to enable lenient locking
                    validate => 'decimal',
                    criteria => 'between',
                    minimum  => -666_666_666.666,
                    maximum  => 666_666_666.666,
                  },
            );
        } @$allComponents;
        Columnset(
            name     => 'Fiddle terms',
            number   => 1099,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            columns  => \@fiddles,
        );
        @postRoundingFiddles{@$allComponents} = @fiddles;
    }
    elsif ( $model->{override} ) {
        my @overrides = map {
            Dataset(
                name               => $_,
                rows               => $allTariffs,
                rowFormats         => $rowFormats{$_},
                cols               => $tariffTable->{$_}{cols},
                defaultFormat      => /day/ ? '0.00hard' : '0.000hard',
                data               => [ map { 0; } @{ $allTariffs->{list} } ],
                usePlaceholderData => 1,
                validation         =>
                  {    # some validation is needed to enable lenient locking
                    validate => 'decimal',
                    criteria => 'between',
                    minimum  => -666_666_666.666,
                    maximum  => 666_666_666.666,
                  },
            );
        } @$allComponents;
        Columnset(
            name     => 'Override tariffs',
            number   => 1097,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            columns  => \@overrides,
        );
        @postRoundingOverrides{@$allComponents} = @overrides;
    }

    my $newTariffTable = {
        map {
            $_ => Arithmetic(
                name => $_,
                m%p/k(W|VAr)h% ? ()
                : ( defaultFormat => '0.00soft' ),
                arithmetic => $postRoundingOverrides{$_}
                  && $model->{override} !~ /diff/i ? '=A7'
                : $model->{model100} ? '=A2*(1-A1)'
                : (   ( $postRoundingFiddles{$_} ? '=A3+' : '=' )
                    . 'ROUND('
                      . ( $model->{pcd}{preRoundingFiddles}{$_} ? 'A4+' : '' )
                      . 'A2*(1-A1),'
                      . ( /kWh|kVArh/ ? 3 : 2 )
                      . ')' )
                  . (
                         $postRoundingOverrides{$_}
                      && $model->{override} =~ /diff/i ? '-A7' : ''
                  ),
                rows       => $allTariffs,
                rowFormats => $rowFormats{$_},
                cols       => $tariffTable->{$_}{cols},
                arguments  => {
                    A2 => $tariffTable->{$_},
                    A1 => /fix/i ? $model->{pcd}{discountFixed}
                    : $model->{pcd}{discount},
                    $postRoundingFiddles{$_}
                    ? ( A3 => $postRoundingFiddles{$_} )
                    : (),
                    $model->{pcd}{preRoundingFiddles}{$_}
                    ? ( A4 => $model->{pcd}{preRoundingFiddles}{$_} )
                    : (),
                    $postRoundingOverrides{$_}
                    ? ( A7 => $postRoundingOverrides{$_} )
                    : (),
                },
            );
        } @$allComponents
    };

    $allTariffs, $allTariffs, $volumeData, $unitsInYear, $newTariffTable;

}

1;
