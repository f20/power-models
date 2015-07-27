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

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Shortcuts ':all';

sub operating {

    my (
        $model,                         $assetLevels,
        $drmLevels,                     $drmExitLevels,
        $operatingLevels,               $operatingDrmLevels,
        $operatingDrmExitLevels,        $customerLevels,
        $operatingCustomerLevels,       $forecastSml,
        $allTariffsByEndUser,           $unitsInYear,
        $daysInYear,                    $lineLossFactors,
        $diversityAllowances,           $componentMap,
        $volumeData,                    $modelGrossAssetsByLevel,
        $modelCostToSml,                $modelSml,
        $serviceModelAssetsPerCustomer, $serviceModelAssetsPerAnnualMwh,
        $siteSpecificSoleUseAssets
    ) = @_;

    my $operatingExpenditureCodedByLevel =
      $model->{opCoded} && $model->{opCoded} =~ /lev/i
      ? Dataset(
        name => 'Transmission exit charges and other'
          . ' operating expenditure coded'
          . ' by network level (£/year)',
        validation => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
        lines         => ['Source: forecast.'],
        number        => 1055,
        appendTo      => $model->{inputTables},
        dataset       => $model->{dataset},
        cols          => $operatingLevels,
        defaultFormat => '0hardnz',
        data => [ 10_000_000, map { [0] } 2 .. @{ $operatingLevels->{list} } ],
      )
      : Stack(
        name => 'Operating expenditure coded by network level (£/year)',
        defaultFormat => '0copynz',
        cols          => $operatingLevels,
        sources       => [
            Dataset(
                name       => 'Transmission exit charges (£/year)',
                validation => {
                    validate => 'decimal',
                    criteria => '>=',
                    value    => 0,
                },
                lines    => ['Source: forecast.'],
                number   => 1055,
                appendTo => $model->{inputTables},
                dataset  => $model->{dataset},
                cols     => Labelset( list => [ $operatingLevels->{list}[0] ] ),
                defaultFormat => '0hardnz',
                data          => [10_000_000],
            ),
            Constant(
                name          => 'Zero for levels other than transmission exit',
                data          => [ map { [0] } @{ $operatingLevels->{list} } ],
                cols          => $operatingLevels,
                defaultFormat => '0connz'
            )
        ]
      );

    $model->{edcmTables}[0][5] = Stack(
        defaultFormat => '0hard',
        sources       => [ $operatingExpenditureCodedByLevel->{sources}[0] ]
    ) if $model->{edcmTables};

    my $operatingLvCustomer =
      Labelset( list =>
          [ grep { /LV customer/i } @{ $operatingCustomerLevels->{list} } ] );

    my $umsOperatingExpenditure;

    if ( $model->{opCoded} && $model->{opCoded} =~ /ums/i ) {

        push @{ $model->{optionLines} },
          'Pre-coded unmetered operating expenditure';

        $umsOperatingExpenditure = Dataset(
            name => 'Pre-coded operating expenditure'
              . ' on unmetered customer assets (£/year)',
            validation => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
            defaultFormat => '0hard',
            cols          => $operatingLvCustomer,
            data          => [800_000]
        );

        push @{ $model->{operatingExpenditure} },
          Columnset(
            name     => 'Operating expenditure coded by category',
            number   => 1055,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            columns  => [
                  $operatingExpenditureCodedByLevel->{sources}
                ? $operatingExpenditureCodedByLevel->{sources}[0]
                : $operatingExpenditureCodedByLevel,
                $umsOperatingExpenditure
            ]
          );

    }

    push @{ $model->{operatingExpenditure} }, $operatingExpenditureCodedByLevel
      unless $umsOperatingExpenditure
      && !$operatingExpenditureCodedByLevel->{sources};

    my $operatingExpenditureToSplit;

    if ( $model->{opAlloc} ) {
        $operatingExpenditureToSplit = Dataset(
            number     => 1057,
            name       => 'Other operating expenditure (£/year)',
            validation => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
            appendTo      => $model->{inputTables},
            dataset       => $model->{dataset},
            lines         => [ 'Source: forecast.', ],
            defaultFormat => '0hard',
            data => [ 90_000_000 - ( $umsOperatingExpenditure ? 800_000 : 0 ) ],
        );
    }
    else {
        my @otex = (
            Dataset(
                name          => 'Direct cost (£/year)',
                defaultFormat => '0hard',
                data =>
                  [ 90_000_000 - ( $umsOperatingExpenditure ? 800_000 : 0 ) ],
                validation => {
                    validate => 'decimal',
                    criteria => '>=',
                    value    => 0,
                }
            ),
            Dataset(
                name          => 'Indirect cost (£/year)',
                defaultFormat => '0hard',
                data          => [0],
                validation    => {
                    validate => 'decimal',
                    criteria => '>=',
                    value    => 0,
                }
            ),
            Dataset(
                name          => 'Indirect cost proportion',
                defaultFormat => '%hard',
                data          => [0.6],
                validation    => {
                    validate => 'decimal',
                    criteria => 'between',
                    minimum  => 0,
                    maximum  => 1,
                }
            ),
            Dataset(
                name          => 'Network rates (£/year)',
                defaultFormat => '0hard',
                data          => [0],
                validation    => {
                    validate => 'decimal',
                    criteria => '>=',
                    value    => 0,
                }
            )
        );

        @{ $model->{edcmTables}[0] }[ 6 .. 8 ] =
          map { Stack( defaultFormat => '0hard', sources => [$_] ); }
          @otex[ 0, 1, 3 ]
          if $model->{edcmTables};

        Columnset(
            name          => 'Other expenditure',
            singleRowName => 'Other expenditure',
            appendTo      => $model->{inputTables},
            dataset       => $model->{dataset},
            columns       => \@otex,
            number        => 1059
        );
        $operatingExpenditureToSplit = Arithmetic(
            name => 'Amount of expenditure to be allocated'
              . ' according to asset values (£/year)',
            singleRowName => 'Other expenditure',
            defaultFormat => '0soft',
            arithmetic    => '=A1+A4+A2*A3',
            arguments     => { map { ( "A$_" => $otex[ $_ - 1 ] ) } 1 .. 4 }
        );

    }

    my @serviceModelColumns;
    my $serviceModelAssets;

    push @serviceModelColumns,
      $serviceModelAssets = SumProduct(
        name          => 'Service model assets (£) scaled by user count',
        matrix        => $serviceModelAssetsPerCustomer,
        vector        => $volumeData->{'Fixed charge p/MPAN/day'},
        defaultFormat => '0soft'
      ) if $serviceModelAssetsPerCustomer;

    my $serviceModelAssetsFromAnnualMwh;
    my $unmeteredUnits;

    if ($serviceModelAssetsPerAnnualMwh) {
        $unmeteredUnits = Stack(
            name    => 'Annual consumption by tariff for unmetered users (MWh)',
            sources => [$unitsInYear],
            rows    => $serviceModelAssetsPerAnnualMwh->{rows} || Labelset(
                list => [
                    grep { /un-?met/i && !/^ldno/i }
                      @{ $unitsInYear->{rows}{list} }
                ]
            ),
        );
        push @serviceModelColumns,
          $serviceModelAssetsFromAnnualMwh =
          $serviceModelAssetsPerAnnualMwh->{rows}
          ? SumProduct(
            name          => 'Service model assets (£) scaled by annual MWh',
            matrix        => $serviceModelAssetsPerAnnualMwh,
            vector        => $unmeteredUnits,
            defaultFormat => '0soft'
          )
          : Arithmetic(
            name          => 'Service model assets (£) scaled by annual MWh',
            defaultFormat => '0soft',
            arithmetic    => '=A1*A2',
            arguments     => {
                A1 => $serviceModelAssetsPerAnnualMwh,
                A2 => GroupBy(
                    name          => 'Total unmetered units',
                    defaultFormat => '0softnz',
                    source        => $unmeteredUnits
                ),
            },
          );

        if ($serviceModelAssets) {

            push @serviceModelColumns,
              my $reshapedServiceModelAssetsFromAnnualMwh = Stack(
                name    => 'Service model assets (£) scaled by annual MWh',
                rows    => 0,
                cols    => $serviceModelAssets->{cols},
                sources => [$serviceModelAssetsFromAnnualMwh]
              );

            push @serviceModelColumns,
              $serviceModelAssets = Arithmetic(
                name       => 'Service model assets (£)',
                arithmetic => '=A1+A2',
                arguments  => {
                    A1 => $serviceModelAssets,
                    A2 => $reshapedServiceModelAssetsFromAnnualMwh
                },
                defaultFormat => '0soft'
              );

        }
        else {
            $serviceModelAssets = $serviceModelAssetsFromAnnualMwh;
        }
    }

    Columnset(
        name    => 'Service model asset data',
        columns => \@serviceModelColumns
    ) if @serviceModelColumns;

    my $modelAssetsByLevel = Stack
      name          => 'Model assets (£) scaled by demand forecast',
      cols          => $assetLevels,
      defaultFormat => '0copynz',
      rows          => 0,
      sources       => [
        $siteSpecificSoleUseAssets ? Arithmetic(
            name => 'Network model assets scaled by load forecast'
              . ', plus site specific sole use assets (£)',
            arithmetic    => '=IF(A5,A1*A2/A3/1000,0)+A4',
            defaultFormat => '0soft',
            cols          => $drmLevels,
            rows          => 0,
            arguments     => {
                A1 => $forecastSml,
                A2 => $modelGrossAssetsByLevel,
                A3 => $modelSml,
                A4 => $siteSpecificSoleUseAssets,
                A5 => $modelSml,
            }
          ) : Arithmetic(
            name          => 'Network model assets (£) scaled by load forecast',
            arithmetic    => '=IF(A5,A1*A2/A3/1000,0)',
            defaultFormat => '0soft',
            cols          => $drmLevels,
            rows          => 0,
            arguments     => {
                A1 => $forecastSml,
                A2 => $modelGrossAssetsByLevel,
                A3 => $modelSml,
                A5 => $modelSml,
            }
          ),
        $serviceModelAssets ? $serviceModelAssets : ()
      ];

    if ( $model->{opAlloc} && $model->{opAlloc} =~ /OptA/i ) {
        $modelAssetsByLevel = Stack(
            name    => 'Network model assets (£)',
            cols    => $assetLevels,
            rows    => 0,
            sources => [$modelGrossAssetsByLevel]
        );
    }

    my $modelAssetsByLevelPossiblyScaled = $modelAssetsByLevel;

    my $multipliers;

    if ( $model->{opAlloc} && $model->{opAlloc} =~ /LS/i ) {

        $multipliers = Dataset(
            name       => 'Operating expenditure intensity multipliers',
            validation => {
                validate => 'decimal',
                criteria => '>',
                value    => 0,
            },
            number   => 1056,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            cols     => 0,
            rows     => $assetLevels->{accepts}[0] || $assetLevels,
            data     => [
                map {
                        /gsp/i ? undef
                      : /132/i ? .5
                      : /EHV/  ? .8
                      : /HV/i  ? 1
                      : 1.4
                } @{ $assetLevels->{list} }
            ],
        );

        $modelAssetsByLevelPossiblyScaled = Arithmetic(
            name          => 'Scaling factor after applying multipliers',
            arithmetic    => '=A1*A2',
            defaultFormat => '0soft',
            arguments     => {
                A1 => $modelAssetsByLevel,
                A2 => $multipliers,
            }
        );

    }

    my $modelAssetsPossiblyScaledTotal = GroupBy(
        name          => 'Denominator for allocation of operating expenditure',
        source        => $modelAssetsByLevelPossiblyScaled,
        defaultFormat => $modelAssetsByLevelPossiblyScaled->{defaultFormat}
    );

    if ( $model->{opAlloc} && $model->{opAlloc} =~ /OptC$/i ) {
        $modelAssetsPossiblyScaledTotal = Arithmetic(
            name       => 'Denominator for allocation of operating expenditure',
            rows       => 0,
            cols       => Labelset( list => [ $modelSml->{rows}{list}[0] ] ),
            arithmetic => '=A1*A2',
            arguments  => {
                A1 => $forecastSml,
                A2 => GroupBy(
                    name   => 'Total network model £/kW SML/year',
                    source => $modelCostToSml
                )
            },
            defaultFormat => $modelAssetsByLevelPossiblyScaled->{defaultFormat}
        );
    }

    push @{ $model->{edcmTables} },
      Stack(
        name          => 'Assets in CDCM model (£)',
        defaultFormat => '0hard',
        number        => 1131,
        sources       => [$modelAssetsByLevelPossiblyScaled]
      ) if $model->{edcmTables};

    push @{ $model->{operatingExpenditure} },
      Columnset(
        name    => 'Data for allocation of operating expenditure',
        columns => [
            $modelAssetsByLevelPossiblyScaled, $modelAssetsPossiblyScaledTotal
        ]
      );

    my $abaters;

    $abaters = Dataset(
        name => 'Abatement of forecast operating expenditure'
          . ' to reflect forward looking assets',
        validation => {
            validate => 'decimal',
            criteria => 'between',
            minimum  => -4,
            maximum  => 1,
        },
        cols => 0,
        rows => $assetLevels->{accepts}[0] || $assetLevels,
        ,
        data => [ map { /^LV/ ? 0.15 : 0 } @{ $assetLevels->{list} } ],
        defaultFormat => '%hardnz'
    ) if $model->{opAlloc} && $model->{opAlloc} =~ /FS/i;

    Columnset(
        number   => 1056,
        appendTo => $model->{inputTables},
        dataset  => $model->{dataset},
        name     => 'Operating expenditure allocation parameters',
        columns =>
          [ $multipliers ? $multipliers : (), $abaters ? $abaters : () ]
    ) if $multipliers && $abaters;

    push @{ $model->{optionLines} }, 'Operating expenditure multipliers'
      if $multipliers;

    push @{ $model->{optionLines} }, 'Operating expenditure abaters'
      if $abaters;

    push @{ $model->{optionLines} },
      'Operating expenditure' . ' allocated by asset values'
      unless $multipliers || $abaters;

    if (undef) {

        $abaters = Arithmetic(
            name       => "$abaters->{name} (transposed)",
            arithmetic => '=A1',
            cols       => $operatingLevels,
            rows       => 0,
            arguments  => { A1 => $abaters }
        ) if $abaters;

        $multipliers = Arithmetic(
            name       => "$multipliers->{name} (transposed)",
            arithmetic => '=A1',
            cols       => $operatingLevels,
            rows       => 0,
            arguments  => { A1 => $multipliers }
        ) if $multipliers;

    }

    push @{ $model->{operatingExpenditure} },
      my $operatingExpenditureTotalByLevel = Arithmetic
      name => 'Total operating expenditure by network level '
      . ( $umsOperatingExpenditure ? 'excluding unmetered services ' : '' )
      . ' (£/year)',
      cols          => $operatingLevels,
      rows          => 0,
      defaultFormat => '0softnz',
      arithmetic    => '=A1+A2/A3*A4' . ( $abaters ? '*(1-A6)' : '' ),
      arguments     => {
        A1 => $operatingExpenditureCodedByLevel,
        A2 => $operatingExpenditureToSplit,
        A3 => $modelAssetsPossiblyScaledTotal,
        A4 => $modelAssetsByLevelPossiblyScaled,
        $abaters ? ( A6 => $abaters ) : (),
      };

    push @{ $model->{operatingExpenditure} },
      my $operatingExpenditurePercentages = Arithmetic(
        name          => 'Operating expenditure percentage by network level',
        defaultFormat => '%softnz',
        cols          => $operatingLevels,
        arithmetic    => '=IF(A2="","",IF(A3>0,A1/A4,0))',
        arguments     => {
            A1 => $operatingExpenditureTotalByLevel,
            A2 => $modelAssetsByLevel,
            A3 => $modelAssetsByLevel,
            A4 => $modelAssetsByLevel,
        }
      );

    my ($siteSpecificOperatingCost);

    if ($siteSpecificSoleUseAssets) {
        push @{ $model->{siteSpecific} },
          $siteSpecificOperatingCost = Arithmetic(
            name => 'Operating expenditure for site-specific'
              . ' sole use assets (£/year)',
            arithmetic    => '=A1*A5',
            defaultFormat => '0softnz',
            cols          => $operatingDrmLevels,
            arguments     => {
                A1 => $operatingExpenditurePercentages,
                A5 => $siteSpecificSoleUseAssets,
            }
          );
        $operatingExpenditureTotalByLevel = Stack(
            name => 'Total operating expenditure by network level '
              . 'excluding site-specific assets '
              . ' (£/year)',
            cols    => $operatingLevels,
            sources => [
                Arithmetic(
                    name => 'Operating expenditure excluding site-specific'
                      . ' sole use assets (£/year)',
                    arithmetic    => '=A7-A1*A5',
                    defaultFormat => '0softnz',
                    cols          => $operatingDrmLevels,
                    arguments     => {
                        A1 => $operatingExpenditurePercentages,
                        A5 => $siteSpecificSoleUseAssets,
                        A7 => $operatingExpenditureTotalByLevel,
                    }
                ),
                $operatingExpenditureTotalByLevel
            ]
        );
    }

    my $operatingCostToSml = Arithmetic(
        name => 'Unit operating expenditure based'
          . ' on simultaneous maximum load (£/kW/year)',
        arithmetic => '=IF(A3>0,A1/A2,0)',
        rows       => 0,
        cols       => $operatingDrmExitLevels,
        arguments  => {
            A2 => $forecastSml,
            A3 => $forecastSml,
            A1 => $operatingExpenditureTotalByLevel,
        }
    );

    push @{ $model->{operatingExpenditure} }, $operatingCostToSml;

    my ( $operatingCostByCustomer, $operatingCostByAnnualMwh );

    if ($serviceModelAssetsPerCustomer) {

        $operatingCostByCustomer = GroupBy(
            name =>
              'Operating expenditure for customer assets p/MPAN/day total',
            rows   => $allTariffsByEndUser,
            cols   => 0,
            source => Arithmetic(
                name       => 'Operating expenditure p/MPAN/day by level',
                rows       => $allTariffsByEndUser,
                cols       => $operatingCustomerLevels,
                arithmetic => '=100/A2*A1*A3',
                arguments  => {
                    A1 => $operatingExpenditurePercentages,
                    A2 => $daysInYear,
                    A3 => $serviceModelAssetsPerCustomer,
                }
            )
        );

        push @{ $model->{operatingExpenditure} },
          Columnset(
            name => 'Operating expenditure for customer assets p/MPAN/day',
            columns =>
              [ $operatingCostByCustomer->{source}, $operatingCostByCustomer ]
          );

    }

    push @{ $model->{operatingExpenditure} },
      $operatingCostByAnnualMwh = Arithmetic(
        name => 'Operating expenditure for unmetered customer assets (p/kWh)',
        rows => $unmeteredUnits->{rows},
        cols => $operatingLvCustomer,
        arithmetic => $umsOperatingExpenditure ? '=0.1*(A3*A1+A4/A5)'
        : '=0.1*A1*A3',
        arguments => {
            A1 => $operatingExpenditurePercentages,
            A3 => $serviceModelAssetsPerAnnualMwh,
            $umsOperatingExpenditure
            ? (
                A4 => $umsOperatingExpenditure,
                A5 => GroupBy(
                    defaultFormat => '0softnz',
                    name   => 'Total unmetered MWh/year for customer assets',
                    source => $unmeteredUnits
                )
              )
            : ()
        }
      ) if $serviceModelAssetsPerAnnualMwh;

    $operatingCostToSml, $operatingCostByCustomer, $operatingCostByAnnualMwh,
      $siteSpecificOperatingCost;

}

1;
