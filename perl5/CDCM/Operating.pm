package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 DCUSA Limited and contributors. All rights reserved.

Redistribution and use in source and binary forms, with or without modification,
are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation and/or
other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY DCUSA LIMITED AND CONTRIBUTORS "AS IS" AND ANY
EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL DCUSA LIMITED OR CONTRIBUTORS BE LIABLE FOR ANY
DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

=cut

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Shortcuts ':all';

sub operating {

    my (
        $model,                          $assetLevels,
        $drmLevels,                      $drmExitLevels,
        $operatingLevels,                $operatingDrmLevels,
        $operatingDrmExitLevels,         $customerLevels,
        $operatingCustomerLevels,        $forecastSml,
        $allTariffsByEndUser,            $unitsInYear,
        $loadFactors,                    $daysInYear,
        $lineLossFactors,                $diversityAllowances,
        $componentMap,                   $volumeData,
        $modelGrossAssetsByLevel,        $modelCostToSml,
        $modelSml,                       $serviceModelAssetsPerCustomer,
        $serviceModelAssetsPerAnnualMwh, $siteSpecificSoleUseAssets
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
            arithmetic    => '=IV1+IV4+IV2*IV3',
            arguments     => { map { ( "IV$_" => $otex[ $_ - 1 ] ) } 1 .. 4 }
        );

    }

    my @serviceModelColumns;
    my $serviceModelAssets;

    push @serviceModelColumns, $serviceModelAssets = SumProduct(
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
        push @serviceModelColumns, $serviceModelAssetsFromAnnualMwh =
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
            arithmetic    => '=IV1*IV2',
            arguments     => {
                IV1 => $serviceModelAssetsPerAnnualMwh,
                IV2 => GroupBy(
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

            push @serviceModelColumns, $serviceModelAssets = Arithmetic(
                name       => 'Service model assets (£)',
                arithmetic => '=IV1+IV2',
                arguments  => {
                    IV1 => $serviceModelAssets,
                    IV2 => $reshapedServiceModelAssetsFromAnnualMwh
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
            arithmetic    => '=IF(IV5,IV1*IV2/IV3/1000,0)+IV4',
            defaultFormat => '0soft',
            cols          => $drmLevels,
            rows          => 0,
            arguments     => {
                IV1 => $forecastSml,
                IV2 => $modelGrossAssetsByLevel,
                IV3 => $modelSml,
                IV4 => $siteSpecificSoleUseAssets,
                IV5 => $modelSml,
            }
          ) : Arithmetic(
            name          => 'Network model assets (£) scaled by load forecast',
            arithmetic    => '=IF(IV5,IV1*IV2/IV3/1000,0)',
            defaultFormat => '0soft',
            cols          => $drmLevels,
            rows          => 0,
            arguments     => {
                IV1 => $forecastSml,
                IV2 => $modelGrossAssetsByLevel,
                IV3 => $modelSml,
                IV5 => $modelSml,
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
            arithmetic    => '=IV1*IV2',
            defaultFormat => '0soft',
            arguments     => {
                IV1 => $modelAssetsByLevel,
                IV2 => $multipliers,
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
            arithmetic => '=IV1*IV2',
            arguments  => {
                IV1 => $forecastSml,
                IV2 => GroupBy(
                    name   => 'Total network model £/kW SML/year',
                    source => $modelCostToSml
                )
            },
            defaultFormat => $modelAssetsByLevelPossiblyScaled->{defaultFormat}
        );
    }

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
            arithmetic => '=IV1',
            cols       => $operatingLevels,
            rows       => 0,
            arguments  => { IV1 => $abaters }
        ) if $abaters;

        $multipliers = Arithmetic(
            name       => "$multipliers->{name} (transposed)",
            arithmetic => '=IV1',
            cols       => $operatingLevels,
            rows       => 0,
            arguments  => { IV1 => $multipliers }
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
      arithmetic    => '=IV1+IV2/IV3*IV4' . ( $abaters ? '*(1-IV6)' : '' ),
      arguments     => {
        IV1 => $operatingExpenditureCodedByLevel,
        IV2 => $operatingExpenditureToSplit,
        IV3 => $modelAssetsPossiblyScaledTotal,
        IV4 => $modelAssetsByLevelPossiblyScaled,
        $abaters ? ( IV6 => $abaters ) : (),
      };

    push @{ $model->{operatingExpenditure} },
      my $operatingExpenditurePercentages = Arithmetic(
        name          => 'Operating expenditure percentage by network level',
        defaultFormat => '%softnz',
        cols          => $operatingLevels,
        arithmetic    => '=IF(IV2="","",IF(IV3>0,IV1/IV4,0))',
        arguments     => {
            IV1 => $operatingExpenditureTotalByLevel,
            IV2 => $modelAssetsByLevel,
            IV3 => $modelAssetsByLevel,
            IV4 => $modelAssetsByLevel,
        }
      );

    my ($siteSpecificOperatingCost);

    if ($siteSpecificSoleUseAssets) {
        push @{ $model->{siteSpecific} },
          $siteSpecificOperatingCost = Arithmetic(
            name => 'Operating expenditure for site-specific'
              . ' sole use assets (£/year)',
            arithmetic    => '=IV1*IV5',
            defaultFormat => '0softnz',
            cols          => $operatingDrmLevels,
            arguments     => {
                IV1 => $operatingExpenditurePercentages,
                IV5 => $siteSpecificSoleUseAssets,
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
                    arithmetic    => '=IV7-IV1*IV5',
                    defaultFormat => '0softnz',
                    cols          => $operatingDrmLevels,
                    arguments     => {
                        IV1 => $operatingExpenditurePercentages,
                        IV5 => $siteSpecificSoleUseAssets,
                        IV7 => $operatingExpenditureTotalByLevel,
                    }
                ),
                $operatingExpenditureTotalByLevel
            ]
        );
    }

    my $operatingCostToSml = Arithmetic(
        name => 'Unit operating expenditure based'
          . ' on simultaneous maximum load (£/kW/year)',
        arithmetic => '=IF(IV3>0,IV1/IV2,0)',
        rows       => 0,
        cols       => $operatingDrmExitLevels,
        arguments  => {
            IV2 => $forecastSml,
            IV3 => $forecastSml,
            IV1 => $operatingExpenditureTotalByLevel,
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
                arithmetic => '=100/IV2*IV1*IV3',
                arguments  => {
                    IV1 => $operatingExpenditurePercentages,
                    IV2 => $daysInYear,
                    IV3 => $serviceModelAssetsPerCustomer,
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
        arithmetic => $umsOperatingExpenditure ? '=0.1*(IV3*IV1+IV4/IV5)'
        : '=0.1*IV1*IV3',
        arguments => {
            IV1 => $operatingExpenditurePercentages,
            IV3 => $serviceModelAssetsPerAnnualMwh,
            $umsOperatingExpenditure
            ? (
                IV4 => $umsOperatingExpenditure,
                IV5 => GroupBy(
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
