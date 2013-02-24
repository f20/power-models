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

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Shortcuts ':all';

use CDCM::Aggregation;
use CDCM::AML;
use CDCM::Contributions;
use CDCM::Loads;
use CDCM::Matching;
use CDCM::NetworkSizer;
use CDCM::Operating;
use CDCM::Reactive;
use CDCM::Revenue;
use CDCM::Routeing;
use CDCM::ServiceModels;
use CDCM::Setup;
use CDCM::Sheets;
use CDCM::SML;
use CDCM::Standing;
use CDCM::Summary;
use CDCM::Tariffs;
use CDCM::TimeOfDay;
use CDCM::Yardsticks;

sub new {

    my $class = shift;
    my $model = {@_};
    bless $model, $class;

    $model->{timebands} = 3 unless $model->{timebands};
    $model->{timebands} = 10 if $model->{timebands} > 10;
    $model->{drm} = 'top500gsp' unless $model->{drm};

    # Keep CDCM::DataPreprocess out of the scope of revision numbers.
    if ( $model->{dataset}
        && keys %{ $model->{dataset} } )
    {
        if ( eval { require CDCM::DataPreprocess; } ) {
            $model->preprocessDataset;
        }
        else {
            warn $@;
        }
    }

    $model->{inputTables} = [];

    my ( $daysInYear, $daysBefore, $daysAfter, $modelLife, $annuityRate,
        $powerFactorInModel, $drmLevels, $drmExitLevels, )
      = $model->setUp;

    my ( $allTariffs, $allTariffsByEndUser, $allEndUsers, $allComponents,
        $nonExcludedComponents, $componentMap )
      = $model->tariffs;

    if ( $model->{pcd} ) {

        delete $model->{portfolio};
        delete $model->{boundary};

        $model->{pcd} = {
            allTariffsByEndUser => $allTariffsByEndUser,
            allTariffs          => $allTariffs
        };

        map { $_->{name} =~ s/> //g; } @{ $allEndUsers->{list} };

        $allTariffs = $allTariffsByEndUser = $allEndUsers;

    }

    my $coreLevels     = $drmLevels;
    my $coreExitLevels = $drmExitLevels;
    my ( $rerouteing13211, $rerouteingMatrix, $drmRerouteing );

    if ( $model->{extraLevels} ) {

        push @{ $model->{optionLines} }, 'Include a 132kV/HV network level';

        $drmLevels = Labelset(
            name => 'DRM network levels with 132kV/HV',
            list => [
                qw(132kV 132kV/EHV EHV EHV/HV 132kV/HV HV HV/LV),
                'LV circuits'
            ]
        );

        $drmExitLevels = Labelset
          name => 'DRM and transmission exit levels with 132kV/HV',
          list => [ 'GSPs', @{ $drmLevels->{list} } ];

        $rerouteing13211 = Dataset(
            name => 'Proportion of relevant load going through'
              . ' 132kV/HV direct transformation',
            number     => 1018,
            appendTo   => $model->{inputTables},
            dataset    => $model->{dataset},
            validation => {
                validate      => 'decimal',
                criteria      => 'between',
                minimum       => 0,
                maximum       => 1,
                input_title   => 'Proportion of load:',
                input_message => 'Between 0% and 100%',
                error_message => 'The proportion of load going through 132kV/HV'
                  . ' must be between 0% and 100%.'
            },
            defaultFormat => '%hardnz',
            cols          => Labelset( list => [ $drmLevels->{list}[4] ] ),
            rows          => Labelset( list => [ $coreLevels->{list}[3] ] ),
            data          => [ [0] ]
        );

        my $rerouteingMap = Constant(
            name          => 'Rerouteing matrix: default elements',
            defaultFormat => '%connz',
            cols          => $drmLevels,
            rows          => $coreLevels,
            byrow         => 1,
            data          => [
                [qw(1 0 0 0 0 0 0 0)], [qw(0 1 0 0 0 0 0 0)],
                [qw(0 0 1 0 0 0 0 0)], [qw(0 0 0 1 0 0 0 0)],
                [qw(0 0 0 0 0 1 0 0)], [qw(0 0 0 0 0 0 1 0)],
                [qw(0 0 0 0 0 0 0 1)],
            ]
        );

        my @sources = (
            $rerouteing13211,
            Arithmetic(
                name          => 'Proportion going through 132kV/EHV',
                cols          => Labelset( list => [ $drmLevels->{list}[1] ] ),
                rows          => Labelset( list => [ $coreLevels->{list}[1] ] ),
                defaultFormat => '%softnz',
                arithmetic    => '=1-IV1',
                arguments     => { IV1 => $rerouteing13211 }
            ),
            Arithmetic(
                name          => 'Proportion going through EHV',
                cols          => Labelset( list => [ $drmLevels->{list}[2] ] ),
                rows          => Labelset( list => [ $coreLevels->{list}[2] ] ),
                defaultFormat => '%softnz',
                arithmetic    => '=1-IV1',
                arguments     => { IV1 => $rerouteing13211 }
            ),
            Arithmetic(
                name          => 'Proportion going through EHV/HV',
                cols          => Labelset( list => [ $drmLevels->{list}[3] ] ),
                rows          => Labelset( list => [ $coreLevels->{list}[3] ] ),
                defaultFormat => '%softnz',
                arithmetic    => '=1-IV1',
                arguments     => { IV1 => $rerouteing13211 }
            ),
            $rerouteingMap
        );

        $drmRerouteing = Stack(
            name          => 'Rerouteing matrix for DRM network levels',
            rows          => $coreLevels,
            cols          => $drmLevels,
            defaultFormat => '%copynz',
            sources       => \@sources
        );

        $rerouteingMatrix = Stack(
            name          => 'Rerouteing matrix for all network levels',
            cols          => $coreExitLevels,
            rows          => $drmExitLevels,
            defaultFormat => '%copynz',
            sources       => [
                @sources,
                Constant(
                    name => 'Map GSP to GSP',
                    cols => Labelset( list => [ $drmExitLevels->{list}[0] ] ),
                    rows => Labelset( list => [ $coreExitLevels->{list}[0] ] ),
                    data => [ [1] ]
                )
            ]
          )

    }

    my (
        $assetCustomerLevels,     $assetDrmLevels,
        $assetLevels,             $chargingDrmExitLevels,
        $chargingLevels,          $customerChargingLevels,
        $customerLevels,          $networkLevels,
        $operatingCustomerLevels, $operatingDrmExitLevels,
        $operatingDrmLevels,      $operatingLevels,
        $routeingFactors,         $lineLossFactorsToGsp,
        $lineLossFactorsLevel,    $lineLossFactorsNetwork,
        $lineLossFactors,         $unitsLossAdjustment,
      )
      = $model->routeing( $allTariffsByEndUser, $coreLevels, $coreExitLevels,
        $drmLevels, $drmExitLevels, $rerouteingMatrix );

    # Now rebuild the 500 MW model using drmSizer

    my ( $modelCostToSml, $modelCostToAml, $diversityAllowances,
        $modelGrossAssetsByLevel, $modelSml, $networkModelCostToSml );

    (
        $coreLevels, $modelCostToSml, $diversityAllowances,
        $modelGrossAssetsByLevel, $modelSml
      )
      = $model->drmSizer(
        $modelLife,     $annuityRate,          $powerFactorInModel,
        $coreLevels,    $coreExitLevels,       $assetDrmLevels,
        $drmExitLevels, $lineLossFactorsLevel, $drmRerouteing
      );

    # Get load profiles

    my (
        $unitsEndUsers,              $generationEndUsers,
        $demandEndUsers,             $standingForFixedEndUsers,
        $generationCapacityEndUsers, $generationUnitsEndUsers,
        $fFactors,                   $loadCoefficients,
        $loadFactors
    ) = $model->loadProfiles( $allEndUsers, $componentMap );

    my (
        $proportionCoveredByContributions,          $proportionChargeable,
        $allLevelsProportionCoveredByContributions, $replacementShare
      )
      = $model->customerContributions(
        $assetDrmLevels,         $drmExitLevels,
        $operatingDrmExitLevels, $chargingDrmExitLevels,
        $allTariffsByEndUser
      );

    push @{ $model->{contributions} }, $replacementShare = Stack(
        name => 'Share of amount that relates'
          . ' to replacement of customer contributed assets',
        defaultFormat => '%connz',
        rows          => $allTariffsByEndUser,
        cols          => $chargingLevels,
        sources       => [
            Constant(
                name          => '100 per cent for service model annuities',
                defaultFormat => '%connz',
                rows          => $allTariffsByEndUser,
                cols          => $assetCustomerLevels,
                data          => [
                    map {
                        [ map { 1 } @{ $allTariffsByEndUser->{list} } ]
                    } @{ $assetCustomerLevels->{list} }
                ]
            ),
            Constant(
                name          => 'Zero for operating expenditure',
                defaultFormat => '%connz',
                rows          => $allTariffsByEndUser,
                cols          => $operatingLevels,
                data          => [
                    map {
                        [ map { 0 } @{ $allTariffsByEndUser->{list} } ]
                    } @{ $operatingLevels->{list} }
                ]
            ),
            $replacementShare
        ]
    ) if $model->{noReplacement} && $model->{noReplacement} =~ /hybrid/i;

    Columnset(
        name     => 'Financial and general assumptions',
        number   => 1010,
        appendTo => $model->{inputTables},
        dataset  => $model->{dataset},
        lines    => <<'EOT',
Sources: financial assumptions; calendar; network model.
These financial assumptions determine the annuity rate applied to convert the asset values of the network model into an annual charge.
EOT
        columns => [
            @{ $annuityRate->{arguments} }{qw(IV1 IV2)},
            $proportionChargeable,
            $powerFactorInModel,
            $daysInYear,
            grep { $_ && $_ != $daysInYear && ( ( ref $_ ) =~ /Dataset/ ); } (
                ref $daysBefore eq 'ARRAY' ? @$daysBefore : $daysBefore,
                $daysAfter,
            ),
        ]
    );

    unshift @{ $model->{networkModel} }, $annuityRate;

    my (
        $serviceModelAssets,             $serviceModelCostByCustomer,
        $serviceModelAssetsPerAnnualMwh, $serviceModelCostPerAnnualMwh,
      )
      = $model->serviceModels( $daysInYear, $modelLife, $annuityRate,
        $proportionChargeable, $allTariffsByEndUser, $componentMap );

    my (
        $volumeData,  $volumesAdjusted, $volumesByEndUser,
        $unitsInYear, $unitsByEndUser
      )
      = $model->volumes(
        $allTariffsByEndUser, $allEndUsers, $nonExcludedComponents,
        $componentMap,        $unitsLossAdjustment
      );

    my (
        $volumeDataAfter, $volumesAdjustedAfter, $revenueBefore,
        $revenuesBefore,  $tariffsBefore,        $unitsInYearAfter
    );

    if ( $model->{pcd} ) {

        pop @{ $model->{inputTables} };
        pop @{ $model->{volumeData} };

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
            lines =>
              [ 'Source: separate price control disaggregation model.', ],
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

        push @{ $model->{volumeData} }, $model->{pcd}{discount} = SumProduct(
            name => 'Discount for each tariff (except for fixed charges)',
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

        push @{ $model->{volumeData} }, $model->{pcd}{discountFixed} = Stack(
            name          => 'Discount for each tariff  for fixed charges only',
            defaultFormat => '%copynz',
            rows          => $model->{pcd}{discount}{rows},
            cols          => $model->{pcd}{discount}{cols},
            sources       => [
                Constant(
                    name =>
                      '100 per cent discount for generators on LDNO networks',
                    defaultFormat => '%connz',
                    rows          => $ldnoGenerators,
                    data => [ [ map { 1 } @{ $ldnoGenerators->{list} } ] ]
                ),
                $model->{pcd}{discount}
            ]
        );

        if ( $model->{portfolio} && $model->{portfolio} > 4 ) {

# take from EDCM
# supplement table 1037 with table 1181, or replace everything with a new table 1038 (in which separate discounts can be shown for demand, generation credits and generation fixed charges, at each level)

            die 'Not implemented yet';

        }

        ( $model->{pcd}{volumeData} ) = $model->volumes(
            $model->{pcd}{allTariffsByEndUser}, $allEndUsers,
            $nonExcludedComponents,             $componentMap
        );

        pop @{ $model->{volumeData} };

        if ( $model->{inYear} ) {

            # require CDCM::InYearAdjust;
            if ( $model->{inYear} =~ /after/i ) {
                $model->inYear_inPcdAdjust_after(
                    $nonExcludedComponents, $allEndUsers,
                    $componentMap,          \$revenueBefore,
                    \$unitsInYearAfter,     \$volumeDataAfter,
                    \$volumesAdjustedAfter,
                );
            }
            else {
                $model->inYearAdjust(
                    $nonExcludedComponents, $model->{pcd}{volumeData},
                    $allEndUsers,           $componentMap,
                    $daysAfter,             $daysBefore,
                    $daysInYear,            \$revenueBefore,
                    \$revenuesBefore,       \$tariffsBefore,
                    \$unitsInYearAfter,     \$volumeDataAfter,
                    \$volumesAdjustedAfter,
                );
            }
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

        $volumesAdjusted = $volumesByEndUser = $volumeData = {
            map {
                $_ => GroupBy(
                    name => SpreadsheetModel::Object::_shortName(
                        $intermediate{$_}{name}
                    ),
                    rows   => $allEndUsers,
                    source => $intermediate{$_}
                );
            } @$nonExcludedComponents
        };

        push @{ $model->{volumeData} },
          Columnset(
            name    => 'Equivalent volume for each end user',
            columns => [ @{$volumesAdjusted}{@$nonExcludedComponents} ]
          );

        $unitsInYear = $unitsByEndUser = Arithmetic(
            noCopy     => 1,
            name       => 'All units (MWh)',
            arithmetic => '='
              . join( '+', map { "IV$_" } 1 .. $model->{maxUnitRates} ),
            arguments => {
                map { ( "IV$_" => $volumesByEndUser->{"Unit rate $_ p/kWh"} ) }
                  1 .. $model->{maxUnitRates}
            },
            defaultFormat => '0softnz'
        );

    }

    elsif ( $model->{inYear} ) {

        # require CDCM::InYearAdjust;
        $model->inYearAdjust(
            $nonExcludedComponents, $volumeData,        $allEndUsers,
            $componentMap,          $daysAfter,         $daysBefore,
            $daysInYear,            \$revenueBefore,    \$revenuesBefore,
            \$tariffsBefore,        \$unitsInYearAfter, \$volumeDataAfter,
        );
        die 'inYearAdjust has created $model->{pcd}' if $model->{pcd};
    }

    my $unitsTariffsByEndUser = $model->{pcd} ? $unitsEndUsers : Labelset(
        name   => 'Units tariffs by end user',
        groups => $unitsEndUsers->{list}
    );

    my $demandTariffsByEndUser = $model->{pcd} ? $demandEndUsers : Labelset(
        name   => 'Demand tariffs by end user',
        groups => $demandEndUsers->{list}
    );

    my $capacityTariffsByEndUser = Labelset(
        name => 'Tariffs with capacity charges',
        ( $model->{pcd} ? 'list' : 'groups' ) => [
            grep { $componentMap->{$_}{'Capacity charge p/kVA/day'} }
              @{ $demandEndUsers->{list} }
        ]
    );

    my $standingForFixedTariffsByEndUser =
      $model->{pcd} ? $standingForFixedEndUsers : Labelset(
        name => 'Tariffs with fixed charges based on standing charges factors',
        groups => $standingForFixedEndUsers->{list}
      );

    my $generationCapacityTariffsByEndUser =
      $model->{pcd} ? $generationCapacityEndUsers : Labelset(
        name   => 'Generation capacity tariffs by end user',
        groups => $generationCapacityEndUsers->{list}
      );

    my $generationUnitsTariffsByEndUser =
      $model->{pcd} ? $generationUnitsEndUsers : Labelset(
        name   => 'Generation unit tariffs by end user',
        groups => $generationUnitsEndUsers->{list}
      );

    my ( $pseudoLoadCoefficientsAgainstSystemPeak, $pseudoLoadCoefficients );
    if ( $model->{maxUnitRates} && $model->{maxUnitRates} > 1 ) {
        if ( my $todmethod = $model->{timeOfDay} ) {
            $pseudoLoadCoefficients = $model->$todmethod(
                $drmExitLevels, $componentMap,     $allEndUsers,
                $daysInYear,    $loadCoefficients, $volumesByEndUser,
                $unitsByEndUser
            );
        }
        else {
            (
                $pseudoLoadCoefficientsAgainstSystemPeak,
                $pseudoLoadCoefficients
              )
              = $model->timeOfDay(
                $drmExitLevels, $componentMap,     $allEndUsers,
                $daysInYear,    $loadCoefficients, $volumesByEndUser,
                $unitsByEndUser
              );
        }
    }

    my ( $forecastSml, $simultaneousMaximumLoadUnits,
        $simultaneousMaximumLoadCapacity )
      = $model->networkUse(
        $networkLevels,          $drmLevels,
        $drmExitLevels,          $operatingLevels,
        $operatingDrmExitLevels, $unitsInYear,
        $loadCoefficients,       $pseudoLoadCoefficients,
        $daysInYear,             $lineLossFactors,
        $allTariffsByEndUser,    $componentMap,
        $routeingFactors,        $volumesAdjusted,
        $fFactors,               $generationCapacityTariffsByEndUser,
      );

    my (
        $standingFactors,  $forecastSmlAdjusted,
        $forecastAmlUnits, $diversityAllowancesAdjusted
      )
      = $model->diversity(
        $demandEndUsers,      $demandTariffsByEndUser,
        $unitsInYear,         $loadFactors,
        $daysInYear,          $lineLossFactors,
        $diversityAllowances, $componentMap,
        $volumesAdjusted,     $powerFactorInModel,
        $forecastSml,         $drmExitLevels,
        $coreExitLevels,      $rerouteing13211
      );

    my ( $siteSpecificSoleUseAssets, $siteSpecificReplacement );
    ( $siteSpecificSoleUseAssets, $siteSpecificReplacement ) =
      $model->siteSpecificSoleUse( $assetDrmLevels, $modelLife, $annuityRate,
        $proportionChargeable )
      if $model->{ehv} && $model->{ehv} =~ /s/i;

    my (
        $operatingCostToSml,       $operatingCostByCustomer,
        $operatingCostByAnnualMwh, $siteSpecificOperatingCost,
      )
      = $model->operating(
        $assetLevels,           $assetDrmLevels,
        $drmExitLevels,         $operatingLevels,
        $operatingDrmLevels,    $operatingDrmExitLevels,
        $assetCustomerLevels,   $operatingCustomerLevels,
        $forecastSmlAdjusted,   $allTariffsByEndUser,
        $unitsInYear,           $loadFactors,
        $daysInYear,            $lineLossFactors,
        $diversityAllowances,   $componentMap,
        $volumesAdjusted,       $modelGrossAssetsByLevel,
        $networkModelCostToSml, $modelSml,
        $serviceModelAssets,    $serviceModelAssetsPerAnnualMwh,
        $siteSpecificSoleUseAssets,
      );

    push @{ $model->{yardsticks} }, my $costToSml = Stack
      name => 'Unit cost at each level, £/kW/year'
      . ' (relative to system simultaneous maximum load)',
      cols    => $chargingDrmExitLevels,
      sources => [ $modelCostToSml, $operatingCostToSml ];

    my ( $yardstickCapacityRates, $paygUnitYardstick, $paygUnitRates ) =
      $model->yardsticks(
        $drmExitLevels,
        $chargingDrmExitLevels,
        $unitsTariffsByEndUser,
        $generationCapacityTariffsByEndUser,
        $allTariffsByEndUser,
        $loadCoefficients,
        $pseudoLoadCoefficients,
        $costToSml,
        $fFactors,
        $lineLossFactors,
        $allLevelsProportionCoveredByContributions,
        $daysInYear
      );

=head Development note

$yardstickCapacityComponents is available as $yardstickCapacityRates->{source}
$yardstickUnitsComponents is available as $paygUnitYardstick->{source}

=cut

    my ( $capacityCharges, $unauthorisedDemandCharges, $capacityUser,
        $unitYardstick, @unitRates )
      = $model->standingCharges(
        $standingFactors,                           $drmLevels,
        $drmExitLevels,                             $operatingDrmExitLevels,
        $chargingDrmExitLevels,                     $demandEndUsers,
        $demandTariffsByEndUser,                    $standingForFixedEndUsers,
        $standingForFixedTariffsByEndUser,          $loadFactors,
        $unitsByEndUser,                            $volumesByEndUser,
        $unitsInYear,                               $volumesAdjusted,
        $modelCostToAml,                            $modelCostToSml,
        $operatingCostToSml,                        $costToSml,
        $diversityAllowancesAdjusted,               $lineLossFactors,
        $allLevelsProportionCoveredByContributions, $powerFactorInModel,
        $daysInYear,                                $forecastAmlUnits,
        $componentMap,                              $paygUnitYardstick,
        @$paygUnitRates
      );

    my $sourceMap = {
        'Unit rate 1 p/kWh' => {
            'Standard 1 kWh'         => [ $unitRates[0] ],
            'Standard yardstick kWh' => [$unitYardstick],
            'PAYG 1 kWh & customer'  => [
                $operatingCostByAnnualMwh, $serviceModelCostPerAnnualMwh,
                $paygUnitRates->[0],
            ],
            'PAYG 1 kWh'                    => [ $paygUnitRates->[0] ],
            'PAYG yardstick kWh & customer' => [
                $operatingCostByAnnualMwh, $serviceModelCostPerAnnualMwh,
                $paygUnitYardstick,
            ],
            'PAYG yardstick kWh' => [$paygUnitYardstick],
        },
        (
            map {
                (
                    "Unit rate $_ p/kWh" => {
                        "Standard $_ kWh"        => [ $unitRates[ $_ - 1 ] ],
                        "PAYG $_ kWh & customer" => [
                            $operatingCostByAnnualMwh,
                            $serviceModelCostPerAnnualMwh,
                            $paygUnitRates->[ $_ - 1 ],
                        ],
                        "PAYG $_ kWh" => [ $paygUnitRates->[ $_ - 1 ] ],
                    }
                  )
            } 2 .. $model->{maxUnitRates}
        ),
        $fFactors
        ? ( 'Generation capacity rate p/kW/day' =>
              { 'PAYG yardstick kW' => [$yardstickCapacityRates] } )
        : (),
        'Capacity charge p/kVA/day' => { 'Capacity' => [$capacityCharges] },
        (
            $model->{unauth} && $model->{unauth} =~ /day/
            ? 'Exceeded capacity charge p/kVA/day'
            : 'Unauthorised demand charge p/kVAh'
          ) => { 'Capacity' => [$unauthorisedDemandCharges] },
        'Fixed charge p/MPAN/day' => {
            'Fixed from network'            => [$capacityUser],
            'Fixed from network & customer' => [
                $operatingCostByCustomer, $serviceModelCostByCustomer,
                $capacityUser,
            ],
            'Customer' =>
              [ $operatingCostByCustomer, $serviceModelCostByCustomer ],
        },
        'Reactive power charge p/kVArh' => {
            'Standard kVArh' => [],
            'PAYG kVArh'     => [],
        },
    };

    my ($tariffsExMatching) = $model->aggregation(
        $componentMap,          $allTariffsByEndUser, $chargingLevels,
        $nonExcludedComponents, $allComponents,       $sourceMap,
    );

    my $componentLabelset = {};

    $model->reactive(
        $drmExitLevels,                             $chargingDrmExitLevels,
        $chargingLevels,                            $componentMap,
        $allTariffsByEndUser,                       $unitYardstick,
        $costToSml,                                 $loadCoefficients,
        $lineLossFactorsToGsp,                      $lineLossFactorsNetwork,
        $allLevelsProportionCoveredByContributions, $daysInYear,
        $powerFactorInModel,                        $tariffsExMatching,
        $componentLabelset,                         $sourceMap
    );

    push @{ $model->{preliminaryAggregation} }, Columnset
      name    => 'Summary of charges before revenue matching',
      columns => [ @{$tariffsExMatching}{@$allComponents} ];

    my ( $revenueShortfall, $totalRevenuesSoFar, $revenuesSoFar,
        $allowedRevenue, $revenueFromElsewhere, $totalSiteSpecificReplacement, )
      = $model->revenueShortfall(
        $allTariffsByEndUser,
        $nonExcludedComponents,
        $daysAfter,
        ( $model->{pcd} ? $volumesAdjustedAfter : $volumeDataAfter )
          || $volumesAdjusted,
        $revenueBefore,
        $tariffsExMatching,
        $siteSpecificOperatingCost,
        $siteSpecificReplacement,
      );

    my ( $totalRevenuesFromMatching, $siteSpecificCharges, @matchingTables ) =
      $model->{scaler} && $model->{scaler} =~ /DCP123/i
      ? $model->matching2012(
        $revenueShortfall,
        $componentMap,
        $allTariffsByEndUser,
        $demandTariffsByEndUser,
        $allEndUsers,
        $chargingLevels,
        $nonExcludedComponents,
        $allComponents,
        $daysAfter,
        ( $model->{pcd} ? $volumesAdjustedAfter : $volumeDataAfter )
          || $volumesAdjusted,
        $loadCoefficients,
        $tariffsExMatching,
        $daysInYear,
        $model->{pcd} ? $volumesAdjusted : $volumeData,
      )
      : $model->matching(
        $revenueShortfall,
        $revenuesSoFar,
        $totalRevenuesSoFar,
        $totalSiteSpecificReplacement,
        $componentMap,
        $allTariffsByEndUser,
        $demandTariffsByEndUser,
        $allEndUsers,
        $chargingLevels,
        $nonExcludedComponents,
        $allComponents,
        $daysAfter,
        ( $model->{pcd} ? $volumesAdjustedAfter : $volumeDataAfter )
          || $volumesAdjusted,
        $revenueBefore,
        $loadCoefficients,
        $lineLossFactorsToGsp,
        $tariffsExMatching,
        $unitsInYearAfter || $unitsInYear,
        $generationCapacityTariffsByEndUser,
        $fFactors,
        $annuityRate,
        $modelLife,
        $costToSml,
        $modelCostToSml,
        $operatingCostToSml,
        $routeingFactors,
        $replacementShare,
        $siteSpecificOperatingCost,
        $siteSpecificReplacement,
        $simultaneousMaximumLoadUnits,
        $simultaneousMaximumLoadCapacity,
        @$pseudoLoadCoefficientsAgainstSystemPeak
      );

    foreach my $table (@matchingTables) {
        foreach my $comp ( keys %$table ) {
            foreach my $src ( keys %{ $sourceMap->{$comp} } ) {
                push @{ $sourceMap->{$comp}{$src} }, $table->{$comp};
            }
        }
    }

    push @{ $model->{tariffSummary} },
      Arithmetic(
        name          => 'Charging rate for site-specific sole use assets',
        defaultFormat => '%softnz',
        arithmetic    => '=IF(IV2,IV1/IV3,"")',
        arguments     => {
            IV2 => $siteSpecificSoleUseAssets,
            IV1 => $siteSpecificCharges,
            IV3 => $siteSpecificSoleUseAssets
        }
      ) if $siteSpecificCharges;

    my ( $tariffTable, $tariffsBeforeRounding, $lossesAdjTable ) =
      $model->roundingAndFinishing(
        $allComponents,
        $tariffsExMatching,
        $componentLabelset,
        $allTariffs,
        $componentMap,
        $sourceMap,
        $daysAfter,
        $nonExcludedComponents,
        $unitsLossAdjustment,
        ( $model->{pcd} ? $volumesAdjustedAfter : $volumeDataAfter )
          || $volumeData,
        $allTariffsByEndUser,
        $totalRevenuesSoFar,
        $totalRevenuesFromMatching,
        $allowedRevenue,
        $revenueBefore,
        $revenueFromElsewhere,
        $siteSpecificCharges,
        $chargingLevels,
        @matchingTables,
      );

    $model->makeMatrixClosure(
        $lossesAdjTable,
        $allComponents,
        $componentLabelset,
        $allTariffs,
        $componentMap,
        $sourceMap,
        $model->{matrices} =~ /partyear/
        ? ( $daysAfter, $volumeDataAfter, $unitsInYearAfter )
        : ( $daysInYear, $volumeData, $unitsInYear ),
        $chargingLevels,
        @matchingTables
    ) if $model->{matrices};

    $model->displayWholeYearTarget( $allComponents, $daysInYear, $volumeData,
        $tariffsBeforeRounding, $allowedRevenue, $revenueFromElsewhere,
        $siteSpecificCharges, )
      if $model->{inYear} && $model->{inYear} =~ /target/;

    if ( $model->{pcd} ) {

        $allTariffs = $allTariffsByEndUser = $model->{pcd}{allTariffsByEndUser};

        $volumesAdjusted = $volumeData = $model->{pcd}{volumeData};

        $unitsInYear = Arithmetic(
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

        $tariffTable = {
            map {
                $_ => Arithmetic(
                    name => SpreadsheetModel::Object::_shortName(
                        $tariffTable->{$_}{name}
                    ),
                    defaultFormat => $tariffTable->{$_}{defaultFormat},
                    arithmetic    => $model->{model100} ? '=IV2*(1-IV1)'
                    : ( '=ROUND(IV2*(1-IV1),' . ( /kWh|kVArh/ ? 3 : 2 ) . ')' ),
                    rows      => $allTariffs,
                    cols      => $tariffTable->{$_}{cols},
                    arguments => {
                        IV2 => $tariffTable->{$_},
                        IV1 => /fix/i ? $model->{pcd}{discountFixed}
                        : $model->{pcd}{discount}
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

        my $atwTariffSummary = pop @{ $model->{tariffSummary} };
        push @{ $model->{roundingResults} }, $atwTariffSummary;

        unshift @{ $model->{tariffSummary} }, Columnset(
            name => 'Tariffs',
            $model->{noLLFCs}
            ? ()
            : (
                dataset               => $model->{dataset},
                doNotCopyInputColumns => 1,
                number                => 3701,
            ),    # hacks to get the LLFCs to copy
            columns => [
                $model->{noLLFCs} ? () : (
                    Dataset(
                        rows          => $allTariffsReordered,
                        defaultFormat => 'texthard',
                        data => [ map { '' } @{ $allTariffs->{list} } ],
                        name => 'Open LLFCs',
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
            ]
        );

    }

    if ( $model->{summary} ) {

        push @{ $model->{optionLines} }, ' ';

        push @{ $model->{overallSummary} },
          Columnset(
            name => 'Workbook build options and main parameters',
            1 ? () : ( singleRowName => 'Parameter value' ),
            lines   => $model->{optionLines},
            columns => $model->{summaryColumns},
          );

        my $revenuesByTariff;

        if ( $model->{summary} =~ /partyear/i ) {
            $revenuesByTariff =
              $model->summaryOfRevenues( $tariffTable, $volumeDataAfter,
                $daysAfter, $nonExcludedComponents, $componentMap, $allTariffs,
                $unitsInYearAfter, )
              if $model->{summary} !~ /disclosure/i;
        }
        elsif ( $model->{summary} !~ /hybrid/i ) {
            $revenuesByTariff =
              $model->summaryOfRevenues( $tariffTable, $volumeData, $daysInYear,
                $nonExcludedComponents, $componentMap, $allTariffs,
                $unitsInYear, );
        }

        if ( $model->{summary} =~ /hybrid|disclosure/i ) {

            # require CDCM::InYearSummaries;
            my @revenuesByTariffHybrid = $model->summaryOfRevenuesHybrid(
                $allTariffs,       $nonExcludedComponents,
                $componentMap,     $volumeDataAfter,
                $daysBefore,       $daysAfter,
                $unitsInYearAfter, $tariffTable,
                $volumeData,       $daysInYear,
                $unitsInYear,
            );
            $revenuesByTariff =
              $revenuesByTariffHybrid[ $#revenuesByTariffHybrid - 1 ]
              if $model->{summary} =~ /partyear/i;
            $revenuesByTariff =
              $revenuesByTariffHybrid[$#revenuesByTariffHybrid]
              if $model->{summary} =~ /hybrid/i;
        }

        $model->consultationSummary(
            $revenuesByTariff,
            $tariffTable,
            $model->{summary} =~ /partyear/i ? $volumeDataAfter : $volumeData,
            $model->{summary} =~ /partyear/i ? $daysAfter       : $daysInYear,
            $nonExcludedComponents,
            $componentMap,
            $allTariffs,
            $allTariffsByEndUser,
            $model->{summary} =~ /partyear/i ? $unitsInYearAfter : $unitsInYear,
            $unitsLossAdjustment,
            $tariffsBefore,
            $revenuesBefore,
            $unitsInYear,
        ) if $model->{summary} =~ /consul/i;

    }

    $model;

}

1;
