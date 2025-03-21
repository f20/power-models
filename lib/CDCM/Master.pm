﻿package CDCM;

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

use CDCM::Aggregation;
use CDCM::AML;
use CDCM::Contributions;
use CDCM::Loads;
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
use CDCM::Yardsticks;

sub requiredModulesForRuleset {

    my ( $class, $ruleset ) = @_;

    my $todmodule = ucfirst( $ruleset->{timeOfDay} || 'timeOfDay' );
    die "Time of day module CDCM::$todmodule is unsafe"
      unless $todmodule =~ /^[a-zA-Z0-9_]+$/s;

    "CDCM::$todmodule",

      $ruleset->{tariffSpec} ? 'CDCM::TariffSpec' : 'CDCM::TariffList',

      $ruleset->{tariffGrouping} ? 'CDCM::Grouping' : (),

      $ruleset->{pcd} ? 'CDCM::Discounts' : (),

      $ruleset->{inYear} ? qw(CDCM::InYearAdjust CDCM::InYearSummaries)
      : $ruleset->{addVolumes}
      && $ruleset->{addVolumes} =~ /matching/i ? 'CDCM::InYearAdjust'
      : (),

      $ruleset->{targetRevenue}
      && $ruleset->{targetRevenue} =~ /dcp132|2012/i
      ? 'CDCM::Table1001_2012'

      : $ruleset->{targetRevenue}
      && $ruleset->{targetRevenue} =~ /dcp249|dcp273|2016/i
      ? 'CDCM::Table1001_2016'

      : $ruleset->{targetRevenue}
      && $ruleset->{targetRevenue} =~ /dcp334|2019/i
      ? 'CDCM::Table1001_2019'

      : $ruleset->{targetRevenue}
      && $ruleset->{targetRevenue} =~ /dcp421|2024/i
      ? 'CDCM::Table1001_2024'

      : (),

       !$ruleset->{scaler}                  ? ()
      : $ruleset->{scaler} =~ /tcrd/i       ? 'CDCM::DomesticTcr'
      : $ruleset->{scaler} =~ /tcr|dcp361/i ? 'CDCM::Matching361'
      : $ruleset->{scaler} =~ /ppuflex/i    ? 'CDCM::MatchingFlex'
      : $ruleset->{scaler} =~ /dcp123/i     ? 'CDCM::Matching123'
      : 'CDCM::Matching',

      !$ruleset->{summary}
      ? ()
      : $ruleset->{summary} =~ /stat(?:istic)?s/i
      ? 'CDCM::Statistics' . ( $ruleset->{summary} =~ /(1203|1204)/ ? $1 : '' )
      : $ruleset->{summary} =~ /consul/i ? 'CDCM::SummaryDeprecated'
      : (),

      $ruleset->{matrices}
      && $ruleset->{matrices} =~ /charts/i ? 'CDCM::MatrixCharts' : (),

      $ruleset->{tcdb} ? qw(CDCM::TCDB) : (),

      $ruleset->{checksums} ? qw(SpreadsheetModel::Checksum) : (),

      $ruleset->{timebandDetails} ? qw(CDCM::Timebands) : (),

      $ruleset->{unroundedTariffAnalysis}
      ? (
        qw(CDCM::TariffAnalysis),
        $ruleset->{unroundedTariffAnalysis} =~ /modelg/i
        ? qw(CDCM::ModelG)
        : ()
      )
      : (),

      $ruleset->{embeddedModelG}
      ? eval { CDCM->requiredModulesForRuleset( $ruleset->{embeddedModelG} ); }
      : (),

      $ruleset->{embeddedModelM}
      ? eval {
        require ModelM::Master;
        ModelM->requiredModulesForRuleset( $ruleset->{embeddedModelM} );
      }
      : (),

      ;

}

sub new {

    my $class = shift;
    my $model = {@_};
    bless $model, $class;

    die 'This system will not build an orange '
      . 'CDCM model without a suitable disclaimer.' . "\n--"
      if $model->{colour}
      && $model->{colour} =~ /orange/i
      && !($model->{extraNotice}
        && length( $model->{extraNotice} ) > 299
        && $model->{extraNotice} =~ /DCUSA/ );

    $model->{inputTables} ||= [];
    $model->{table1000} = Dataset(
        number             => 1000,
        dataset            => $model->{dataset},
        name               => 'Company, charging year, data version',
        cols               => Labelset( list => [qw(Company Year Version)] ),
        defaultFormat      => 'puretexthard',
        data               => [ 'no company', 'no year', 'no data version' ],
        usePlaceholderData => 1,
        forwardLinks       => {},
        appendTo           => $model->{inputTables},
    );

    $model->{edcmTables} = [
        [
            Constant(
                name => 'Generation O&M charging rate (£/kW/year)',
                data => [0.2],
            ),
        ],
        Stack(
            name =>
              'EDCM input data ⇒1100. Company, charging year, data version',
            sources => [ $model->{table1000} ],
        ),
      ]
      if $model->{edcmTables};

    $model->{timebands} = 3 unless $model->{timebands};
    $model->{timebands} = 10 if $model->{timebands} > 10;
    $model->{drm}       = 'top500gsp' unless $model->{drm};

   # Keep CDCM::DataPreprocess out of the scope of revision number construction.

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

    if ( ref $model->{sharingObjectRef} ) {
        $model->{sharedData} = ${ $model->{sharingObjectRef} }
          if UNIVERSAL::can( ${ $model->{sharingObjectRef} }, 'addStats' );
        ${ $model->{sharingObjectRef} }->registerCdcmModel($model)
          if UNIVERSAL::can( ${ $model->{sharingObjectRef} },
            'registerCdcmModel' );
    }

    $model->timebandDetails if $model->{timebandDetails};

    my ( $daysInYear, $daysBefore, $daysAfter, $modelLife, $annuityRate,
        $powerFactorInModel, $drmLevels, $drmExitLevels, )
      = $model->setUp;

    my ( $allTariffs, $allTariffsByEndUser, $allEndUsers, $allComponents,
        $nonExcludedComponents, $componentMap )
      = $model->tariffs;

    ( $allEndUsers, $allTariffsByEndUser, $allTariffs ) =
      $model->setUpGrouping( $componentMap, $allEndUsers,
        $allTariffsByEndUser, $allTariffs )
      if $model->{tariffGrouping};

    $model->{ldnoWord} = $model->{portfolio}
      && $model->{portfolio} =~ /qno/i ? 'QNO' : 'LDNO';

    ( $allEndUsers, $allTariffsByEndUser, $allTariffs ) =
      $model->pcdSetUp( $allEndUsers, $allTariffsByEndUser, $allTariffs )
      if $model->{pcd};

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
                error_message =>
                  'The proportion of load going through 132kV/HV'
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
                arithmetic    => '=1-A1',
                arguments     => { A1 => $rerouteing13211 }
            ),
            Arithmetic(
                name          => 'Proportion going through EHV',
                cols          => Labelset( list => [ $drmLevels->{list}[2] ] ),
                rows          => Labelset( list => [ $coreLevels->{list}[2] ] ),
                defaultFormat => '%softnz',
                arithmetic    => '=1-A1',
                arguments     => { A1 => $rerouteing13211 }
            ),
            Arithmetic(
                name          => 'Proportion going through EHV/HV',
                cols          => Labelset( list => [ $drmLevels->{list}[3] ] ),
                rows          => Labelset( list => [ $coreLevels->{list}[3] ] ),
                defaultFormat => '%softnz',
                arithmetic    => '=1-A1',
                arguments     => { A1 => $rerouteing13211 }
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
        );

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
      = $model->routeing( $allTariffsByEndUser, $coreLevels,
        $coreExitLevels, $drmLevels, $drmExitLevels, $rerouteingMatrix );

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
            @{ $annuityRate->{arguments} }{qw(A1 A2)},
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
        $volumeData,           $volumesAdjusted, $volumesByEndUser,
        $unitsInYear,          $unitsByEndUser,  $volumeDataAfter,
        $volumesAdjustedAfter, $revenueBefore,   $revenuesBefore,
        $tariffsBefore,        $unitsInYearAfter,
    );

    if ( $model->{pcd} ) {
        $model->pcdPreprocessedVolumes(
            $allEndUsers,      $componentMap,
            $daysAfter,        $daysBefore,
            $daysInYear,       $nonExcludedComponents,
            \$revenueBefore,   \$revenuesBefore,
            \$tariffsBefore,   \$unitsByEndUser,
            \$unitsInYear,     \$unitsInYearAfter,
            \$volumeData,      \$volumeDataAfter,
            \$volumesAdjusted, \$volumesAdjustedAfter,
            \$volumesByEndUser,
        );
    }
    else {

        (
            $volumeData,  $volumesAdjusted, $volumesByEndUser,
            $unitsInYear, $unitsByEndUser
          )
          = $model->volumes(
            $allTariffsByEndUser, $allEndUsers, $nonExcludedComponents,
            $componentMap,        $unitsLossAdjustment
          );

        if ( $model->{inYear} ) {
            if ( $model->{inYear} =~ /after/i ) {
                $model->inYearAdjustUsingAfter(
                    $nonExcludedComponents, $volumeData,
                    $allEndUsers,           $componentMap,
                    \$revenueBefore,        \$unitsInYearAfter,
                    \$volumeDataAfter,
                );
            }
            else {
                $model->inYearAdjustUsingBefore(
                    $nonExcludedComponents, $volumeData,
                    $allEndUsers,           $componentMap,
                    $daysAfter,             $daysBefore,
                    $daysInYear,            \$revenueBefore,
                    \$revenuesBefore,       \$tariffsBefore,
                    \$unitsInYearAfter,     \$volumeDataAfter,
                );
            }
        }
        elsif ($model->{addVolumes}
            && $model->{addVolumes} =~ /matching/i )
        {
            $model->inYearAdjustUsingAfter( $nonExcludedComponents,
                $volumeData, $allEndUsers,       $componentMap,
                undef,       \$unitsInYearAfter, \$volumeDataAfter, );
        }
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
        if ( my $timeOfDay = $model->{timeOfDay} ) {
            $pseudoLoadCoefficients = $model->$timeOfDay(
                $drmExitLevels, $componentMap,     $allEndUsers,
                $daysInYear,    $loadCoefficients, $volumesByEndUser,
                $unitsByEndUser
            );
        }
        else {
           # Legacy code which creates $pseudoLoadCoefficientsAgainstSystemPeak,
           # which is possibly used by some old revenue matching methods.
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

    $loadFactors = $model->impliedLoadFactors(
        $allEndUsers,              $demandEndUsers,
        $standingForFixedEndUsers, $componentMap,
        $volumesByEndUser,         $unitsByEndUser,
        $daysInYear,               $powerFactorInModel,
        $loadFactors,
    ) if $model->{impliedLoadFactors};

    my (
        $standingFactors,  $forecastSmlAdjusted,
        $forecastAmlUnits, $diversityAllowancesAdjusted
      )
      = $model->diversity(
        $demandEndUsers,                   $demandTariffsByEndUser,
        $standingForFixedTariffsByEndUser, $unitsInYear,
        $loadFactors,                      $daysInYear,
        $lineLossFactors,                  $diversityAllowances,
        $componentMap,                     $volumesAdjusted,
        $powerFactorInModel,               $forecastSml,
        $drmExitLevels,                    $coreExitLevels,
        $rerouteing13211
      );

    my ( $siteSpecificSoleUseAssets, $siteSpecificReplacement );
    ( $siteSpecificSoleUseAssets, $siteSpecificReplacement ) =
      $model->siteSpecificSoleUse( $assetDrmLevels, $modelLife,
        $annuityRate, $proportionChargeable )
      if $model->{ehv} && $model->{ehv} =~ /s/i;

    my (
        $operatingCostToSml,       $operatingCostByCustomer,
        $operatingCostByAnnualMwh, $siteSpecificOperatingCost,
      )
      = $model->operating(
        $assetLevels,                    $assetDrmLevels,
        $drmExitLevels,                  $operatingLevels,
        $operatingDrmLevels,             $operatingDrmExitLevels,
        $assetCustomerLevels,            $operatingCustomerLevels,
        $forecastSmlAdjusted,            $allTariffsByEndUser,
        $unitsInYear,                    $daysInYear,
        $lineLossFactors,                $diversityAllowances,
        $componentMap,                   $volumesAdjusted,
        $modelGrossAssetsByLevel,        $networkModelCostToSml,
        $modelSml,                       $serviceModelAssets,
        $serviceModelAssetsPerAnnualMwh, $siteSpecificSoleUseAssets,
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
        $componentLabelset,                         $sourceMap,
        $rerouteingMatrix,
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
         $model->{scaler}
      && $model->{scaler} =~ /tcr/i && $model->{scaler} !~ /tcrd/i
      ? $model->matching361(
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
      : $model->{scaler} && $model->{scaler} =~ /DCP123/i
      ? $model->matchingdcp123(
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
        @$pseudoLoadCoefficientsAgainstSystemPeak,
      )
      unless $model->{scaler}
      && $model->{scaler} eq 'none';

    $model->domesticTcr(
        $allComponents,
        $daysAfter,
        ( $model->{pcd} ? $volumesAdjustedAfter : $volumeDataAfter )
          || $volumesAdjusted,
        \@matchingTables,
      )
      if $model->{scaler}
      && $model->{scaler} =~ /tcrd/i;

    foreach my $table (@matchingTables) {
        foreach my $comp ( keys %$table ) {
            foreach my $src ( keys %{ $sourceMap->{$comp} } ) {
                push @{ $sourceMap->{$comp}{$src} }, $table->{$comp};
            }
        }
    }

    if ( $model->{unroundedTariffAnalysis} ) {
        my @utaTables = $model->unroundedTariffAnalysis(
            $allComponents, $allTariffsByEndUser, $componentLabelset,
            $daysAfter,     $tariffsExMatching,   @matchingTables,
        );
        return $model->modelG( $nonExcludedComponents, $daysAfter,
            $volumeData, $allEndUsers, @utaTables )
          if $model->{unroundedTariffAnalysis} =~ /modelg/i;
        push @{ $model->{utaTables} },
          grep { ref $_; } map { values %$_; } @utaTables;
    }

    push @{ $model->{tariffSummary} },
      Arithmetic(
        name          => 'Charging rate for site-specific sole use assets',
        defaultFormat => '%softnz',
        arithmetic    => '=IF(A2,A1/A3,"")',
        arguments     => {
            A2 => $siteSpecificSoleUseAssets,
            A1 => $siteSpecificCharges,
            A3 => $siteSpecificSoleUseAssets
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
        (
              $model->{pcd}
            ? $volumesAdjustedAfter
            : $volumeDataAfter
          )
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

    ${ $model->{sharingObjectRef} }
      ->tariffAnalysis( $model, $allTariffs, $allComponents,
        $componentLabelset, $tariffTable, $componentMap, $sourceMap, )
      if ref $model->{sharingObjectRef}
      && UNIVERSAL::can( ${ $model->{sharingObjectRef} }, 'tariffAnalysis' );

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

    $model->displayWholeYearTarget( $allComponents, $daysInYear,
        $volumeData,     $tariffsBeforeRounding,
        $allowedRevenue, $revenueFromElsewhere, $siteSpecificCharges, )
      if $model->{inYear} && $model->{inYear} =~ /target/;

    (
        $allTariffs, $allTariffsByEndUser, $volumeData, $unitsInYear,
        $tariffTable,
      )
      = $model->pcdApplyDiscounts(
        $allComponents,
        $model->{generationReactiveRouteingFix}
        ? $tariffsBeforeRounding
        : $tariffTable,
        $daysInYear,
        $componentMap,
      ) if $model->{pcd};

    (
        $allTariffs, $allTariffsByEndUser, $volumeData, $unitsInYear,
        $tariffTable,
      )
      = $model->degroupTariffs( $allComponents, $tariffTable )
      if $model->{tariffGrouping};

    my $allTariffsReordered  = $allTariffs;
    my $tariffTableReordered = $tariffTable;

    unless ( $model->{tariffOrder} ) {

        push @{ $model->{postPcdApplicationResults} },
          Columnset(
            name    => 'Tariffs',
            columns => [ @{$tariffTable}{@$allComponents} ],
          );

        my @allTariffNames =
          grep { !/Non-Domestic Aggregated \(Related MPAN\) Band [1-4]/; }
          map { $allTariffs->{list}[$_] } $allTariffs->indices;

        $allTariffsReordered = Labelset(
            name => 'All tariffs',
            list => [
                ( grep { !/(?:ID|LD|Q)NO/i; } @allTariffNames ),
                ( grep { /(?:ID|LD|Q)NO LV:/i; } @allTariffNames ),
                ( grep { /(?:ID|LD|Q)NO HV:/i; } @allTariffNames ),
                ( grep { /(?:ID|LD|Q)NO HVplus:/i; } @allTariffNames ),
                ( grep { /(?:ID|LD|Q)NO EHV:/i; } @allTariffNames ),
                ( grep { /(?:ID|LD|Q)NO 132kV\/EHV:/i; } @allTariffNames ),
                ( grep { /(?:ID|LD|Q)NO 132kV:/i; } @allTariffNames ),
                ( grep { /(?:ID|LD|Q)NO 0000:/i; } @allTariffNames ),
                (
                    grep {
                        /(?:ID|LD|Q)NO/i
                          && !
/(?:ID|LD|Q)NO (?:LV|HV|HVplus|EHV|132kV\/EHV|132kV|0000):/i
                    } @allTariffNames
                )
            ]
        );

        $tariffTableReordered = {
            map {
                my $component = $_;
                (
                    $component => Stack(
                        name          => $tariffTable->{$component}{name},
                        defaultFormat => (
                            map {
                                local $_ = $_;
                                s/soft/copy/ if defined $_;
                                $_;
                            } $tariffTable->{$component}{defaultFormat}
                        ),
                        rows       => $allTariffsReordered,
                        rowFormats => [
                            map {
                                $componentMap->{$_}{$component}
                                  ? undef
                                  : 'unavailable';
                            } @{ $allTariffsReordered->{list} }
                        ],
                        cols    => $tariffTable->{$component}{cols},
                        sources => [ $tariffTable->{$component} ],
                    )
                );
            } @$allComponents
        };

    }

    my @allTariffColumns = @{$tariffTableReordered}{@$allComponents};
    $model->{allTariffColumns} = \@allTariffColumns;

    unshift @{ $model->{tariffSummary} },
      $model->{tariffColumnset3701} = Columnset(
        $model->{noLLFCs}
        ? ( name => '' )
        : (
            name                  => 'Tariffs',
            dataset               => $model->{dataset},
            doNotCopyInputColumns => 1,
            number => 3701,    # needed as this is an input data table for LLFCs
        ),
        columns => [
            $model->{noLLFCs} ? () : (
                Dataset(
                    rows          => $allTariffsReordered,
                    defaultFormat => 'puretexthard',
                    data => [ map { '' } @{ $allTariffsReordered->{list} } ],
                    name => 'Open LLFCs',
                ),
                $model->{noPCs} ? () : Constant(
                    rows          => $allTariffsReordered,
                    defaultFormat => 'puretextcon',
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
                defaultFormat => 'puretexthard',
                data => [ map { '' } @{ $allTariffsReordered->{list} } ],
                name => 'Closed LLFCs',
            ),
            $model->{checksums} ? (
                map {
                    my $digits    = /([0-9])/                ? $1 : 6;
                    my $recursive = /table|recursive|model/i ? 1  : 0;
                    $model->{"checksum_${recursive}_$digits"} =
                      SpreadsheetModel::Checksum->new(
                        name      => $_,
                        recursive => $recursive,
                        digits    => $digits,
                        columns   => \@allTariffColumns,
                        factors   => [
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

    $model->{componentMapForTcdb} = $componentMap if $model->{tcdb};

    if ( $model->{summary} ) {

        if ( $model->{summary} =~ /consultation|headline/i ) {
            push @{ $model->{optionLines} },
              'The list of options above is not comprehensive',
              'This is just padding', 'This is just padding', ' ';
            push @{ $model->{overallSummary} },
              Columnset(
                name => $model->{model100}
                ? 'Workbook build options and main parameters'
                : 'Headline parameters',
                $model->{model100} ? ( lines => $model->{optionLines} )
                : (),
                columns => $model->{summaryColumns},
              );
        }

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

        $model->comparisonSummary(
            $revenuesByTariff,
            $tariffTable,
            $model->{summary} =~ /partyear/i ? $volumeDataAfter
            : $volumeData,
            $model->{summary} =~ /partyear/i ? $daysAfter : $daysInYear,
            $nonExcludedComponents,
            $componentMap,
            $allTariffs,
            $allTariffsByEndUser,
            $model->{summary} =~ /partyear/i ? $unitsInYearAfter
            : $unitsInYear,
            $unitsLossAdjustment,
            $tariffsBefore,
            $revenuesBefore,
            $unitsInYear,
        ) if $model->{summary} =~ /1201/;

        if ( $model->{summary} =~ /stat(?:istic)?s/i ) {
            my $statsMethod = 'makeStatisticsTables'
              . ( $model->{summary} =~ /(1203|1204)/ ? $1 : '' );
            $model->$statsMethod(
                $tariffTableReordered,
                $daysInYear,
                $nonExcludedComponents,
                $componentMap,
                $model->{summary} =~ /partyear/i
                ? $volumeDataAfter
                : $volumeData,
            );
        }
        elsif ( $model->{summary} =~ /consul/i ) {
            $model->consultationSummaryDeprecated(
                $revenuesByTariff,
                $tariffTable,
                $model->{summary} =~ /partyear/i ? $volumeDataAfter
                : $volumeData,
                $model->{summary} =~ /partyear/i ? $daysAfter
                : $daysInYear,
                $nonExcludedComponents,
                $componentMap,
                $allTariffs,
                $allTariffsByEndUser,
                $model->{summary} =~ /partyear/i ? $unitsInYearAfter
                : $unitsInYear,
                $unitsLossAdjustment,
                $tariffsBefore,
                $revenuesBefore,
                $unitsInYear,
            );
        }

    }

    $model->{sharedData}
      ->addStats( 'DNO-wide aggregates', $model, $totalRevenuesFromMatching )
      if $model->{sharedData};

    $model;

}

1;
