package EDCM;

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

=head Table numbers used in this file

1111 (shared)
1168 (shared)
1192

=cut

=head Table number availability

110x Inputs Assets Scaling Sheets
111x Inputs Master
112x Assets
113x Assets
114x Scaling
115x 
116x Ldno Master
117x
118x
119x Master

=cut

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Shortcuts ':all';
use EDCM::Ldno;
use EDCM::Inputs;
use EDCM::Assets;
use EDCM::Locations;
use EDCM::Charges;
use EDCM::Scaling;
use EDCM::Summary;
use EDCM::ProcessData;
use EDCM::Sheets;

sub new {

    my $class = shift;
    my $model = {@_};
    $model->{inputTables} = [];
    bless $model, $class;

    if ( $model->{dataset} ) {
        $model->processData;
        if ( $model->{randomise} ) {
            $model->randomiseCut if $model->{randomise} =~ /cut/i;
            if ( $model->{randomise} =~ /aggressive/i ) {
                $model->randomiseAggressive;
            }
            else {
                $model->randomiseNormal;
            }
        }
    }

    if ( $model->{ldnoRev} && $model->{ldnoRev} =~ /only/i ) {
        $model->{daysInYear} = Dataset(
            name          => 'Days in year',
            defaultFormat => '0hard',
            data          => [365],
            dataset       => $model->{dataset},
            appendTo      => $model->{inputTables},
            number        => 1111,
            validation    =>
              { validate => 'decimal', criteria => '>=', value => 0, }
        );
        $model->{ldnoRevTables} = [ $model->ldnoRev ];
        return $model;
    }

    print "$model->{datafile} $model->{method}: "
      . "$model->{numLocations} locations, $model->{numTariffs} tariffs\n"
      if defined $model->{numLocations} && defined $model->{numTariffs};

    die unless $model->{vedcm} == 53;

    my (
        $edcmScope,      $daysInYear,            $direct,
        $indirect,       $indirectProp,          $rates,
        $ehvIntensity,   $transmissionExitTotal, $systemPeakLoad,
        $allowedRevenue, $powerFactorInModel,    $ehvAssetLevelset,
      )
      = $model->generalInputs;

    $model->{ldnoRevTables} = [ $model->ldnoRev() ] if $model->{ldnoRev};

    my $hoursInRed = Dataset(
        name          => 'Annual hours in super red',
        defaultFormat => '0.0hardnz',
        data          => [90],
        dataset       => $model->{dataset},
    );

    my $hoursInRedGeneration = 1 ? $hoursInRed : Arithmetic(
        name          => 'Annual hours in which generation credits are payable',
        arithmetic    => '=24*IV1',
        defaultFormat => '0.0hardnz',
        arguments     => { IV1 => $daysInYear, }
    );

    my (
        $llfcs,                   $tariffs,
        $tariffDorG,              $customerClass,
        $tariffDaysInYearNot,     $tariffHoursInRedNot,
        $totalAgreedCapacity,     $nonChargeableCapacity,
        $activeCoincidence,       $reactiveCoincidence,
        $activeUnits,             $creditableCapacity,
        $tariffSoleUseMeav,       $tariffSoleUsePropex,
        $tariffProportionPre2005, $tariffPoint,
        $tariffGroup,             $tariffNetworkSupportFactor,
        $tariffCategory,          $useProportions,
        $importForGenerator,      $importForLdno,
        $exceededCapacity,        $previousIncome,
      )
      = $model->tariffInputs($ehvAssetLevelset);

    my $totalAgreedCapacity953 = $totalAgreedCapacity;
    my $chargeableCapacity     = Arithmetic(
        name          => 'Chargeable capacity (kVA)',
        defaultFormat => '0soft',
        arguments     =>
          { IV1 => $totalAgreedCapacity, IV2 => $nonChargeableCapacity, },
        arithmetic => '=IV1-IV2',
    );
    my $chargeableCapacity953 = $chargeableCapacity;
    my $activeCoincidence953  = $activeCoincidence;
    my $exceededCapacity953   = $exceededCapacity;

    $totalAgreedCapacity = Arithmetic(
        name => 'Maximum export/import capacity adjusted for part-year (kVA)',
        defaultFormat => '0soft',
        arithmetic    => '=IV1*(1-IV2/IV3)',
        arguments     => {
            IV1 => $totalAgreedCapacity,
            IV2 => $tariffDaysInYearNot,
            IV3 => $daysInYear,
        }
    );

    $chargeableCapacity = Arithmetic(
        name          => 'Chargeable capacity adjusted for part-year (kVA)',
        defaultFormat => '0soft',
        arithmetic    => '=IV1*(1-IV2/IV3)',
        arguments     => {
            IV1 => $chargeableCapacity,
            IV2 => $tariffDaysInYearNot,
            IV3 => $daysInYear,
        }
    );

    $creditableCapacity = Arithmetic(
        name =>
          'Capacity eligible for GSP credits adjusted for part-year (kVA)',
        defaultFormat => '0soft',
        arithmetic    => '=IV1*(1-IV2/IV3)',
        arguments     => {
            IV1 => $creditableCapacity,
            IV2 => $tariffDaysInYearNot,
            IV3 => $daysInYear,
        }
    );

    push @{ $model->{tablesD} }, $exceededCapacity = Arithmetic(
        name          => 'Exceeded capacity adjusted for part-year (kVA)',
        defaultFormat => '0soft',
        arithmetic    => '=IV1*(1-IV2/IV3)',
        arguments     => {
            IV1 => $exceededCapacity,
            IV2 => $tariffDaysInYearNot,
            IV3 => $daysInYear,
        }
    );

    $activeCoincidence = Arithmetic(
        name       => 'Super red kW divided by kVA adjusted for part-year',
        arithmetic => '=IV1*(1-IV2/IV3)/(1-IV4/IV5)',
        arguments  => {
            IV1 => $activeCoincidence,
            IV2 => $tariffHoursInRedNot,
            IV3 => $hoursInRed,
            IV4 => $tariffDaysInYearNot,
            IV5 => $daysInYear,
        }
    );

    $reactiveCoincidence = Arithmetic(
        name       => 'Super red kW divided by kVA adjusted for part-year',
        arithmetic => '=IV1*(1-IV2/IV3)/(1-IV4/IV5)',
        arguments  => {
            IV1 => $reactiveCoincidence,
            IV2 => $tariffHoursInRedNot,
            IV3 => $hoursInRed,
            IV4 => $tariffDaysInYearNot,
            IV5 => $daysInYear,
        }
      )
      if $reactiveCoincidence;

    my $tariffSUraw = Arithmetic(
        name          => 'Chargeable sole use asset MEAV (£)',
        defaultFormat => '0soft',
        arguments     =>
          { IV1 => $tariffSoleUseMeav, IV2 => $tariffSoleUsePropex, },
        arithmetic => '=IV1*(1-IV2)',
    );

    my $tariffSU = Arithmetic(
        name          => 'Chargeable sole use asset MEAV for part-year (£)',
        defaultFormat => '0soft',
        arithmetic    => '=IV1*(1-IV2/IV3)',
        arguments     => {
            IV1 => $tariffSUraw,
            IV2 => $tariffDaysInYearNot,
            IV3 => $daysInYear,
        }
    );

    my $tariffLoc =
        $model->{method} =~ /LRIC/i ? $tariffPoint
      : $model->{method} =~ /FCP/i  ? $tariffGroup
      : undef;

    my $included = Arithmetic(
        name          => 'EDCM tariff?',
        arithmetic    => '=ISNUMBER(FIND(","&IV1&",",","&IV2&","))',
        arguments     => { IV2 => $edcmScope, IV1 => $customerClass },
        defaultFormat => 'boolsoft',
    );

    my ( $cdcmAssets, $cdcmEhvAssets, $cdcmHvLvShared, $cdcmHvLvService, ) =
      $model->cdcmAssets;

    unless ( $model->{noCapping} ) {

        0 and $activeCoincidence = Arithmetic(
            name       => 'Super-red kW/agreed kVA (capped)',
            arithmetic => '=MIN(MAX(0,IV1),1)',
            arguments  => { IV1 => $activeCoincidence }
        );

        $reactiveCoincidence = Arithmetic(
            name       => 'Super-red kVAr/agreed kVA (capped)',
            arithmetic =>
              '=MAX(MIN(SQRT(1-MIN(1,IV2)^2),IV1),0-SQRT(1-MIN(1,IV3)^2))',
            arguments => {
                IV1 => $reactiveCoincidence,
                IV2 => $activeCoincidence,
                IV3 => $activeCoincidence,
            }
        );

    }

    my (
        $cdcmUse,                 $lossFactors,
        $diversity,               $redUseRate,
        $capUseRate,              $assetsFixed,
        $assetsCapacity,          $assetsConsumption,
        $totalAssetsFixed,        $totalAssetsCapacity,
        $totalAssetsConsumption,  $totalAssetsGenerationSoleUse,
        $totalEdcmAssets,         $assetsCapacityCooked,
        $assetsConsumptionCooked, $assetsCapacityDoubleCooked,
        $assetsConsumptionDoubleCooked,
      )
      = $model->notionalAssets(
        $llfcs,             $tariffs,             $included,
        $activeCoincidence, $reactiveCoincidence, $totalAgreedCapacity,
        $exceededCapacity,  $powerFactorInModel,  $tariffCategory,
        $tariffDorG,        $tariffSU,            $cdcmAssets,
        $useProportions,    $ehvAssetLevelset,    $importForGenerator,
      );

    my $edcmRedUse = SumProduct(
        name          => 'Total EDCM peak time consumption (kW)',
        vector        => $redUseRate,
        matrix        => $totalAgreedCapacity,
        defaultFormat => '0softnz'
    );

    my $cdcmRedUse = Stack(
        cols => Labelset( list => [ $cdcmUse->{cols}{list}[0] ] ),
        name    => 'Total CDCM peak time consumption (kW)',
        sources => [$cdcmUse]
    );

    my $overallRedUse = Arithmetic(
        name          => 'Estimate total peak-time consumption (MW)',
        defaultFormat => '0softnz',
        arithmetic    => '=IV1+IV2',
        arguments     => { IV1 => $cdcmRedUse, IV2 => $edcmRedUse }
    );

    0 and push @{ $model->{allocDnoTables} },
      Columnset(
        name    => 'DNO-wide network use aggregates',
        columns => [ $edcmRedUse, $cdcmRedUse, $overallRedUse ]
      );

    my $chargeExit =
      $transmissionExitTotal;    # Stack( sources => [$transmissionExitTotal] );

    my $chargeDirect = $direct;  # Stack( sources => [$direct] );

    my $chargeIndirect = Arithmetic(
        name          => 'Charge for indirect costs (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=IV1*IV2',
        arguments     => { IV1 => $indirect, IV2 => $indirectProp }
    );

    my $chargeRates = $rates;    # Stack( sources => [$rates] );

    my $rateExit = Arithmetic(
        name       => 'Transmission exit charging rate (£/kW/year)',
        arithmetic => '=IV1/IV2',
        arguments  => { IV1 => $chargeExit, IV2 => $overallRedUse }
    );

    my $rateDirect = Arithmetic(
        name          => 'Direct cost charging rate',
        arithmetic    => '=IV1/(IV2+IV3+(IV4+IV5)/IV6)',
        defaultFormat => '%soft',
        arguments     => {
            IV1 => $chargeDirect,
            IV2 => $totalEdcmAssets,
            IV3 => $cdcmEhvAssets,
            IV4 => $cdcmHvLvShared,
            IV5 => $cdcmHvLvService,
            IV6 => $ehvIntensity,
        }
    );

    my $rateRates = Arithmetic(
        name          => 'Network rates charging rate',
        arithmetic    => '=IV1/(IV2+IV3+IV4+IV5)',
        defaultFormat => '%soft',
        arguments     => {
            IV1 => $chargeRates,
            IV2 => $totalEdcmAssets,
            IV3 => $cdcmEhvAssets,
            IV4 => $cdcmHvLvShared,
            IV5 => $cdcmHvLvService,
        }
    );

    my $rateIndirect = Arithmetic(
        name          => 'Indirect cost charging rate',
        arithmetic    => '=IV1/(IV20+IV3+(IV4+IV5)/IV6)',
        defaultFormat => '%soft',
        arguments     => {
            IV1  => $chargeIndirect,
            IV21 => $totalAssetsCapacity,
            IV22 => $totalAssetsConsumption,
            IV3  => $cdcmEhvAssets,
            IV4  => $cdcmHvLvShared,
            IV6  => $ehvIntensity,
            IV20 => $totalEdcmAssets,
            IV5  => $cdcmHvLvService,
        }
    );

    my $edcmIndirect = Arithmetic(
        name          => 'Indirect costs on EDCM demand (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=IV1*(IV20-IV23)',
        arguments     => {
            IV1  => $rateIndirect,
            IV21 => $totalAssetsCapacity,
            IV22 => $totalAssetsConsumption,
            IV20 => $totalEdcmAssets,
            IV23 => $totalAssetsGenerationSoleUse,
        },
    );

    my $edcmDirect = Arithmetic(
        name =>
'Direct costs on EDCM demand except through sole use asset charges (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=IV1*(IV20+IV23)',
        arguments     => {
            IV1  => $rateDirect,
            IV20 => $totalAssetsCapacity,
            IV23 => $totalAssetsConsumption,
        },
    );

    my $edcmRates = Arithmetic(
        name =>
'Network rates on EDCM demand except through sole use asset charges (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=IV1*(IV20+IV23)',
        arguments     => {
            IV1  => $rateRates,
            IV20 => $totalAssetsCapacity,
            IV23 => $totalAssetsConsumption,
        },
    );

    my $fixed3contribution = Arithmetic(
        name          => 'Fixed pot contribution p/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(IV1,100/IV2*IV3*(IV6+IV7+IV88),0)',
        arguments     => {
            IV1  => $included,
            IV3  => $tariffSU,
            IV6  => $rateDirect,
            IV7  => $rateIndirect,
            IV88 => $rateRates,
            IV2  => $daysInYear,
            IV4  => $tariffDorG,
        }
    );

    my $fixed3charge = Arithmetic(
        name          => 'Fixed charge p/day (scaled for part year)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(IV1,100/IV2*IV3*(IV6+IV88),0)',
        arguments     => {
            IV1  => $included,
            IV3  => $tariffSU,
            IV6  => $rateDirect,
            IV88 => $rateRates,
            IV2  => $daysInYear,
        }
    );

    my $fixed3chargeTrue =
      $model->{noGen}
      ? Arithmetic(
        name          => 'Fixed charge for demand p/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(AND(IV1,IV7="Demand"),100/IV2*IV3*(IV6+IV88),0)',
        arguments     => {
            IV1  => $included,
            IV3  => $tariffSUraw,
            IV6  => $rateDirect,
            IV88 => $rateRates,
            IV2  => $daysInYear,
            IV7  => $tariffDorG,
        }
      )
      : Arithmetic(
        name          => 'Fixed charge p/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(IV1,100/IV2*IV3*(IV6+IV88),0)',
        arguments     => {
            IV1  => $included,
            IV3  => $tariffSUraw,
            IV6  => $rateDirect,
            IV88 => $rateRates,
            IV2  => $daysInYear,
        }
      );

    my $charges1 = my $charges2 = [];

    my $exportCapacity = Constant(
        name => 'No FCP/LRIC export capacity charges',
        rows => $tariffs->{rows},
        data => [ [ map { 0 } 1 .. $model->{numTariffs} ] ],
    );

    my ( $acCoef, $reCoef );

    $model->{texit} = 4;

    my (
        $locations, $locLevel, $locDorG, $locParent,
        $c1,        $c2,       $a1d,     $r1d,
        $a1g,       $r1g,      $a2d,     $r2d,
        $a2g,       $r2g,      $gspa,    $gspb,
      )
      = $model->loadFlowInputs;

    ( $charges1, $charges2, $acCoef, $reCoef ) = $model->charge1and2(
        $tariffs, $llfcs,
        $included,
        $tariffLoc,
        $locations,
        $locParent,
        $c1,  $c2,  $a1d, $r1d, $a1g, $r1g, $a2d,
        $r2d, $a2g, $r2g,
        $model->preprocess(
            $locations,
            $a1d, $r1d, $a1g, $r1g, $a2d, $r2d, $a2g, $r2g,
        ),
        undef, $gspa, $gspb,
        $rateExit,
        $tariffDorG,
    );

    Columnset(
        name    => 'Annual hours in super red',
        columns => [
            $hoursInRed, $hoursInRed == $hoursInRedGeneration
            ? ()
            : $hoursInRedGeneration,
        ],
        number   => 1168,
        dataset  => $model->{dataset},
        appendTo => $model->{inputTables},
    );

    my (
        $fcpLricDemandCapacityChargeBig,
        $genCredit, $fcpLricGenerationCapacityChargeBig,
        $unitRateFcpLric, $genCreditCapacity,
      )
      = $model->chargesFcpLric(
        $acCoef,              $activeCoincidence,
        $charges1,            $charges2,
        $daysInYear,          $included,
        $llfcs,               $reactiveCoincidence,
        $reCoef,              $tariffNetworkSupportFactor,
        $tariffDorG,          $tariffs,
        $hoursInRed,          $hoursInRedGeneration,
        $totalAgreedCapacity, $creditableCapacity,
        $redUseRate,
      );

    $exportCapacity = Arithmetic(
        name       => 'Export capacity charge adjusted for GSM (p/kVA/day)',
        arithmetic =>
          '=IF(AND(IV1,IV2="Generation"),IF(IV6=0,1,IV4/IV5)*IV3,0)',
        defaultFormat => '0.00softnz',
        arguments     => {
            IV1 => $included,
            IV2 => $tariffDorG,
            IV3 => $fcpLricGenerationCapacityChargeBig,
            IV4 => $chargeableCapacity,
            IV5 => $totalAgreedCapacity,
            IV6 => $totalAgreedCapacity,
        }
    );

    my ($gPot) = $model->gPot(
        $included,   $tariffProportionPre2005,
        $tariffDorG, $totalAgreedCapacity
    );

    my ($generationScalingCharge) = $model->genScaling(
        $llfcs, $tariffs,
        $included,
        $tariffDorG,
        $totalAgreedCapacity,
        undef,
        $exportCapacity,
        $daysInYear,
        Arithmetic(
            name          => 'FCP/LRIC generation charges (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => '=IF(AND(IV8,IV9="Generation"),IV1*IV2*IV3/100,0)',
            arguments     => {
                IV1 => $exportCapacity,
                IV2 => $totalAgreedCapacity,
                IV3 => $daysInYear,
                IV8 => $included,
                IV9 => $tariffDorG,
            }
        ),
        $gPot,
    );

    my $exportCapacityScaled = Arithmetic(
        name          => 'Export capacity charge p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(AND(IV8,IV9="Generation"),IV1+IV2,0)',
        arguments     => {
            IV1 => $exportCapacity,
            IV2 => $generationScalingCharge,
            IV8 => $included,
            IV9 => $tariffDorG,
        },
    );

    push @{ $model->{tablesG} },
      Columnset(
        name    => 'Generation scaling',
        columns => [
            ( map { Stack( sources => [$_] ) } ( $llfcs, $tariffs ) ),
            $generationScalingCharge,
        ]
      );

    my $chargeOther =
      $model->{noGen}
      ? Arithmetic(
        name =>
          'Revenue less costs plus/minus non-CDCM generation bits (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=IV1-IV2-IV3-IV4-IV5',
        arguments     => {
            IV1 => $allowedRevenue,
            IV2 => $chargeDirect,
            IV3 => $chargeIndirect,
            IV4 => $chargeRates,
            IV5 => Dataset(
                name =>
                  'Net forecast income from non-CDCM generation (£/year)',
                defaultFormat => '0hardnz',
                dataset       => $model->{dataset},
                appendTo      => $model->{inputTables},
                number        => 1114,
                data          => [ [666] ],
            ),
        }
      )
      : Arithmetic(
        name => 'Revenue less costs plus/minus EDCM generation bits (£/year)',
        defaultFormat => '0softnz',
        arithmetic    =>
'=IV1-IV2-IV3-IV4-SUMPRODUCT(IV51_IV52,IV53_IV54)/100-SUMPRODUCT(IV61_IV62,IV63_IV64)*IV65/100-SUMPRODUCT(IV71_IV72,IV73_IV74)*IV75/100-SUMIF(IV81_IV82,"Generation",IV83_IV84)*IV85/100',
        arguments => {
            IV1       => $allowedRevenue,
            IV2       => $chargeDirect,
            IV3       => $chargeIndirect,
            IV4       => $chargeRates,
            IV51_IV52 => $genCredit,
            IV53_IV54 => $activeUnits,
            IV61_IV62 => $genCreditCapacity,
            IV63_IV64 => $totalAgreedCapacity,
            IV65      => $daysInYear,
            IV71_IV72 => $exportCapacityScaled,
            IV73_IV74 => $totalAgreedCapacity,
            IV75      => $daysInYear,
            IV81_IV82 => $tariffDorG,
            IV83_IV84 => 1 ? $fixed3charge : $fixed3contribution,
            IV85      => $daysInYear,
        }
      );

    my $rateOther = Arithmetic(
        name          => 'Other revenue charging rate',
        arithmetic    => '=IV1/((IV21+IV22)+IV3+IV4)',
        defaultFormat => '%soft',
        arguments     => {
            IV1  => $chargeOther,
            IV21 => $totalAssetsCapacity,
            IV22 => $totalAssetsConsumption,
            IV3  => $cdcmEhvAssets,
            IV4  => $cdcmHvLvShared,
            IV2  => $totalEdcmAssets,
            IV5  => $cdcmHvLvService,
            IV9  => $totalAssetsGenerationSoleUse,
        }
    );

    my $capacity3 = Arithmetic(
        name          => 'Capacity pot contribution p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic    =>
'=IF(AND(IV1,IV2="Demand"),100/IV3*((IV51+IV53)*(IV6+IV7+IV8+IV9)+IV41*IV42),0)',
        arguments => {
            IV1  => $included,
            IV2  => $tariffDorG,
            IV3  => $daysInYear,
            IV51 => $assetsCapacity,
            IV53 => $assetsConsumption,
            IV41 => $rateExit,
            IV42 => $model->{lossesFix}
              && $model->{lossesFix} > 2 ? $redUseRate : $activeCoincidence,
            IV6 => $rateDirect,
            IV7 => $rateIndirect,
            IV8 => $rateRates,
            IV9 => $rateOther,
        }
    );

    my $revenue3 = Arithmetic(
        name          => 'Pot contribution £/year',
        defaultFormat => '0softnz',
        arithmetic    => '=IF(AND(IV5,IV6="Demand"),IV9*0.01*(IV1+IV2*IV3),0)',
        arguments     => {
            IV1 => $fixed3contribution,
            IV2 => $capacity3,
            IV3 => $totalAgreedCapacity,
            IV9 => $daysInYear,
            IV5 => $included,
            IV6 => $tariffDorG,
        }
    );

    my $totalRevenue3 = GroupBy(
        name          => 'Pot £/year',
        defaultFormat => '0softnz',
        source        => $revenue3
    );

    my ( $scalingChargeFixed, $scalingChargeCapacity );

    my $capacityChargeT = Arithmetic(
        name          => 'Capacity charge p/kVA/day (exit only)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(AND(IV1,IV2="Demand"),100/IV3*IV41*IV42,0)',
        arguments     => {
            IV1  => $included,
            IV2  => $tariffDorG,
            IV3  => $daysInYear,
            IV41 => $rateExit,
            IV42 => $model->{lossesFix} ? $redUseRate : $activeCoincidence,
        }
    );

    $model->{summaryInformationColumns}[1] = Arithmetic(
        name          => 'Transmission exit charge (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=IF(AND(IV4,IV5="Demand"),0.01*(IV9-IV7)*IV1*IV2,0)',
        arguments     => {
            IV1 => $totalAgreedCapacity953,
            IV2 => $capacityChargeT,
            IV9 => $daysInYear,
            IV7 => $tariffDaysInYearNot,
            IV4 => $included,
            IV5 => $tariffDorG,
        },
    );

    push @{ $model->{tablesD} }, $unitRateFcpLric = Arithmetic(
        name          => 'Super red unit rate adjusted for DSM (p/kWh)',
        arithmetic    => '=IF(AND(IV1,IV2="Demand"),IF(IV6=0,1,IV4/IV5)*IV3,0)',
        defaultFormat => '0.000softnz',
        arguments     => {
            IV1 => $included,
            IV2 => $tariffDorG,
            IV3 => $unitRateFcpLric,
            IV4 => $chargeableCapacity,
            IV5 => $totalAgreedCapacity,
            IV6 => $totalAgreedCapacity,
        }
    );

    $capacityChargeT = Arithmetic(
        name       => 'Import capacity charge before scaling (p/VA/day)',
        arithmetic =>
          '=IF(AND(IV1,IV2="Demand"),IV7+IF(IV6=0,1,IV4/IV5)*IV3,0)',
        defaultFormat => '0.00softnz',
        arguments     => {
            IV1 => $included,
            IV2 => $tariffDorG,
            IV3 => $fcpLricDemandCapacityChargeBig,
            IV4 => $chargeableCapacity,
            IV5 => $totalAgreedCapacity,
            IV6 => $totalAgreedCapacity,
            IV7 => $capacityChargeT,
        }
    );

    $model->{summaryInformationColumns}[5] = Arithmetic(
        name          => 'FCP/LRIC charge (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=IF(AND(IV4,IV5="Demand"),0.01*IV9*IV1*(IV2-IV3),0)',
        arguments     => {
            IV1 => $totalAgreedCapacity,
            IV2 => $capacityChargeT,
            IV3 => $capacityChargeT->{arguments}{IV7},
            IV9 => $daysInYear,
            IV4 => $included,
            IV5 => $tariffDorG,
        },
    );

    $model->{Thursday31} = [
        Arithmetic(
            name       => 'FCP/LRIC capacity-based charge (£/year)',
            arithmetic => '=IF(AND(IV1,IV2="Demand"),IV3*(IV4+IV5)*IV9/100,0)',
            defaultFormat => '0softnz',
            arguments     => {
                IV1 => $included,
                IV2 => $tariffDorG,
                IV3 => $model->{demandCapacityFcpLric},
                IV4 => $chargeableCapacity,
                IV5 => $exceededCapacity,
                IV9 => $daysInYear,
            }
        ),
        Arithmetic(
            name          => 'FCP/LRIC unit-based charge (£/year)',
            arithmetic    => '=IF(AND(IV1,IV2="Demand"),IV3*IV4*IV9/100,0)',
            defaultFormat => '0softnz',
            arguments     => {
                IV1 => $included,
                IV2 => $tariffDorG,
                IV3 => $model->{demandConsumptionFcpLric},
                IV4 => $chargeableCapacity,
                IV9 => $daysInYear,
            }
        ),
        Arithmetic(
            name          => 'FCP/LRIC generation charges (£/year)',
            defaultFormat => '0softnz',
            arithmetic    =>
              '=IF(AND(IV8,IV9="Generation"),IV1*(IV4+IV5)*IV3/100,0)',
            arguments => {
                IV1 => $fcpLricGenerationCapacityChargeBig,
                IV4 => $chargeableCapacity,
                IV5 => $exceededCapacity,
                IV3 => $daysInYear,
                IV8 => $included,
                IV9 => $tariffDorG,
            }
        ),

    ];

    0 and my $demandScalingShortfallWithoutYnonError = Arithmetic(
        name          => 'Additional amount to be recovered (£/year)',
        defaultFormat => '0softnz',
        arithmetic    =>
'=IV1*(1-SUMPRODUCT(IV41_IV42,IV43_IV44)/100/IV7*(IV62+IV63+IV64+IV65)/(IV51+IV52+IV53+IV54))-(SUMIF(IV23_IV24,"Demand",IV21_IV22)',
        arguments => {
            IV1       => $totalRevenue3,
            IV31_IV32 => $capacityChargeT,
            IV33_IV34 => $totalAgreedCapacity,
            IV41_IV42 => $genCredit,
            IV43_IV44 => $activeUnits,
            IV62      => $totalEdcmAssets,
            IV63      => $cdcmEhvAssets,
            IV64      => $cdcmHvLvShared,
            IV65      => $cdcmHvLvService,
            IV51      => $totalAssetsFixed,
            IV52      => $totalAssetsCapacity,
            IV53      => $totalAssetsConsumption,
            IV54      => $cdcmEhvAssets,
            IV7       => $allowedRevenue,
            IV9       => $daysInYear,
            IV23_IV24 => $tariffDorG,
            IV21_IV22 => $fixed3charge,
        }
    );

    my $demandScalingShortfall = Arithmetic(
        name          => 'Additional amount to be recovered (£/year)',
        defaultFormat => '0softnz',
        arithmetic    =>
'=IV1-(SUMIF(IV23_IV24,"Demand",IV21_IV22)+SUMPRODUCT(IV31_IV32,IV33_IV34))*IV9/100',
        arguments => {
            IV1       => $totalRevenue3,
            IV31_IV32 => $capacityChargeT,
            IV33_IV34 => $totalAgreedCapacity,
            IV9       => $daysInYear,
            IV23_IV24 => $tariffDorG,
            IV21_IV22 => $fixed3charge,
        }
    );

    warn 'this is broken' && $model->fudge41(
        $included,                      $tariffDorG,
        $activeCoincidence,             $totalAgreedCapacity,
        $edcmIndirect,                  $edcmDirect,
        $edcmRates,                     $daysInYear,
        \$capacityChargeT,              \$demandScalingShortfall,
        $importForLdno,                 $assetsCapacityDoubleCooked,
        $assetsConsumptionDoubleCooked, $reactiveCoincidence,
        $powerFactorInModel,
      )
      if $model->{vedcm} == 41
      || $model->{vedcm} == 42
      || $model->{vedcm} == 49
      || $model->{vedcm} == 50
      || $model->{vedcm} == 51
      || $model->{vedcm} == 52
      || $model->{vedcm} == 43
      || $model->{vedcm} == 44
      || $model->{vedcm} == 45
      || $model->{vedcm} == 46
      || $model->{vedcm} == 47;

    $model->fudge41(
        $included,                      $tariffDorG,
        $activeCoincidence,             $totalAgreedCapacity,
        $edcmIndirect,                  $edcmDirect,
        $edcmRates,                     $daysInYear,
        \$capacityChargeT,              \$demandScalingShortfall,
        $importForLdno,                 $assetsCapacityDoubleCooked,
        $assetsConsumptionDoubleCooked, $reactiveCoincidence,
        $powerFactorInModel,
      )
      if $model->{vedcm} > 52 && $model->{vedcm} < 61;

    ($scalingChargeCapacity) = $model->demandScaling41(
        $included,             $tariffDorG,
        $totalAgreedCapacity,  $demandScalingShortfall,
        $daysInYear,           $assetsFixed,
        $assetsCapacityCooked, $assetsConsumptionCooked,
        $capacityChargeT,      $fixed3charge,
      )
      if $model->{vedcm} == 42
      || $model->{vedcm} > 52 && $model->{vedcm} < 61
      || $model->{vedcm} == 49
      || $model->{vedcm} == 50
      || $model->{vedcm} == 51
      || $model->{vedcm} == 52
      || $model->{vedcm} == 43
      || $model->{vedcm} == 44
      || $model->{vedcm} == 45
      || $model->{vedcm} == 46
      || $model->{vedcm} == 47;

    $model->{summaryInformationColumns}[2] = Arithmetic(
        name          => 'Direct cost allocation (£/year)',
        defaultFormat => '0softnz',
        arithmetic => '=IF(AND(IV4,IV5="Demand"),IV1*IV2*IV3*0.01*IV7/IV9,0)',
        arguments  => {
            IV1 => $totalAgreedCapacity,
            IV2 => $scalingChargeCapacity,
            IV3 => $daysInYear,
            IV4 => $included,
            IV5 => $tariffDorG,
            IV7 => $edcmDirect,
            IV8 => $edcmRates,
            IV9 => $demandScalingShortfall,
        },
    );

    $model->{summaryInformationColumns}[4] = Arithmetic(
        name          => 'Network rates allocation (£/year)',
        defaultFormat => '0softnz',
        arithmetic => '=IF(AND(IV4,IV5="Demand"),IV1*IV2*IV3*0.01*IV8/IV9,0)',
        arguments  => {
            IV1 => $totalAgreedCapacity,
            IV2 => $scalingChargeCapacity,
            IV3 => $daysInYear,
            IV4 => $included,
            IV5 => $tariffDorG,
            IV7 => $edcmDirect,
            IV8 => $edcmRates,
            IV9 => $demandScalingShortfall,
        },
    );

    $model->{summaryInformationColumns}[7] = Arithmetic(
        name          => 'Demand scaling asset based (£/year)',
        defaultFormat => '0softnz',
        arithmetic    =>
          '=IF(AND(IV4,IV5="Demand"),IV1*IV2*IV3*0.01*(1-(IV8+IV7)/IV9),0)',
        arguments => {
            IV1 => $totalAgreedCapacity,
            IV2 => $scalingChargeCapacity,
            IV3 => $daysInYear,
            IV4 => $included,
            IV5 => $tariffDorG,
            IV7 => $edcmDirect,
            IV8 => $edcmRates,
            IV9 => $demandScalingShortfall,
        },
    );

    0 and push @{ $model->{demandChargeTables} },
      Columnset(
        name    => 'Demand scaling',
        columns => [
            1
            ? ()
            : ( map { Stack( sources => [$_] ) } ( $llfcs, $tariffs ) ),
            $scalingChargeFixed    ? $scalingChargeFixed    : (),
            $scalingChargeCapacity ? $scalingChargeCapacity : ()
        ]
      );

    0
      and push @{ $model->{demandChargeTables} },
      Arithmetic(
        name          => 'Amount recovered through demand scaling (£/year)',
        defaultFormat => '0softnz',
        arithmetic    =>
          '=(SUM(IV21_IV22)+SUMPRODUCT(IV31_IV32,IV33_IV34))*IV9/100',
        arguments => {
            IV31_IV32 => $scalingChargeCapacity,
            IV33_IV34 => $totalAgreedCapacity,
            IV9       => $daysInYear,
            IV21_IV22 => $scalingChargeFixed,
        }
      )
      if $scalingChargeFixed
      and $scalingChargeCapacity;

    0
      and push @{ $model->{demandChargeTables} },
      Arithmetic(
        name          => 'Amount recovered through demand scaling (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=SUMPRODUCT(IV31_IV32,IV33_IV34)*IV9/100',
        arguments     => {
            IV31_IV32 => $scalingChargeCapacity,
            IV33_IV34 => $totalAgreedCapacity,
            IV9       => $daysInYear,
        }
      )
      if !$scalingChargeFixed
      and $scalingChargeCapacity;

    #   }

    my $importCapacityScaled =
      $scalingChargeCapacity
      ? Arithmetic(
        name          => 'Total import capacity charge p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(AND(IV8,IV9="Demand"),IV1+IV2,0)',
        arguments     => {
            IV1 => $capacityChargeT,
            IV2 => $scalingChargeCapacity,
            IV8 => $included,
            IV9 => $tariffDorG,
        }
      )
      : Stack( sources => [$capacityChargeT] );

    0 and push @{ $model->{demandChargeTables} },
      Columnset(
        name    => 'Application of demand scaling',
        columns => [
            1
            ? ()
            : (
                Arithmetic(
                    name          => $llfcs->{name},
                    arithmetic    => '=IF(IV1,IV2,"")',
                    arguments     => { IV1 => $included, IV2 => $llfcs },
                    defaultFormat => 'textcopy',
                ),
                Arithmetic(
                    name          => $tariffs->{name},
                    arithmetic    => '=IF(IV1,IV2,"")',
                    arguments     => { IV1 => $included, IV2 => $tariffs },
                    defaultFormat => 'textcopy',
                )
            ),
            $importCapacityScaled
        ]
      );

    $unitRateFcpLric = Arithmetic(
        name       => 'Super red rate p/kWh',
        arithmetic =>
          '=IF(AND(IV1,IV2="Demand"),IF(IV3=0,IV9,MIN(IV4,IV5/IV6*IV7/IV8)),0)',
        arguments => {
            IV1 => $included,
            IV2 => $tariffDorG,
            IV3 => $activeCoincidence,
            IV4 => $unitRateFcpLric,
            IV9 => $unitRateFcpLric,
            IV5 => $importCapacityScaled,
            IV6 => $activeCoincidence,
            IV7 => $daysInYear,
            IV8 => $hoursInRed,
        }
      )
      if $unitRateFcpLric;

    push @{ $model->{tablesD} },
      my $importCapacityScaledSaved = $importCapacityScaled;

    $importCapacityScaled = Arithmetic(
        name       => 'Import capacity p/kVA/day',
        arithmetic => '=IF(AND(IV1,IV2="Demand"),IV3-IV4*IV7*IV8/IV9,0)',
        arguments  => {
            IV1 => $included,
            IV2 => $tariffDorG,
            IV3 => $importCapacityScaled,
            IV9 => $daysInYear,
            IV7 => $activeCoincidence,
            IV4 => $unitRateFcpLric,
            IV8 => $hoursInRed,
        },
        defaultFormat => '0.00softnz'
      )
      if $unitRateFcpLric;

    my $importCapacityExceeded = Arithmetic(
        name          => 'Exceeded import capacity charge (p/kVA/day)',
        defaultFormat => '0.00softnz',
        arithmetic    =>
          '=IV7+IF(AND(IV1,IV2="Demand"),IF(IV6=0,0,1-IV4/IV5)*IV3,0)',
        defaultFormat => '0.000softnz',
        arguments     => {
            IV1 => $included,
            IV2 => $tariffDorG,
            IV3 => $fcpLricDemandCapacityChargeBig,
            IV4 => $chargeableCapacity,
            IV5 => $totalAgreedCapacity,
            IV6 => $totalAgreedCapacity,
            IV7 => $importCapacityScaled,
        },
        defaultFormat => '0.00softnz'
    );

    my $exportCapacityExceeded = Arithmetic(
        name          => 'Exceeded export capacity charge (p/kVA/day)',
        defaultFormat => '0.00softnz',
        arithmetic    =>
          '=IV7+IF(AND(IV1,IV2="Generation"),IF(IV6=0,0,1-IV4/IV5)*IV3,0)',
        arguments => {
            IV1 => $included,
            IV2 => $tariffDorG,
            IV3 => $fcpLricGenerationCapacityChargeBig,
            IV4 => $chargeableCapacity,
            IV5 => $totalAgreedCapacity,
            IV6 => $totalAgreedCapacity,
            IV7 => $exportCapacityScaled,
        },
    );

    push @{ $model->{tablesG} }, $genCredit, $genCreditCapacity,
      $exportCapacityScaled;

    push @{ $model->{tariffTables} }, Columnset(
        name    => 'EDCM charge',
        columns => [
            Arithmetic(
                name          => $llfcs->{name},
                arithmetic    => '=IF(IV1,IV2,"")',
                arguments     => { IV1 => $included, IV2 => $llfcs },
                defaultFormat => 'textcopy',
            ),
            Arithmetic(
                name          => $tariffs->{name},
                arithmetic    => '=IF(IV1,IV2,"")',
                arguments     => { IV1 => $included, IV2 => $tariffs },
                defaultFormat => 'textcopy',
            ),
            $unitRateFcpLric ? $unitRateFcpLric : (),
            $fixed3chargeTrue,
            $importCapacityScaled,
            $importCapacityExceeded,
            $model->{noGen}
            ? ()
            : (
                Stack(
                    sources       => [$exportCapacityScaled],
                    defaultFormat => '0.00copynz'
                ),
                $exportCapacityExceeded,
                Stack( sources => [$genCredit] ),
                Stack(
                    sources       => [$genCreditCapacity],
                    defaultFormat => '0.00copynz'
                ),
            )
          ]

    );

    return $model unless $model->{summaries};

    my @revenueBits = (

        Arithmetic(
            name          => 'Capacity charge for demand (£/year)',
            defaultFormat => '0softnz',
            arithmetic => '=IF(AND(IV4,IV5="Demand"),0.01*(IV9-IV7)*IV8*IV1,0)',
            arguments  => {
                IV1 => $importCapacityScaled,
                IV9 => $daysInYear,
                IV7 => $tariffDaysInYearNot,
                IV4 => $included,
                IV5 => $tariffDorG,
                IV8 => $totalAgreedCapacity953,
            }
        ),

        Arithmetic(
            name          => 'Super red charge for demand (£/year)',
            defaultFormat => '0softnz',
            arithmetic    =>
              '=IF(AND(IV4,IV5="Demand"),0.01*(IV9-IV7)*IV1*IV6*IV8,0)',
            arguments => {
                IV1 => $unitRateFcpLric,
                IV9 => $hoursInRed,
                IV7 => $tariffHoursInRedNot,
                IV4 => $included,
                IV5 => $tariffDorG,
                IV6 => $totalAgreedCapacity953,
                IV8 => $activeCoincidence953,
            }
        ),

        Arithmetic(
            name          => 'Fixed charge for demand (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => '=IF(AND(IV4,IV5="Demand"),0.01*(IV9-IV7)*IV1,0)',
            arguments     => {
                IV1 => $fixed3chargeTrue,
                IV9 => $daysInYear,
                IV7 => $tariffDaysInYearNot,
                IV4 => $included,
                IV5 => $tariffDorG,
            }
        ),

        Arithmetic(
            name          => 'Exceeded capacity charge for demand (£/year)',
            defaultFormat => '0softnz',
            arithmetic => '=IF(AND(IV4,IV5="Demand"),0.01*(IV9-IV7)*IV8*IV1,0)',
            arguments  => {
                IV1 => $importCapacityExceeded,
                IV9 => $daysInYear,
                IV7 => $tariffDaysInYearNot,
                IV4 => $included,
                IV5 => $tariffDorG,
                IV8 => $exceededCapacity953,
            }
        ),

    );

    my $r1 = Arithmetic(
        name       => $previousIncome->{name},
        arithmetic => '=IF(AND(IV1,IV3="Demand"),IV2,0)',
        arguments  =>
          { IV1 => $included, IV2 => $previousIncome, IV3 => $tariffDorG, },
        defaultFormat => '0copynz',
    );

    my $r2 = Arithmetic(
        name          => 'Annual charge (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=' . join( '+', map { "IV$_" } 1 .. @revenueBits ),
        arguments     =>
          { map { ( "IV$_" => $revenueBits[ $_ - 1 ] ) } 1 .. @revenueBits },
    );

    my $change1 = Arithmetic(
        name          => 'Change (£/year)',
        arithmetic    => '=IF(AND(IV1,IV2="Demand"),IV3-IV4,0)',
        defaultFormat => '0softpm',
        arguments     =>
          { IV1 => $included, IV2 => $tariffDorG, IV3 => $r2, IV4 => $r1 }
    );

    my $change2 = Arithmetic(
        name          => 'Change (%)',
        arithmetic    => '=IF(IV1,IV3/IV4-1,"")',
        defaultFormat => '%softpm',
        arguments     => { IV1 => $r1, IV3 => $r2, IV4 => $r1 }
    );

    push @{ $model->{revenueTables} },
      Columnset(
        name    => 'Summary information part 1',
        columns => [
            Arithmetic(
                name       => $llfcs->{name},
                arithmetic => '=IF(AND(IV1,IV3="Demand"),IV2,"")',
                arguments  =>
                  { IV1 => $included, IV2 => $llfcs, IV3 => $tariffDorG, },
                defaultFormat => 'textcopy',
            ),
            Arithmetic(
                name       => $tariffs->{name},
                arithmetic => '=IF(AND(IV1,IV3="Demand"),IV2,"")',
                arguments  => {
                    IV1 => $included,
                    IV2 => $tariffs,
                    IV3 => $tariffDorG,
                },
                defaultFormat => 'textcopy',
            ),
            @revenueBits
        ]
      );

    push @{ $model->{revenueTables} },
      Columnset(
        name    => 'Summary information part 2',
        columns => [
            Arithmetic(
                name       => $llfcs->{name},
                arithmetic => '=IF(AND(IV1,IV3="Demand"),IV2,"")',
                arguments  =>
                  { IV1 => $included, IV2 => $llfcs, IV3 => $tariffDorG, },
                defaultFormat => 'textcopy',
            ),
            Arithmetic(
                name       => $tariffs->{name},
                arithmetic => '=IF(AND(IV1,IV3="Demand"),IV2,"")',
                arguments  => {
                    IV1 => $included,
                    IV2 => $tariffs,
                    IV3 => $tariffDorG,
                },
                defaultFormat => 'textcopy',
            ),
            $r2, $r1, $change1, $change2,
            Arithmetic(
                name          => "Super red units (kWh)",
                defaultFormat => '0softnz',
                arithmetic => '=IF(AND(IV1,IV2="Demand"),IV4*(IV3-IV7)*IV5,0)',
                arguments  => {
                    IV1 => $included,
                    IV2 => $tariffDorG,
                    IV3 => $hoursInRed,
                    IV7 => $tariffHoursInRedNot,
                    IV4 => $activeCoincidence953,
                    IV5 => $totalAgreedCapacity953,
                }
            ),
            map { Stack( sources => [$_] ) } $chargeableCapacity953
        ]
      );

    $model->{summaryInformationColumns}[0] = Stack(
        name    => 'Sole use asset charge (£/year)',
        sources => [ $revenueBits[2] ]
    );

    push @{ $model->{revenueTables} }, Columnset(
        name    => 'Summary information part 3',
        columns => [
            Arithmetic(
                name       => $llfcs->{name},
                arithmetic => '=IF(AND(IV1,IV3="Demand"),IV2,"")',
                arguments  =>
                  { IV1 => $included, IV2 => $llfcs, IV3 => $tariffDorG, },
                defaultFormat => 'textcopy',
            ),
            Arithmetic(
                name       => $tariffs->{name},
                arithmetic => '=IF(AND(IV1,IV3="Demand"),IV2,"")',
                arguments  => {
                    IV1 => $included,
                    IV2 => $tariffs,
                    IV3 => $tariffDorG,
                },
                defaultFormat => 'textcopy',
            ),
            ( grep { $_ } @{ $model->{summaryInformationColumns} } ),
            Arithmetic(
                name          => 'Check (£/year)',
                defaultFormat => '0softnz',
                arithmetic    => join(
                    '',
                    '=IV1-IV9',
                    map {
                        $model->{summaryInformationColumns}[$_]
                          ? ( "-IV" . ( 20 + $_ ) )
                          : ()
                      } 0 .. $#{ $model->{summaryInformationColumns} }
                ),
                arguments => {
                    IV1 => $r2,
                    IV9 => $revenueBits[$#revenueBits],
                    map {
                        $model->{summaryInformationColumns}[$_]
                          ? (
                            "IV" . ( 20 + $_ ),
                            $model->{summaryInformationColumns}[$_]
                          )
                          : ()
                      } 0 .. $#{ $model->{summaryInformationColumns} }
                }
            ),
        ]
    );

    return $model if $model->{noGen};

    my $revenue = $model->revenue(
        $daysInYear,             $llfcs,
        $tariffs,                $tariffDorG,
        $included,               $totalAgreedCapacity,
        $exceededCapacity,       $activeUnits,
        $fixed3charge,           $importCapacityScaledSaved,
        $exportCapacityScaled,   $importCapacityExceeded,
        $exportCapacityExceeded, $genCredit,
        $genCreditCapacity,      $importCapacityScaled,
        $unitRateFcpLric,        $activeCoincidence,
        $hoursInRed,
    );

    $model->summary(
        $llfcs,               $tariffs,           $tariffDorG,
        $included,            $revenue,           $previousIncome,
        $totalAgreedCapacity, $activeCoincidence, $charges1,
        $charges2,
    );

    $model;

}

1;
