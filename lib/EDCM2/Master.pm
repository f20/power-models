package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2012 Energy Networks Association Limited and others.
Copyright 2013-2017 Franck Latrémolière, Reckon LLP and others.

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

sub requiredModulesForRuleset {
    my ( $class, $ruleset ) = @_;
    qw(
      EDCM2::Adjust
      EDCM2::Assets
      EDCM2::Charge1
      EDCM2::Charging
      EDCM2::Generation
      EDCM2::Inputs
      EDCM2::Locations
      EDCM2::PotsAndRates
      EDCM2::Scaling
      EDCM2::Sheets
      ),
      $ruleset->{ldnoRev}   ? qw(EDCM2::Ldno)      : (),
      $ruleset->{summaries} ? qw(EDCM2::Summaries) : (),
      $ruleset->{transparency}
      && $ruleset->{transparency} =~ /impact/i ? qw(EDCM2::Impact)   : (),
      $ruleset->{customerTemplates}            ? qw(EDCM2::Template) : (),
      $ruleset->{mitigateUndueSecrecy} ? qw(EDCM2::SecrecyMitigation) : (),
      $ruleset->{voltageRulesTransparency} ? () : qw(EDCM2::AssetCalcHard),
      $ruleset->{layout} ? qw(SpreadsheetModel::MatrixSheet EDCM2::Layout) : (),
      $ruleset->{checksums} ? qw(SpreadsheetModel::Checksum) : ();
}

sub new {

    my $class = shift;
    my $model = bless {@_}, $class;

    die 'This system will not build an orange '
      . 'EDCM model without a suitable disclaimer.' . "\n--"
      if $model->{colour}
      && $model->{colour} =~ /orange/i
      && !($model->{extraNotice}
        && length( $model->{extraNotice} ) > 299
        && $model->{extraNotice} =~ /DCUSA/ );

    $model->{inputTables} = [];
    $model->{method} ||= 'none';

    # The EDCM timeband is called purple in this code;
    # its display name defaults to super-red.
    $model->{TimebandName} ||=
      ucfirst( $model->{timebandName} ||= 'super-red' );

  # Keep EDCM2::DataPreprocess out of the scope of revision number construction.
    if ( $model->{dataset}
        && keys %{ $model->{dataset} } )
    {
        if ( eval { require EDCM2::DataPreprocess; } ) {
            $model->preprocessDataset;
        }
        else {
            warn $@;
        }
    }

    $model->{numLocations} ||= $model->{numLocationsDefault};
    $model->{numTariffs}   ||= $model->{numTariffsDefault};

    my $ehvAssetLevelset = Labelset(
        name => 'EHV asset levels',
        list => [ split /\n/, <<EOT ] );
GSP
132kV circuits
132kV/EHV
EHV circuits
EHV/HV
132kV/HV
EOT

    my (
        $tariffs,
        $importCapacity935,
        $exportCapacity935Exempt,
        $exportCapacity935ChargeablePre2005,
        $exportCapacity935Chargeable20052010,
        $exportCapacity935ChargeablePost2010,
        $tariffSoleUseMeav,
        $dcp189Input,
        $tariffLoc,
        $tariffCategory,
        $useProportions,
        $activeCoincidence935,
        $reactiveCoincidence935,
        $indirectExposure,
        $nonChargeableCapacity935,
        $activeUnits,
        $creditableCapacity935,
        $tariffNetworkSupportFactor,
        $tariffDaysInYearNot,
        $tariffHoursInPurpleNot,
        $previousChargeImport,
        $previousChargeExport,
        $llfcImport,
        $llfcExport,
        $actualRedDemandRate,
    ) = $model->tariffInputs($ehvAssetLevelset);

    my ( $locations, $locParent, $c1, $a1d, $r1d, $a1g, $r1g ) =
      $model->loadFlowInputs;

    if (   $model->{transparency}
        && $model->{transparency} !~ /outputonly/i
        && !$model->{legacy201} )
    {
        my $flagInput = Dataset(
            name => 'Is this the master model containing all the tariff data?',
            defaultFormat => 'boolhard',
            singleRowName => 'Enter TRUE or FALSE',
            validation    => {
                validate => 'list',
                value    => [ 'TRUE', 'FALSE' ],
            },
            data     => ['TRUE'],
            dataset  => $model->{dataset},
            appendTo => $model->{inputTables},
            number   => 1190,
        );
        $model->{transparencyMasterFlag} = Arithmetic(
            name          => 'Is this the master model?',
            defaultFormat => 'boolsoft',
            arithmetic    => '=IF(ISERROR(A5),TRUE,'
              . 'IF(A4="FALSE",FALSE,IF(A3=FALSE,FALSE,TRUE)))',
            arguments => {
                A3 => $flagInput,
                A4 => $flagInput,
                A5 => $flagInput,
            },
        );
    }

    if ( $model->{transparency} ) {

        if ( $model->{transparency} =~ /impact/i ) {

            $model->impactNotes;

            ( $locations, $locParent, $c1, $a1d, $r1d, $a1g, $r1g ) =
              $model->mangleLoadFlowInputs( $locations, $locParent, $c1, $a1d,
                $r1d, $a1g, $r1g );

            (
                $tariffs,
                $importCapacity935,
                $exportCapacity935Exempt,
                $exportCapacity935ChargeablePre2005,
                $exportCapacity935Chargeable20052010,
                $exportCapacity935ChargeablePost2010,
                $tariffSoleUseMeav,
                $tariffLoc,
                $tariffCategory,
                $useProportions,
                $activeCoincidence935,
                $reactiveCoincidence935,
                $indirectExposure,
                $nonChargeableCapacity935,
                $activeUnits,
                $creditableCapacity935,
                $tariffNetworkSupportFactor,
                $tariffDaysInYearNot,
                $tariffHoursInPurpleNot,
                $previousChargeImport,
                $previousChargeExport,
                $llfcImport,
                $llfcExport,
                $actualRedDemandRate,
                $model->{transparency},
              )
              = $model->mangleTariffInputs(
                $tariffs,
                $importCapacity935,
                $exportCapacity935Exempt,
                $exportCapacity935ChargeablePre2005,
                $exportCapacity935Chargeable20052010,
                $exportCapacity935ChargeablePost2010,
                $tariffSoleUseMeav,
                $tariffLoc,
                $tariffCategory,
                $useProportions,
                $activeCoincidence935,
                $reactiveCoincidence935,
                $indirectExposure,
                $nonChargeableCapacity935,
                $activeUnits,
                $creditableCapacity935,
                $tariffNetworkSupportFactor,
                $tariffDaysInYearNot,
                $tariffHoursInPurpleNot,
                $previousChargeImport,
                $previousChargeExport,
                $llfcImport,
                $llfcExport,
              );

            $model->{transparencyImpact} = 1;

        }
        elsif ( $model->{transparency} =~ /outputonly/i ) {
            $model->{transparency} = {};
        }
        else {
            delete $model->{transparency};
        }

    }

    $model->{transparency} = Arithmetic(
        name  => 'Weighting of each tariff for reconciliation of totals',
        lines => [
            '0 means that the tariff is active and'
              . ' is included in the table 119x aggregates.',
            '-1 means that the tariff is included in'
              . ' the table 119x aggregates but should be removed.',
            '1 means that the tariff is active and'
              . ' is not included in the table 119x aggregates.',
        ],
        arithmetic => '=IF(OR(A3,'
          . 'NOT(ISERROR(SEARCH("[ADDED]",A2))))' . ',1,'
          . 'IF(ISERROR(SEARCH("[REMOVED]",A1)),0,-1)' . ')',
        arguments => {
            A1 => $tariffs,
            A2 => $tariffs,
            A3 => $model->{transparencyMasterFlag},
        },
    ) if $model->{transparencyMasterFlag} && !defined $model->{transparency};

    if ( $model->{transparencyMasterFlag} ) {

        foreach my $set (
            [
                1191,
                [ 'Baseline total EDCM peak time consumption (kW)', '0hard' ],
                [
                    'Baseline total marginal effect'
                      . ' of indirect cost adder (kVA)',
                    '0hard'
                ],
                [
                    'Baseline total marginal effect of demand adder (kVA)',
                    '0hard'
                ],
                [ 'Baseline revenue from demand charge 1 (£/year)', '0hard' ],
                [
                    'Baseline net forecast EDCM generation revenue (£/year)',
                    '0hard'
                ],
                'EDCM demand and revenue aggregates'
            ],
            [
                1192,
                [ 'Baseline total chargeable export capacity (kVA)', '0hard' ],
                [
                    'Baseline total non-exempt 2005-2010 export capacity (kVA)',
                    '0hard'
                ],
                [
                    'Baseline total non-exempt post-2010 export capacity (kVA)',
                    '0hard'
                ],
                'EDCM generation aggregates'
            ],
            [
                1193,
                [ 'Baseline total sole use assets for demand (£)', '0hard' ],
                [
                    'Baseline total sole use assets for generation (£)',
                    '0hard'
                ],
                [ 'Baseline total notional capacity assets (£)',    '0hard' ],
                [ 'Baseline total notional consumption assets (£)', '0hard' ],
                [
                    'Baseline total non sole use'
                      . ' notional assets subject to matching (£)',
                    '0hard'
                ],
                $model->{dcp189} && $model->{dcp189} =~ /preservePot|split/i
                ? [
                    'Baseline total demand sole use assets '
                      . 'qualifying for DCP 189 discount (£)',
                    '0hard'
                  ]
                : (),
                'EDCM notional asset aggregates'
            ],
          )
        {
            my @cols;
            foreach my $col ( 1 .. ( $#$set - 1 ) ) {
                push @cols,
                  $model->{transparency}{baselineItem}{ 100 * $set->[0] + $col }
                  = Dataset(
                    name          => $set->[$col][0],
                    defaultFormat => $set->[$col][1],
                    data          => [ [''] ],
                  );
            }
            Columnset(
                name          => "Baseline $set->[$#$set]",
                singleRowName => $set->[$#$set],
                number        => $set->[0],
                dataset       => $model->{dataset},
                appendTo      => $model->{inputTables},
                columns       => \@cols,
            );
        }

    }

    $model->{mitigateUndueSecrecy} = EDCM2::SecrecyMitigation->new($model)
      if $model->{mitigateUndueSecrecy};

    my ( $cdcmAssets, $cdcmEhvAssets, $cdcmHvLvShared, $cdcmHvLvService, ) =
      $model->cdcmAssets;

    if ( $model->{ldnoRev} && $model->{ldnoRev} =~ /only/i ) {
        $model->{daysInYear} = Dataset(
            name          => 'Days in year',
            defaultFormat => '0hard',
            data          => [365],
            dataset       => $model->{dataset},
            appendTo      => $model->{inputTables},
            number        => $model->{table1101} ? 1110 : 1111,
            validation    => {
                validate => 'decimal',
                criteria => 'between',
                minimum  => 365,
                maximum  => 366,
            },
        );
        $model->ldnoRev;
        return $model;
    }

    $model->{matricesData} = [ [], [] ]
      if $model->{summaries} && $model->{summaries} =~ /matri/i;

    my (
        $daysInYear,            $chargeDirect,
        $chargeIndirect,        $chargeRates,
        $chargeExit,            $ehvIntensity,
        $allowedRevenue,        $powerFactorInModel,
        $genPot20p,             $genPotGP,
        $genPotGL,              $genPotCdcmCap20052010,
        $genPotCdcmCapPost2010, $hoursInPurple,
    ) = $model->generalInputs;

    $model->ldnoRev if $model->{ldnoRev};

    my $exportEligible = Arithmetic(
        name          => 'Has export charges?',
        defaultFormat => 'boolsoft',
        arithmetic    => '=OR(A1<>"VOID",A2<>"VOID",A3<>"VOID")',
        arguments     => {
            A1 => $exportCapacity935ChargeablePre2005,
            A2 => $exportCapacity935Chargeable20052010,
            A3 => $exportCapacity935ChargeablePost2010,
        }
    );

    my $importEligible = Arithmetic(
        name          => 'Has import charges?',
        defaultFormat => 'boolsoft',
        arithmetic    => '=A1<>"VOID"',
        arguments     => {
            A1 => $importCapacity935,
        }
    );

    my (
        $chargeableCapacity,               $exportCapacityChargeable,
        $importCapacity,                   $activeCoincidence,
        $reactiveCoincidence,              $creditableCapacity,
        $exportCapacityChargeable20052010, $exportCapacityChargeablePost2010,
        $exportCapacityChargeablePre2005,  $exportCapacityExempt,
        $demandSoleUseAsset,               $generationSoleUseAsset,
        $demandSoleUseAssetUnscaled,       $generationSoleUseAssetUnscaled
      )
      = $model->preliminaryAdjustments(
        $daysInYear,
        $hoursInPurple,
        $tariffDaysInYearNot,
        $tariffHoursInPurpleNot,
        $importCapacity935,
        $nonChargeableCapacity935,
        $activeCoincidence935,
        $reactiveCoincidence935,
        $creditableCapacity935,
        $exportCapacity935Chargeable20052010,
        $exportCapacity935ChargeablePost2010,
        $exportCapacity935ChargeablePre2005,
        $exportCapacity935Exempt,
        $tariffSoleUseMeav,
      );

    $reactiveCoincidence935 = Arithmetic(
        name       => 'Unadjusted but capped red kVAr/agreed kVA',
        arithmetic => '=MAX(MIN(SQRT(1-MIN(1,A2)^2),'
          . ( $model->{legacy201} ? '' : '0+' )
          . 'A1),0-SQRT(1-MIN(1,A3)^2))',
        arguments => {
            A1 => $reactiveCoincidence935,
            A2 => $activeCoincidence935,
            A3 => $activeCoincidence935,
        }
    );

    my $cdcmUse = $model->{cdcmComboTable} ? Stack(
        name => 'Forecast system simultaneous maximum load (kW)'
          . ' from CDCM users',
        defaultFormat => '0copy',
        cols          => $ehvAssetLevelset,
        rows          => 0,
        rowName       => 'System simultaneous maximum load (kW)',
        sources       => [ $model->{cdcmComboTable} ],
      ) : Dataset(
        name => 'Forecast system simultaneous maximum load (kW)'
          . ' from CDCM users'
          . ( $model->{transparency} ? '' : ' (from CDCM table 2506)' ),
        defaultFormat => '0hardnz',
        cols          => $ehvAssetLevelset,
        data          => [qw(5e6 5e6 5e6 5e6 5e6 5e6)],
        number        => 1122,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        }
      );

    my (
        $lossFactors,            $diversity,
        $accretion,              $purpleUseRate,
        $capUseRate,             $assetsFixed,
        $assetsCapacity,         $assetsConsumption,
        $totalAssetsFixed,       $totalAssetsCapacity,
        $totalAssetsConsumption, $totalAssetsGenerationSoleUse,
        $totalEdcmAssets,        $assetsCapacityCooked,
        $assetsConsumptionCooked,
      )
      = $model->notionalAssets(
        $activeCoincidence,      $reactiveCoincidence,
        $importCapacity,         $powerFactorInModel,
        $tariffCategory,         $demandSoleUseAsset,
        $generationSoleUseAsset, $cdcmAssets,
        $useProportions,         $ehvAssetLevelset,
        $cdcmUse,
      );

    my ( $rateDirect, $rateRates, $rateIndirect, ) = $model->chargingRates(
        $chargeDirect,    $chargeRates,   $chargeIndirect,
        $totalEdcmAssets, $cdcmEhvAssets, $cdcmHvLvShared,
        $cdcmHvLvService, $ehvIntensity,
    );

    push @{ $model->{calc3Tables} }, $cdcmHvLvService, $cdcmEhvAssets,
      $cdcmHvLvShared
      if $model->{legacy201};

    ( $rateDirect, $rateRates, $rateIndirect ) =
      $model->{mitigateUndueSecrecy}
      ->fixedChargeAdj( $rateDirect, $rateRates, $rateIndirect,
        $demandSoleUseAsset, $chargeDirect, $chargeRates, $chargeIndirect, )
      if $model->{mitigateUndueSecrecy};

    my (
        $fixedDcharge,        $fixedDchargeTrue, $fixedGcharge,
        $fixedGchargeUnround, $fixedGchargeTrue,
      )
      = $model->fixedCharges(
        $rateDirect,     $rateRates,
        $daysInYear,     $demandSoleUseAsset,
        $dcp189Input,    $demandSoleUseAssetUnscaled,
        $importEligible, $generationSoleUseAsset,
        $generationSoleUseAssetUnscaled,
      );

    my $totalDcp189DiscountedAssets;

    $totalDcp189DiscountedAssets =
      $model->{transparencyMasterFlag}
      ? (
        $model->{dcp189} =~ /proportion/i
        ? Arithmetic(
            name => 'Total demand sole use assets '
              . 'qualifying for DCP 189 discount (£)',
            defaultFormat => '0softnz',
            arithmetic => '=IF(A123,0,A1)+SUMPRODUCT(A11_A12,A13_A14,A15_A16)',
            arguments  => {
                A123    => $model->{transparencyMasterFlag},
                A1      => $model->{transparency}{baselineItem}{119306},
                A11_A12 => $demandSoleUseAsset,
                A13_A14 => $dcp189Input,
                A15_A16 => $model->{transparency},
            },
          )
        : Arithmetic(
            name => 'Total demand sole use assets '
              . 'qualifying for DCP 189 discount (£)',
            defaultFormat => '0softnz',
            arithmetic    => '=IF(A123,0,A1)+SUMPRODUCT(A11_A12,A15_A16)',
            arguments     => {
                A123    => $model->{transparencyMasterFlag},
                A1      => $model->{transparency}{baselineItem}{119306},
                A11_A12 => Arithmetic(
                    name => 'Demand sole use assets '
                      . 'qualifying for DCP 189 discount (£)',
                    defaultFormat => '0softnz',
                    arithmetic    => '=IF(A4="Y",A1,0)',
                    arguments     => {
                        A1 => $demandSoleUseAsset,
                        A4 => $dcp189Input,
                    }
                ),
                A15_A16 => $model->{transparency},
            },
        )
      )
      : $model->{dcp189} =~ /proportion/i ? SumProduct(
        name => 'Total demand sole use assets '
          . 'qualifying for DCP 189 discount (£)',
        defaultFormat => '0softnz',
        matrix        => $dcp189Input,
        vector        => $demandSoleUseAsset,
      )
      : Arithmetic(
        name => 'Total demand sole use assets '
          . 'qualifying for DCP 189 discount (£)',
        defaultFormat => '0softnz',
        arithmetic    => '=SUMIF(A1_A2,"Y",A3_A4)',
        arguments     => {
            A1_A2 => $dcp189Input,
            A3_A4 => $demandSoleUseAsset,
        }
      ) if $model->{dcp189} && $model->{dcp189} =~ /preservePot|split/i;

    $model->{transparency}{dnoTotalItem}{119306} = $totalDcp189DiscountedAssets
      if $model->{transparency} && $totalDcp189DiscountedAssets;

    my $cdcmPurpleUse = Stack(
        cols => Labelset( list => [ $cdcmUse->{cols}{list}[0] ] ),
        name    => 'Total CDCM peak time consumption (kW)',
        sources => [$cdcmUse]
    );

    my ( $rateExit, $edcmPurpleUse ) =
      $model->exitChargingRate( $cdcmPurpleUse, $purpleUseRate, $importCapacity,
        $chargeExit, );

    ( $rateExit, $edcmPurpleUse ) =
      $model->{mitigateUndueSecrecy}->exitChargeAdj(
        $rateExit,   $cdcmPurpleUse, $edcmPurpleUse,
        $chargeExit, $purpleUseRate, $importCapacity,
      ) if $model->{mitigateUndueSecrecy};

    my ( $charges1, $acCoef, $reCoef ) =
      $model->charge1( $tariffLoc, $locations, $locParent, $c1, $a1d, $r1d,
        $a1g, $r1g,
        $model->preprocessLocationData( $locations, $a1d, $r1d, $a1g, $r1g, ),
      ) if $locations;

    my ( $fcpLricDemandCapacityChargeBig,
        $genCredit, $unitRateFcpLricNonDSM, $genCreditCapacity,
        $demandConsumptionFcpLric, )
      = $model->chargesFcpLric(
        $acCoef,                     $activeCoincidence,
        $charges1,                   $daysInYear,
        $reactiveCoincidence,        $reCoef,
        $tariffNetworkSupportFactor, $hoursInPurple,
        $hoursInPurple,              $importCapacity,
        $exportCapacityChargeable,   $creditableCapacity,
        $rateExit,                   $activeCoincidence935,
        $reactiveCoincidence935,
      );

    $genCredit = Arithmetic(
        name       => 'Generation credit (unrounded) p/kWh',
        groupName  => 'Generation unit rate credit',
        arithmetic => '=IF(A41,(A2*A3/(A4+A5)),0)',
        arguments  => {
            A1  => $exportCapacityChargeable,
            A2  => $genCredit,
            A21 => $genCredit,
            A3  => $exportCapacityChargeable,
            A4  => $exportCapacityChargeable,
            A5  => $exportCapacityExempt,
            A41 => $exportCapacityChargeable,
            A51 => $exportCapacityExempt,
        },
    );

    my $gCharge = $model->gCharge(
        $genPot20p,                        $genPotGP,
        $genPotGL,                         $genPotCdcmCap20052010,
        $genPotCdcmCapPost2010,            $exportCapacityChargeable,
        $exportCapacityChargeable20052010, $exportCapacityChargeablePost2010,
        $daysInYear,
    );

    $gCharge = $model->{mitigateUndueSecrecy}->gChargeAdj($gCharge)
      if $model->{mitigateUndueSecrecy};

    my ( $exportCapacityCharge, $genCreditRound, $exportCapacityChargeRound,
        $netexportCapacityChargeRound, $generationRevenue, )
      = $model->exportCharges( $gCharge, $daysInYear, $exportEligible,
        $genCredit, $genCreditCapacity, $exportCapacityChargeable, $activeUnits,
        $fixedGcharge, );

    my ( $totalRevenue3, $rateOther ) = $model->demandPot(
        $allowedRevenue,              $chargeDirect,
        $chargeRates,                 $chargeIndirect,
        $chargeExit,                  $rateDirect,
        $rateRates,                   $rateIndirect,
        $rateExit,                    $generationRevenue,
        $totalDcp189DiscountedAssets, $totalAssetsCapacity,
        $totalAssetsConsumption,      $cdcmEhvAssets,
        $cdcmHvLvShared,              $demandSoleUseAsset,
        $edcmPurpleUse,               $totalAssetsFixed,
        $daysInYear,                  $assetsCapacity,
        $assetsConsumption,           $purpleUseRate,
        $importCapacity,
    );

    my (
        $purpleRateFcpLricRound,    $fixedDchargeTrueRound,
        $importCapacityScaledRound, $importCapacityExceededRound,
        $exportCapacityExceeded,    $demandScalingShortfall,
      )
      = $model->tariffCalculation(
        $activeCoincidence,           $activeCoincidence935,
        $assetsCapacityCooked,        $assetsConsumptionCooked,
        $assetsFixed,                 $chargeableCapacity,
        $daysInYear,                  $demandConsumptionFcpLric,
        $edcmPurpleUse,               $exportCapacityCharge,
        $exportCapacityChargeRound,   $fcpLricDemandCapacityChargeBig,
        $fixedDcharge,                $fixedDchargeTrue,
        $fixedGchargeTrue,            $genCredit,
        $genCreditCapacity,           $hoursInPurple,
        $importCapacity,              $importCapacity935,
        $importEligible,              $indirectExposure,
        $powerFactorInModel,          $purpleUseRate,
        $rateDirect,                  $rateExit,
        $rateIndirect,                $rateRates,
        $reactiveCoincidence,         $tariffDaysInYearNot,
        $tariffHoursInPurpleNot,      $tariffs,
        $totalAssetsCapacity,         $totalAssetsConsumption,
        $totalAssetsFixed,            $totalAssetsGenerationSoleUse,
        $totalDcp189DiscountedAssets, $totalEdcmAssets,
        $totalRevenue3,               $unitRateFcpLricNonDSM,
      );

    my @tariffColumns = (
        Stack( sources => [$tariffs] ),
        $purpleRateFcpLricRound,
        $fixedDchargeTrueRound,
        $importCapacityScaledRound,
        $importCapacityExceededRound,
        Stack( sources => [$genCreditRound] ),
        Stack( sources => [$fixedGchargeTrue] ),
        Stack( sources => [$netexportCapacityChargeRound] ),
        $exportCapacityExceeded,
    );

    my $allTariffColumns = $model->{checksums}
      ? [
        @tariffColumns,
        map {
            my $digits = /([0-9])/ ? $1 : 6;
            my $recursive = /table|recursive|model/i ? 1 : 0;
            $model->{"checksum_${recursive}_$digits"} =
              SpreadsheetModel::Checksum->new(
                name => $_,
                /table|recursive|model/i ? ( recursive => 1 ) : (),
                digits  => $digits,
                columns => [ @tariffColumns[ 1 .. 8 ] ],
                factors => [qw(1000 100 100 100 1000 100 100 100)]
              );
          } split /;\s*/,
        $model->{checksums}
      ]
      : \@tariffColumns;

    if ( $model->{layout} ) {

        $model->{tableList} = $model->orderedLayout(
            [ $exportEligible, @{ $tariffColumns[5]{sourceLines} } ],
            $model->{transparencyMasterFlag}
            ? [ $model->{transparencyMasterFlag}, ]
            : (),
            1 || $model->{layout} =~ /multisheet/i
            ? ( [ $gCharge, ], [ $rateExit, ], )
            : [ $gCharge, $rateExit, ],
            [
                @{ $tariffColumns[7]{sourceLines} },
                @{ $tariffColumns[8]{sourceLines} },
            ],
            [ $generationSoleUseAsset, $demandSoleUseAsset, ],
            [ $accretion, ],
            [ $assetsCapacity, ],
            [ $assetsConsumption, ],
            [ $rateDirect, $rateIndirect, $rateRates, ],
            [
                @{ $tariffColumns[6]{sourceLines} },
                @{ $tariffColumns[2]{sourceLines} },
                $fixedGcharge,
            ],
            [ $rateOther, ],
            [ $demandScalingShortfall, ],
            [ $assetsCapacityCooked, ],
            [ $assetsConsumptionCooked, ],
            [
                @{ $tariffColumns[1]{sourceLines} },
                @{ $tariffColumns[3]{sourceLines} },
                @{ $tariffColumns[4]{sourceLines} },
            ],
            $model->{mitigateUndueSecrecy}
            ? $model->{mitigateUndueSecrecy}->calcTables
            : (),
        );

    }

    if (    $model->{layout}
        and $model->{layout} =~ /no4501/i || $model->{tariff1Row} )
    {

        SpreadsheetModel::MatrixSheet->new(
            $model->{tariff1Row}
            ? (
                dataRow            => $model->{tariff1Row},
                captionDecorations => [qw(algae purple slime)],
              )
            : (),
          )->addDatasetGroup(
            name    => 'Tariff name',
            columns => [ $allTariffColumns->[0] ],
          )->addDatasetGroup(
            name    => 'Import tariff',
            columns => [ @{$allTariffColumns}[ 1 .. 4 ] ],
          )->addDatasetGroup(
            name    => 'Export tariff',
            columns => [ @{$allTariffColumns}[ 5 .. 8 ] ],
          )->addDatasetGroup(
            name    => 'Checksums',
            columns => [ @{$allTariffColumns}[ 9 .. $#$allTariffColumns ] ]
          );

        push @{ $model->{tariffTables} }, @$allTariffColumns;

    }

    else {
        push @{ $model->{tariffTables} },
          Columnset(
            name    => 'EDCM charge',
            columns => $allTariffColumns,
          );
    }

    if ( $model->{transparency} ) {

        my %dnoTotalItem;
        foreach (
            grep { $_ > 99_999; }
            keys %{ $model->{transparency}{dnoTotalItem} }
          )
        {
            my $number = int( $_ / 100 );
            $dnoTotalItem{$number}[ $_ - $number * 100 - 1 ] =
              Stack( sources => [ $model->{transparency}{dnoTotalItem}{$_} ] );
        }

        foreach ( values %dnoTotalItem ) {
            $_ ||= Constant( name => 'Not used', data => [], ) foreach @$_;
        }

        $model->{mitigateUndueSecrecy}
          ->adjustDnoTotals( $model, \%dnoTotalItem )
          if $model->{mitigateUndueSecrecy};

        $model->{aggregateTables} = [
            (
                map {
                    Columnset(
                        name          => "⇒$_->[0]. $_->[1]",
                        singleRowName => $_->[1],
                        number        => 3600 + $_->[0],
                        columns       => $dnoTotalItem{ $_->[0] },
                      )
                  }[ 1191 => 'EDCM demand and revenue aggregates' ],
                [ 1192 => 'EDCM generation aggregates' ],
                [ 1193 => 'EDCM notional asset aggregates' ],
            ),
            $model->{mitigateUndueSecrecy} ? () : (
                map {
                    my $obj  = $model->{transparency}{dnoTotalItem}{$_};
                    my $name = 'Copy of ' . $obj->{name};
                    $obj->isa('SpreadsheetModel::Columnset')
                      ? Columnset(
                        name    => $name,
                        number  => 3600 + $_,
                        columns => [
                            map { Stack( sources => [$_] ) }
                              @{ $obj->{columns} }
                        ]
                      )
                      : Stack(
                        name    => $name,
                        number  => 3600 + $_,
                        sources => [$obj]
                      );
                  } sort { $a <=> $b; }
                  grep   { $_ < 100_000; }
                  keys %{ $model->{transparency}{dnoTotalItem} }
            )
        ];

    }

    $model->summaries(
        $activeCoincidence935,      $activeUnits,
        $actualRedDemandRate,       $daysInYear,
        $exportCapacityChargeable,  $fixedDchargeTrue,
        $fixedDchargeTrueRound,     $fixedGchargeTrue,
        $genCreditRound,            $hoursInPurple,
        $importCapacity,            $importCapacity935,
        $importCapacityScaledRound, $netexportCapacityChargeRound,
        $previousChargeExport,      $previousChargeImport,
        $purpleRateFcpLricRound,    $tariffDaysInYearNot,
        $tariffHoursInPurpleNot,    $tariffs,
        @tariffColumns,
    ) if $model->{summaries};

    $model->templates(
        $tariffs,                          $importCapacity935,
        $exportCapacityExempt,             $exportCapacityChargeablePre2005,
        $exportCapacityChargeable20052010, $exportCapacityChargeablePost2010,
        $tariffSoleUseMeav,                $tariffLoc,
        $tariffCategory,                   $useProportions,
        $activeCoincidence935,             $reactiveCoincidence,
        $indirectExposure,                 $nonChargeableCapacity935,
        $activeUnits,                      $creditableCapacity,
        $tariffNetworkSupportFactor,       $tariffDaysInYearNot,
        $tariffHoursInPurpleNot,           $previousChargeImport,
        $previousChargeExport,             $llfcImport,
        $llfcExport,                       \@tariffColumns,
        $daysInYear,                       $hoursInPurple,
    ) if $model->{customerTemplates};

    $model;

}

1;
