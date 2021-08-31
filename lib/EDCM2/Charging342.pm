package EDCM2;

# Copyright 2009-2012 Energy Networks Association Limited and others.
# Copyright 2013-2020 Franck Latrémolière and others.
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

sub tariffCalculation342 {

    my (
        $model,                          $activeCoincidence,
        $activeCoincidence935,           $assetsCapacityCooked,
        $assetsConsumptionCooked,        $assetsFixed,
        $chargeableCapacity,             $daysInYear,
        $demandConsumptionFcpLric,       $edcmPurpleUse,
        $exportCapacityCharge,           $exportCapacityChargeRound,
        $fcpLricDemandCapacityChargeBig, $fixedDchargeNotUsed,
        $fixedDchargeTrue,               $fixedGchargeTrue,
        $genCredit,                      $genCreditCapacity,
        $hoursInPurple,                  $importCapacity,
        $importCapacity935,              $importEligible,
        $indirectExposure,               $powerFactorInModel,
        $purpleUseRate,                  $rateDirect,
        $rateExit,                       $rateIndirect,
        $rateRates,                      $reactiveCoincidence,
        $tariffDaysInYearNot,            $tariffHoursInPurpleNot,
        $tariffs,                        $totalAssetsCapacity,
        $totalAssetsConsumption,         $totalAssetsFixed,
        $totalAssetsGenerationSoleUse,   $totalDcp189DiscountedAssets,
        $totalEdcmAssets,                $totalRevenue3,
        $unitRateFcpLricNonDSM,          $tariffScalingExempt,
    ) = @_;

    my $capacityCharge = Arithmetic(
        name          => 'Capacity charge p/kVA/day (exit only)',
        defaultFormat => '0.00soft',
        arithmetic    => '=100/A2*A41*A1',
        arguments     => {
            A2  => $daysInYear,
            A41 => $rateExit,
            A1  => ref $purpleUseRate eq 'ARRAY'
            ? $purpleUseRate->[0]
            : $purpleUseRate,
        }
    );

    $model->{summaryInformationColumns}[1] = Arithmetic(
        name          => 'Transmission exit charge (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=0.01*A9*A1*A2',
        arguments     => {
            A1 => $importCapacity,
            A2 => $capacityCharge,
            A9 => $daysInYear,
            A7 => $tariffDaysInYearNot,
        },
    );

    my $importCapacityExceededAdjustment = Arithmetic(
        name =>
          'Adjustment to exceeded import capacity charge for DSM (p/kVA/day)',
        defaultFormat => '0.00soft',
        arithmetic    => '=IF(A1=0,0,(1-A4/A5)*(A3+'
          . 'IF(A23=0,0,(A2*A21*(A22-A24)/(A9-A91)))))',
        arguments => {
            A3  => $fcpLricDemandCapacityChargeBig,
            A4  => $chargeableCapacity,
            A5  => $importCapacity,
            A1  => $importCapacity,
            A2  => $unitRateFcpLricNonDSM,
            A21 => $activeCoincidence935,
            A23 => $activeCoincidence935,
            A22 => $hoursInPurple,
            A24 => $tariffHoursInPurpleNot,
            A9  => $daysInYear,
            A91 => $tariffDaysInYearNot,
        },
        defaultFormat => '0.00soft'
    );

    push @{ $model->{calc2Tables} },
      my $unitRateFcpLricDSM = Arithmetic(
        name => "$model->{TimebandName} unit rate adjusted for DSM (p/kWh)",
        arithmetic => '=IF(A6=0,1,A4/A5)*A1',
        arguments  => {
            A1 => $unitRateFcpLricNonDSM,
            A4 => $chargeableCapacity,
            A5 => $importCapacity,
            A6 => $importCapacity,
        }
      );

    push @{ $model->{calc2Tables} },
      my $capacityChargeT1 = Arithmetic(
        name          => 'Import capacity charge from charge 1 (p/kVA/day)',
        groupName     => 'Charge 1',
        arithmetic    => '=IF(A6=0,1,A4/A5)*A1',
        defaultFormat => '0.00soft',
        arguments     => {
            A1 => $fcpLricDemandCapacityChargeBig,
            A4 => $chargeableCapacity,
            A5 => $importCapacity,
            A6 => $importCapacity,
        },
      );

    $capacityCharge = Arithmetic(
        name          => 'Import capacity charge before scaling (p/kVA/day)',
        arithmetic    => '=A7+A1',
        defaultFormat => '0.00soft',
        arguments     => {
            A1 => $capacityChargeT1,
            A7 => $capacityCharge,
        }
    );

    my $tariffHoursInPurple = Arithmetic(
        name => "Number of $model->{timebandName} hours connected in year",
        defaultFormat => '0.0soft',
        arithmetic    => '=A2-A1',
        arguments     => {
            A2 => $hoursInPurple,
            A1 => $tariffHoursInPurpleNot,
        }
    );

    my $demandScalingShortfall = Arithmetic(
        name          => 'Additional amount to be recovered (£/year)',
        groupName     => 'Residual EDCM demand revenue',
        defaultFormat => '0soft',
        arithmetic    => '=A1-A2*'
          . (
            $totalDcp189DiscountedAssets ? '(A42-A44)'
            : 'A42'
          )
          . '-A3*A43-A5*A6'
          . ( $model->{removeDemandCharge1} ? '' : '-A9' ),
        arguments => {
            A1  => $totalRevenue3,
            A2  => $rateDirect,
            A3  => $rateRates,
            A42 => $totalAssetsFixed,
            A43 => $totalAssetsFixed,
            $totalDcp189DiscountedAssets
            ? ( A44 => $totalDcp189DiscountedAssets )
            : (),
            A5 => $rateExit,
            A6 => $edcmPurpleUse,
            $model->{removeDemandCharge1} ? ()
            : (
                A9 => $model->{transparencyMasterFlag} ? Arithmetic(
                    name          => 'Revenue from demand charge 1 (£/year)',
                    defaultFormat => '0soft',
                    arithmetic    => '=IF(A123,0,A1)+('
                      . 'SUMPRODUCT(A64_A65,A31_A32,A33_A34)+'
                      . 'SUMPRODUCT(A66_A67,A41_A42,A43_A44,A35_A36,A51_A52)/A54'
                      . ')*A9/100',
                    arguments => {
                        A123    => $model->{transparencyMasterFlag},
                        A1      => $model->{transparency}{baselineItem}{119104},
                        A31_A32 => $capacityChargeT1,
                        A33_A34 => $importCapacity,
                        A9      => $daysInYear,
                        A41_A42 => $unitRateFcpLricDSM,
                        A43_A44 => $activeCoincidence935,
                        A35_A36 => $importCapacity935,
                        A51_A52 => $tariffHoursInPurple,
                        A54     => $daysInYear,
                        A64_A65 => $model->{transparency},
                        A66_A67 => $model->{transparency},
                    },
                  )
                : Arithmetic(
                    name          => 'Revenue from demand charge 1 (£/year)',
                    defaultFormat => '0soft',
                    arithmetic    => '=('
                      . 'SUMPRODUCT(A31_A32,A33_A34)+'
                      . 'SUMPRODUCT(A41_A42,A43_A44,A35_A36,A51_A52)/A54'
                      . ')*A9/100',
                    arguments => {
                        A31_A32 => $capacityChargeT1,
                        A33_A34 => $importCapacity,
                        A9      => $daysInYear,
                        A41_A42 => $unitRateFcpLricDSM,
                        A43_A44 => $activeCoincidence935,
                        A35_A36 => $importCapacity935,
                        A51_A52 => $tariffHoursInPurple,
                        A54     => $daysInYear,
                    },
                ),
            ),
        },
    );

    $model->{transparency}{dnoTotalItem}{1254} = $demandScalingShortfall
      if $model->{transparency};
    $model->{transparency}{dnoTotalItem}{119104} =
      $demandScalingShortfall->{arguments}{A9}
      if $model->{transparency}
      && $demandScalingShortfall->{arguments}{A9};

    my $edcmIndirect = Arithmetic(
        name          => 'Indirect costs on EDCM demand (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=A1*(A20-A23)',
        arguments     => {
            A1  => $rateIndirect,
            A20 => $totalEdcmAssets,
            A23 => $totalAssetsGenerationSoleUse,
        },
    );
    $model->{transparency}{dnoTotalItem}{1253} = $edcmIndirect
      if $model->{transparency};

    my $edcmDirect = Arithmetic(
        name => 'Direct costs on EDCM demand except'
          . ' through sole use asset charges (£/year)',
        groupName     => 'Expenditure allocated to EDCM demand',
        defaultFormat => '0soft',
        arithmetic    => '=A1*(A20+A23)',
        arguments     => {
            A1  => $rateDirect,
            A20 => $totalAssetsCapacity,
            A23 => $totalAssetsConsumption,
        },
    );
    $model->{transparency}{dnoTotalItem}{1252} = $edcmDirect
      if $model->{transparency};

    my $edcmRates = Arithmetic(
        name => 'Network rates on EDCM demand except '
          . 'through sole use asset charges (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=A1*(A20+A23)',
        arguments     => {
            A1  => $rateRates,
            A20 => $totalAssetsCapacity,
            A23 => $totalAssetsConsumption,
        },
    );
    $model->{transparency}{dnoTotalItem}{1255} = $edcmRates
      if $model->{transparency};

    my $adderAmount = Arithmetic(
        name          => 'Amount to be recovered from adders ex costs (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=A1-A7-A91-A92',
        arguments     => {
            A1  => $demandScalingShortfall,
            A7  => $edcmIndirect,
            A91 => $edcmDirect,
            A92 => $edcmRates,
        },
    );

    my $ynonFudge = Constant(
        name => 'Factor for the allocation of capacity scaling',
        data => [0.5],
    );

    my $ynonFudge41 = Constant(
        name => 'Proportion of residual to go into fixed adder',
        data => [0.2],
    );

    $model->{mitigateUndueSecrecy}->fudge41param($ynonFudge41)
      if $model->{mitigateUndueSecrecy};

    $activeCoincidence = Arithmetic(
        name       => 'Peak-time capacity use per kVA of agreed capacity',
        arithmetic => '=SQRT(A1*A2+A3*A4)',
        arguments  => {
            A1 => $activeCoincidence,
            A2 => $activeCoincidence,
            A3 => $reactiveCoincidence,
            A4 => $reactiveCoincidence,
        }
    ) if $model->{dcp183};

    my $slopeFixedAdder =
      $model->{dcp185}
      ? Arithmetic(
        name => 'Marginal revenue effect of demand'
          . ( $model->{dcp185} == 2 ? ' and indirect cost adders' : ' adder' ),
        defaultFormat => '0soft',
        arithmetic    => '=A1*(A4+A5)*IF(A6<0,1,A7)'
          . ( $model->{dcp342} ? '*(1-A9)' : '' ),
        arguments => {
            A1 => $importCapacity,
            A4 => $ynonFudge,
            A5 => $activeCoincidence,
            A6 => $adderAmount,
            A7 => $indirectExposure,
            $model->{dcp342} ? ( A9 => $tariffScalingExempt ) : (),
        }
      )
      : Arithmetic(
        name          => 'Marginal revenue effect of demand adder',
        defaultFormat => '0soft',
        arithmetic    => '=A1*(A4+A5)' . ( $model->{dcp342} ? '*(1-A9)' : '' ),
        arguments     => {
            A1 => $importCapacity,
            A4 => $ynonFudge,
            A5 => $activeCoincidence,
            $model->{dcp342} ? ( A9 => $tariffScalingExempt ) : (),
        }
      );

    my $slopeIndirect =
      $model->{dcp185} && $model->{dcp185} == 2
      ? Arithmetic(
        name       => 'Data for capacity-based allocation of indirect costs',
        groupName  => 'Allocation of indirect costs',
        arithmetic => '=IF(A6<0,1,A1)*(A2+A3)',
        arguments  => {
            A2 => $ynonFudge,
            A3 => $activeCoincidence,
            A1 => $indirectExposure,
            A6 => $adderAmount,
        }
      )
      : Arithmetic(
        name       => 'Data for capacity-based allocation of indirect costs',
        groupName  => 'Allocation of indirect costs',
        arithmetic => '=A1*(A2+A3)',
        arguments  => {
            A2 => $ynonFudge,
            A3 => $activeCoincidence,
            A1 => $indirectExposure,
        }
      );

    my $totalIndirectFudge =
      $model->{transparencyMasterFlag}
      ? Arithmetic(
        name          => 'Total marginal effect of indirect cost adder',
        defaultFormat => '0soft',
        arithmetic    => '=IF(A123,0,A1)+SUMPRODUCT(A2_A3,A4_A5,A6_A7)',
        arguments     => {
            A123  => $model->{transparencyMasterFlag},
            A1    => $model->{transparency}{baselineItem}{119102},
            A2_A3 => $model->{transparency},
            A4_A5 => $slopeIndirect,
            A6_A7 => $importCapacity,
        },
      )
      : SumProduct(
        name          => 'Total marginal effect of indirect cost adder',
        defaultFormat => '0soft',
        matrix        => $slopeIndirect,
        vector        => $importCapacity
      );

    $model->{transparency}{dnoTotalItem}{119102} = $totalIndirectFudge
      if $model->{transparency};

    my $indirectAppRate = Arithmetic(
        name       => 'Indirect costs application rate',
        groupName  => 'EDCM demand adders',
        arithmetic => '=IF(A2,A3/A4,0)',
        arguments  => {
            A2 => $edcmIndirect,
            A3 => $edcmIndirect,
            A4 => $totalIndirectFudge,
        },
        location => 'Charging rates',
    );
    $model->{transparency}{dnoTotalItem}{1262} = $indirectAppRate
      if $model->{transparency};

    ($indirectAppRate) =
      $model->{mitigateUndueSecrecy}
      ->indirectChargeAdj( $indirectAppRate, $slopeIndirect,
        $importCapacity, $edcmIndirect, )
      if $model->{mitigateUndueSecrecy};

    $capacityCharge = Arithmetic(
        arithmetic => '=A1+A3*A4*100/A9',
        name => 'Capacity charge after applying indirect cost charge p/kVA/day',
        arguments => {
            A1 => $capacityCharge,
            A3 => $indirectAppRate,
            A4 => $slopeIndirect,
            A9 => $daysInYear,
        }
    );

    $model->{summaryInformationColumns}[3] = Arithmetic(
        name          => 'Indirect cost allocation (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=A1*A3*A7',
        arguments     => {
            A1 => $importCapacity,
            A3 => $indirectAppRate,
            A7 => $slopeIndirect,
        },
    );

    my $totalSlope =
      $model->{transparencyMasterFlag}
      ? Arithmetic(
        name          => 'Total marginal revenue effect of demand adder',
        defaultFormat => '0soft',
        arithmetic    => '=IF(A123,0,A1)+SUMPRODUCT(A2_A3,A4_A5)',
        arguments     => {
            A123  => $model->{transparencyMasterFlag},
            A1    => $model->{transparency}{baselineItem}{119103},
            A2_A3 => $model->{transparency},
            A4_A5 => $slopeFixedAdder,
        },
      )
      : GroupBy(
        name          => 'Total marginal revenue effect of demand adder',
        defaultFormat => '0soft',
        source        => $slopeFixedAdder
      );

    $model->{transparency}{dnoTotalItem}{119103} = $totalSlope
      if $model->{transparency};

    my $fixedAdderRate = Arithmetic(
        name       => 'Fixed adder ex indirects application rate',
        groupName  => 'EDCM demand adders',
        arithmetic => '=IF(A9,A1*A2/A4,0)',
        arguments  => {
            A1 => $ynonFudge41,
            A2 => $adderAmount,
            A9 => $adderAmount,
            A4 => $totalSlope,
        },
        location => 'Charging rates',
    );
    $model->{transparency}{dnoTotalItem}{1261} = $fixedAdderRate
      if $model->{transparency};

    ($fixedAdderRate) =
      $model->{mitigateUndueSecrecy}
      ->fixedAdderAdj( $fixedAdderRate, $slopeFixedAdder )
      if $model->{mitigateUndueSecrecy};

    $capacityCharge = Arithmetic(
        arithmetic => '=A1+A3*(A7+A4)*100/A9'
          . ( $model->{dcp185} ? '*IF(A6<0,1,A8)' : '' )
          . ( $model->{dcp342} ? '*(1-A666)'      : '' ),
        name =>
          'Capacity charge after applying fixed adder ex indirects p/kVA/day',
        arguments => {
            A1 => $capacityCharge,
            A3 => $fixedAdderRate,
            A4 => $activeCoincidence,
            A7 => $ynonFudge,
            A9 => $daysInYear,
            $model->{dcp185}
            ? (
                A6 => $adderAmount,
                A8 => $indirectExposure,
              )
            : (),
            $model->{dcp342} ? ( A666 => $tariffScalingExempt ) : (),
        }
    );

    $model->{summaryInformationColumns}[6] = Arithmetic(
        name          => 'Demand scaling fixed adder (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=A1*A3*(A71+A72)'
          . ( $model->{dcp185} ? '*IF(A6<0,1,A8)' : '' )
          . ( $model->{dcp342} ? '*(1-A9)'        : '' ),
        arguments => {
            A1  => $importCapacity,
            A3  => $fixedAdderRate,
            A71 => $ynonFudge,
            A72 => $activeCoincidence,
            $model->{dcp185}
            ? (
                A6 => $adderAmount,
                A8 => $indirectExposure,
              )
            : (),
            $model->{dcp342} ? ( A9 => $tariffScalingExempt ) : (),
        },
    );

    my $slopeNotionalAssets = Arithmetic(
        name      => 'Non sole use notional assets subject to allocation (£)',
        groupName => 'Demand scaling',
        defaultFormat => '0soft',
        arithmetic    => '=(A2+A1)*A4',
        arguments     => {
            A2 => $assetsCapacityCooked,
            A1 => $assetsConsumptionCooked,
            A4 => $importCapacity,
        }
    );

    my $totalSlopeNotionalAssets =
      $model->{transparencyMasterFlag}
      ? Arithmetic(
        name => 'Total non sole use notional assets subject to allocation (£)',
        defaultFormat => '0soft',
        arithmetic    => '=IF(A123,0,A1)+SUMPRODUCT(A21_A22,A51_A52)',
        arguments     => {
            A123    => $model->{transparencyMasterFlag},
            A1      => $model->{transparency}{baselineItem}{119305},
            A21_A22 => $model->{transparency},
            A51_A52 => $slopeNotionalAssets,
        }
      )
      : GroupBy(
        name => 'Total non sole use notional assets subject to allocation (£)',
        defaultFormat => '0soft',
        source        => $slopeNotionalAssets,
      );

    $model->{transparency}{dnoTotalItem}{119305} = $totalSlopeNotionalAssets
      if $model->{transparency};

    my $totalSlopeNotionalAssetsExempt = $totalSlopeNotionalAssets;

    if ( $model->{dcp342} ) {
        $totalSlopeNotionalAssetsExempt =
          $model->{transparencyMasterFlag}
          ? Arithmetic(
            name => 'Total non sole use notional assets'
              . ' subject to DCP 342 exemption (£)',
            defaultFormat => '0soft',
            arithmetic => '=IF(A123,0,A1)+SUMPRODUCT(A21_A22,A51_A52,A61_A62)',
            arguments  => {
                A123    => $model->{transparencyMasterFlag},
                A1      => $model->{transparency}{baselineItem}{119307},
                A21_A22 => $model->{transparency},
                A51_A52 => $slopeNotionalAssets,
                A61_A62 => $tariffScalingExempt,
            }
          )
          : GroupBy(
            name =>
              'Total non sole use notional assets subject to allocation (£)',
            defaultFormat => '0soft',
            source        => $slopeNotionalAssets,
          );
        $model->{transparency}{dnoTotalItem}{119307} =
          $totalSlopeNotionalAssetsExempt
          if $model->{transparency};
    }

    die 'This code does not support mitigateUndueSecrecy'
      if $model->{mitigateUndueSecrecy};

    my $directChargingRate = Arithmetic(
        name          => 'Charging rate for direct costs on notional assets',
        groupName     => 'Demand scaling rate',
        defaultFormat => '%soft',
        arithmetic    => '=IF(A4,A1/A2,0)',
        arguments     => {
            A1 => $edcmDirect,
            A4 => $totalSlopeNotionalAssets,
            A2 => $totalSlopeNotionalAssets,
        },
        location => 'Charging rates',
    );
    $model->{transparency}{dnoTotalItem}{1264} = $directChargingRate
      if $model->{transparency};

    my $ratesChargingRate = Arithmetic(
        name          => 'Charging rate for network rates on notional assets',
        groupName     => 'Demand scaling rate',
        defaultFormat => '%soft',
        arithmetic    => '=IF(A4,A1/A2,0)',
        arguments     => {
            A1 => $edcmRates,
            A4 => $totalSlopeNotionalAssets,
            A2 => $totalSlopeNotionalAssets,
        },
        location => 'Charging rates',
    );
    $model->{transparency}{dnoTotalItem}{1265} = $ratesChargingRate
      if $model->{transparency};

    my $assetScalerChargingRate = Arithmetic(
        name          => 'Charging rate for asset scaler on notional assets',
        groupName     => 'Demand scaling rate',
        defaultFormat => '%soft',
        arithmetic    => '=IF(A4,(1-A3)*A1/'
          . ( $model->{dcp342} ? '(A2-A5)' : 'A2' ) . ',0)',
        arguments => {
            A1 => $adderAmount,
            A3 => $ynonFudge41,
            A4 => $totalSlopeNotionalAssets,
            A2 => $totalSlopeNotionalAssets,
            $model->{dcp342} ? ( A5 => $totalSlopeNotionalAssetsExempt ) : (),
        },
        location => 'Charging rates',
    );
    $model->{transparency}{dnoTotalItem}{1266} = $assetScalerChargingRate
      if $model->{transparency};

    my $scalingChargeCapacity = Arithmetic(
        arithmetic => '=(A2+A3+A4'
          . ( $model->{dcp342} ? '*(1-A5)' : '' )
          . ')*(A1+A12)*100/A9',
        defaultFormat => '0.00soft',
        name      => 'Notional asset charge including demand scaling p/kVA/day',
        arguments => {
            A1  => $assetsCapacityCooked,
            A12 => $assetsConsumptionCooked,
            A2  => $directChargingRate,
            A3  => $ratesChargingRate,
            A4  => $assetScalerChargingRate,
            $model->{dcp342} ? ( A5 => $tariffScalingExempt ) : (),
            A9 => $daysInYear,
        }
    );

    my $scalingChargeApplicationFactor = Arithmetic(
        name       => 'Application factor for charges on notional assets',
        arithmetic => '=IF(A1<0,MIN(1,'
          . '0-(A21+(1-A55/A54)/(1-A56/A31)*IF(A52=0,1,A51/A53)*A5)'
          . '/A2),1)',
        arguments => {
            A1  => $scalingChargeCapacity,
            A2  => $scalingChargeCapacity,
            A21 => $capacityCharge,
            A31 => $daysInYear,
            A5  => $demandConsumptionFcpLric,
            A51 => $chargeableCapacity,
            A52 => $importCapacity,
            A53 => $importCapacity,
            A54 => $hoursInPurple,
            A55 => $tariffHoursInPurpleNot,
            A56 => $tariffDaysInYearNot,
        },
    );

    $model->{summaryInformationColumns}[2] = Arithmetic(
        name          => 'Direct cost allocation (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=A1*(A11+A12)*A4*A2',
        arguments     => {
            A1  => $importCapacity,
            A11 => $assetsCapacityCooked,
            A12 => $assetsConsumptionCooked,
            A2  => $scalingChargeApplicationFactor,
            A4  => $directChargingRate,
        },
    );

    $model->{summaryInformationColumns}[4] = Arithmetic(
        name          => 'Network rates allocation (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=A1*(A11+A12)*A4*A2',
        arguments     => {
            A1  => $importCapacity,
            A11 => $assetsCapacityCooked,
            A12 => $assetsConsumptionCooked,
            A2  => $scalingChargeApplicationFactor,
            A4  => $ratesChargingRate,
        },
    );

    $model->{summaryInformationColumns}[7] = Arithmetic(
        name          => 'Demand scaling asset based (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=A1*(A11+A12)*A4*A2'
          . ( $model->{dcp342} ? '*(1-A5)' : '' ),
        arguments => {
            A1  => $importCapacity,
            A11 => $assetsCapacityCooked,
            A12 => $assetsConsumptionCooked,
            A2  => $scalingChargeApplicationFactor,
            A4  => $assetScalerChargingRate,
            $model->{dcp342} ? ( A5 => $tariffScalingExempt ) : (),
        },
    );

    my $importCapacityScaled = Arithmetic(
        name          => 'Total import capacity charge p/kVA/day',
        defaultFormat => '0.00soft',
        arithmetic    => '=A1+A2*A3',
        arguments     => {
            A1 => $capacityCharge,
            A2 => $scalingChargeCapacity,
            A3 => $scalingChargeApplicationFactor,
        }
    );

    my $purpleRateFcpLric = Arithmetic(
        name       => "$model->{TimebandName} rate p/kWh",
        arithmetic => '=IF(A3,IF(A1=0,A9,'
          . 'MAX(0,MIN(A4,A41+(A5/A11*(A7-A71)/(A8-A81))))' . '),0)',
        arguments => {
            A1  => $activeCoincidence,
            A11 => $activeCoincidence935,
            A3  => $importEligible,
            A4  => $unitRateFcpLricDSM,
            A41 => $unitRateFcpLricDSM,
            A9  => $unitRateFcpLricDSM,
            A5  => $importCapacityScaled,
            A7  => $daysInYear,
            A71 => $tariffDaysInYearNot,
            A8  => $hoursInPurple,
            A81 => $tariffHoursInPurpleNot,
        }
    ) if $unitRateFcpLricDSM;

    push @{ $model->{calc4Tables} }, $importCapacityScaled;

    $importCapacityScaled = Arithmetic(
        name       => 'Import capacity charge p/kVA/day',
        groupName  => 'Demand charges after scaling',
        arithmetic => '=IF(A3,MAX(0,A1),0)',
        arguments  => {
            A1 => $importCapacityScaled,
            A3 => $importEligible,
        },
        defaultFormat => '0.00soft'
    );

    my $importCapacityExceeded = Arithmetic(
        name          => 'Exceeded import capacity charge (p/kVA/day)',
        defaultFormat => '0.00soft',
        arithmetic    => '=A7+A2',
        defaultFormat => '0.00soft',
        arguments     => {
            A3 => $fcpLricDemandCapacityChargeBig,
            A2 => $importCapacityExceededAdjustment,
            A4 => $chargeableCapacity,
            A5 => $importCapacity,
            A1 => $importCapacity,
            A7 => $importCapacityScaled,
        },
        defaultFormat => '0.00soft'
    );

    $model->{summaryInformationColumns}[5] = Arithmetic(
        name          => 'FCP/LRIC charge (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=0.01*(A11*A9*A2+A1*A4*A8*(A6-A61)*(A91/(A92-A71)))',
        arguments     => {
            A1  => $importCapacity,
            A2  => $fcpLricDemandCapacityChargeBig,
            A3  => $capacityCharge->{arguments}{A1},
            A9  => $daysInYear,
            A4  => $unitRateFcpLricDSM,
            A41 => $activeCoincidence,
            A6  => $hoursInPurple,
            A61 => $tariffHoursInPurpleNot,
            A8  => $activeCoincidence935,
            A91 => $daysInYear,
            A92 => $daysInYear,
            A71 => $tariffDaysInYearNot,
            A11 => $chargeableCapacity,
            A51 => $importCapacity,
            A62 => $importCapacity,

        },
    );

    push @{ $model->{tablesG} }, $genCredit, $genCreditCapacity,
      $exportCapacityCharge;

    my $fixedDchargeTrueRound = Arithmetic(
        name          => 'Import fixed charge (p/day)',
        defaultFormat => '0.00soft',
        arithmetic    => '=ROUND(A1,2)',
        arguments     => { A1 => $fixedDchargeTrue, },
    );

    my $purpleRateFcpLricRound = Arithmetic(
        name          => "Import $model->{timebandName} unit rate (p/kWh)",
        defaultFormat => '0.000soft',
        arithmetic    => '=ROUND(A1,3)',
        arguments     => { A1 => $purpleRateFcpLric, },
    );

    my $importCapacityScaledRound = Arithmetic(
        name          => 'Import capacity rate (p/kVA/day)',
        defaultFormat => '0.00soft',
        arithmetic    => '=ROUND(A1,2)',
        arguments     => { A1 => $importCapacityScaled, },
    );

    my $exportCapacityExceeded = Arithmetic(
        name          => 'Export exceeded capacity rate (p/kVA/day)',
        defaultFormat => '0.00copy',
        arithmetic    => '=A1',
        arguments     => { A1 => $exportCapacityChargeRound, },
    );

    my $importCapacityExceededRound = Arithmetic(
        name          => 'Import exceeded capacity rate (p/kVA/day)',
        defaultFormat => '0.00soft',
        arithmetic    => '=ROUND(A1,2)',
        arguments     => { A1 => $importCapacityExceeded, },
    );

    push @{ $model->{calc4Tables} }, $purpleRateFcpLric,
      $importCapacityScaled,
      $fixedDchargeTrue, $importCapacityExceeded,
      $exportCapacityChargeRound,
      $fixedGchargeTrue;

    $purpleRateFcpLricRound,      $fixedDchargeTrueRound,
      $importCapacityScaledRound, $importCapacityExceededRound,
      $exportCapacityExceeded,    $demandScalingShortfall;

}

1;
