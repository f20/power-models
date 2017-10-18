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

sub tariffCalculation {

    my (
        $model,                          $activeCoincidence,
        $activeCoincidence935,           $assetsCapacityCooked,
        $assetsConsumptionCooked,        $assetsFixed,
        $chargeableCapacity,             $daysInYear,
        $demandConsumptionFcpLric,       $edcmPurpleUse,
        $exportCapacityCharge,           $exportCapacityChargeRound,
        $fcpLricDemandCapacityChargeBig, $fixedDcharge,
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
        $unitRateFcpLricNonDSM,
    ) = @_;

    my ( $scalingChargeCapacity, $scalingChargeUnits );

    my $capacityChargeT = Arithmetic(
        name          => 'Capacity charge p/kVA/day (exit only)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=100/A2*A41*A1',
        arguments     => {
            A2  => $daysInYear,
            A41 => $rateExit,
            A1  => ref $purpleUseRate eq 'ARRAY'
            ? $purpleUseRate->[0]
            : $purpleUseRate,
        }
    );

    if ( $model->{matricesData} ) {
        push @{ $model->{matricesData}[0] },
          Arithmetic(
            name => "Notional $model->{timebandName} unit rate"
              . ' for transmission exit (p/kWh)',
            rows       => $tariffs->{rows},
            arithmetic => '=100/A2*A41*A9',
            arguments  => {
                A2  => $hoursInPurple,
                A41 => $rateExit,
                A9  => (
                    ref $purpleUseRate eq 'ARRAY'
                    ? $purpleUseRate->[0]
                    : $purpleUseRate
                )->{arguments}{A9},
            },
          );
        $model->{matricesData}[2] = $activeCoincidence;
        $model->{matricesData}[3] = $hoursInPurple;
        $model->{matricesData}[4] = $daysInYear;
    }

    $model->{summaryInformationColumns}[1] = Arithmetic(
        name          => 'Transmission exit charge (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=0.01*A9*A1*A2',
        arguments     => {
            A1 => $importCapacity,
            A2 => $capacityChargeT,
            A9 => $daysInYear,
            A7 => $tariffDaysInYearNot,
        },
    );

    my $importCapacityExceededAdjustment = Arithmetic(
        name =>
          'Adjustment to exceeded import capacity charge for DSM (p/kVA/day)',
        defaultFormat => '0.00softnz',
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
        defaultFormat => '0.00softnz'
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

    push @{ $model->{matricesData}[0] },
      Stack( sources => [$unitRateFcpLricDSM] )
      if $model->{matricesData};

    my (
        $importCapacityScaledRound, $purpleRateFcpLricRound,
        $fixedDchargeTrueRound,     $importCapacityScaledSaved,
        $importCapacityExceeded,    $exportCapacityExceeded,
        $importCapacityScaled,      $purpleRateFcpLric,
    );

    my $demandScalingShortfall;

    if ( $model->{legacy201} ) {

        $capacityChargeT = Arithmetic(
            name       => 'Import capacity charge before scaling (p/kVA/day)',
            arithmetic => '=A7+IF(A6=0,1,A4/A5)*A1',
            defaultFormat => '0.00softnz',
            arguments     => {
                A1 => $fcpLricDemandCapacityChargeBig,
                A4 => $chargeableCapacity,
                A5 => $importCapacity,
                A6 => $importCapacity,
                A7 => $capacityChargeT,
            }
        );

        $model->{Thursday32} = [
            Arithmetic(
                name          => 'FCP/LRIC capacity-based charge (£/year)',
                arithmetic    => '=A1*A4*A9/100',
                defaultFormat => '0softnz',
                arguments     => {
                    A1 => $model->{demandCapacityFcpLric},
                    A4 => $chargeableCapacity,
                    A9 => $daysInYear,
                }
            ),
            Arithmetic(
                name          => 'FCP/LRIC unit-based charge (£/year)',
                arithmetic    => '=A1*A4*A9/100',
                defaultFormat => '0softnz',
                arguments     => {
                    A1 => $model->{demandConsumptionFcpLric},
                    A4 => $chargeableCapacity,
                    A9 => $daysInYear,
                }
            ),
        ];
        my $tariffHoursInPurple = Arithmetic(
            name => "Number of $model->{timebandName} hours connected in year",
            defaultFormat => '0.0softnz',
            arithmetic    => '=A2-A1',
            arguments     => {
                A2 => $hoursInPurple,
                A1 => $tariffHoursInPurpleNot,

            }
        );

        $demandScalingShortfall = Arithmetic(
            name          => 'Additional amount to be recovered (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => '=A1'
              . '-(SUM(A21_A22)+SUMPRODUCT(A31_A32,A33_A34)'
              . '+SUMPRODUCT(A41_A42,A43_A44,A35_A36,A51_A52)/A54'
              . ')*A9/100',
            arguments => {
                A1      => $totalRevenue3,
                A31_A32 => $capacityChargeT,
                A33_A34 => $importCapacity,
                A9      => $daysInYear,
                A21_A22 => $fixedDcharge,
                A41_A42 => $unitRateFcpLricDSM,
                A43_A44 => $activeCoincidence935,
                A35_A36 => $importCapacity935,
                A51_A52 => $tariffHoursInPurple,
                A54     => $daysInYear,
            }
        );

    }
    else {    # not legacy201

        push @{ $model->{calc2Tables} },
          my $capacityChargeT1 = Arithmetic(
            name          => 'Import capacity charge from charge 1 (p/kVA/day)',
            groupName     => 'Charge 1',
            arithmetic    => '=IF(A6=0,1,A4/A5)*A1',
            defaultFormat => '0.00softnz',
            arguments     => {
                A1 => $fcpLricDemandCapacityChargeBig,
                A4 => $chargeableCapacity,
                A5 => $importCapacity,
                A6 => $importCapacity,
            },
          );

        $capacityChargeT = Arithmetic(
            name       => 'Import capacity charge before scaling (p/kVA/day)',
            arithmetic => '=A7+A1',
            defaultFormat => '0.00softnz',
            arguments     => {
                A1 => $capacityChargeT1,
                A7 => $capacityChargeT,
            }
        );

        $model->{Thursday32} = [
            Arithmetic(
                name          => 'FCP/LRIC capacity-based charge (£/year)',
                arithmetic    => '=A1*A4*A9/100',
                defaultFormat => '0softnz',
                arguments     => {
                    A1 => $model->{demandCapacityFcpLric},
                    A4 => $chargeableCapacity,
                    A9 => $daysInYear,
                }
            ),
            Arithmetic(
                name          => 'FCP/LRIC unit-based charge (£/year)',
                arithmetic    => '=A1*A4*A9/100',
                defaultFormat => '0softnz',
                arguments     => {
                    A1 => $model->{demandConsumptionFcpLric},
                    A4 => $chargeableCapacity,
                    A9 => $daysInYear,
                }
            ),
        ];
        my $tariffHoursInPurple = Arithmetic(
            name => "Number of $model->{timebandName} hours connected in year",
            defaultFormat => '0.0softnz',
            arithmetic    => '=A2-A1',
            arguments     => {
                A2 => $hoursInPurple,
                A1 => $tariffHoursInPurpleNot,

            }
        );

        $demandScalingShortfall = Arithmetic(
            name          => 'Additional amount to be recovered (£/year)',
            defaultFormat => '0softnz',
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
                        name => 'Revenue from demand charge 1 (£/year)',
                        defaultFormat => '0softnz',
                        arithmetic    => '=IF(A123,0,A1)+('
                          . 'SUMPRODUCT(A64_A65,A31_A32,A33_A34)+'
                          . 'SUMPRODUCT(A66_A67,A41_A42,A43_A44,A35_A36,A51_A52)/A54'
                          . ')*A9/100',
                        arguments => {
                            A123 => $model->{transparencyMasterFlag},
                            A1 => $model->{transparency}{baselineItem}{119104},
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
                        name => 'Revenue from demand charge 1 (£/year)',
                        defaultFormat => '0softnz',
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

    }

    my $edcmIndirect = Arithmetic(
        name          => 'Indirect costs on EDCM demand (£/year)',
        defaultFormat => '0softnz',
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
        defaultFormat => '0softnz',
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
        defaultFormat => '0softnz',
        arithmetic    => '=A1*(A20+A23)',
        arguments     => {
            A1  => $rateRates,
            A20 => $totalAssetsCapacity,
            A23 => $totalAssetsConsumption,
        },
    );
    $model->{transparency}{dnoTotalItem}{1255} = $edcmRates
      if $model->{transparency};

    $model->fudge41(
        $activeCoincidence, $importCapacity,
        $edcmIndirect,      $edcmDirect,
        $edcmRates,         $daysInYear,
        \$capacityChargeT,  \$demandScalingShortfall,
        $indirectExposure,  $reactiveCoincidence,
        $powerFactorInModel,
    );

    push @{ $model->{calc4Tables} }, $demandScalingShortfall;

    ( $scalingChargeCapacity, $scalingChargeUnits ) = $model->demandScaling41(
        $importCapacity,       $demandScalingShortfall,
        $daysInYear,           $assetsFixed,
        $assetsCapacityCooked, $assetsConsumptionCooked,
        $capacityChargeT,      $fixedDcharge,
    );

    ( $edcmDirect, $edcmRates ) =
      $model->{mitigateUndueSecrecy}
      ->interimRecookTotals( $demandScalingShortfall, $edcmDirect, $edcmRates, )
      if $model->{mitigateUndueSecrecy};

    $model->{summaryInformationColumns}[2] = Arithmetic(
        name          => 'Direct cost allocation (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=A1*MAX(A2,'
          . '0-(A21+IF(A22=0,0,(1-A55/A54)*A31/(A32-A56)*IF(A52=0,1,A51/A53)*A5))'
          . ')*A3*0.01*A7/A9',
        arguments => {
            A1  => $importCapacity,
            A2  => $scalingChargeCapacity,
            A21 => $capacityChargeT,
            A22 => $activeCoincidence935,
            A5  => $demandConsumptionFcpLric,
            A51 => $chargeableCapacity,
            A52 => $importCapacity,
            A53 => $importCapacity,
            A3  => $daysInYear,
            A7  => $edcmDirect,
            A8  => $edcmRates,
            A9  => $demandScalingShortfall,
            A54 => $hoursInPurple,
            A55 => $tariffHoursInPurpleNot,
            A56 => $tariffDaysInYearNot,
            A31 => $daysInYear,
            A32 => $daysInYear,
        },
    );

    $model->{summaryInformationColumns}[4] = Arithmetic(
        name          => 'Network rates allocation (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=A1*MAX(A2,'
          . '0-(A21+IF(A22=0,0,(1-A55/A54)*A31/(A32-A56)*IF(A52=0,1,A51/A53)*A5))'
          . ')*A3*0.01*A8/A9',
        arguments => {
            A1  => $importCapacity,
            A2  => $scalingChargeCapacity,
            A21 => $capacityChargeT,
            A22 => $activeCoincidence935,
            A5  => $demandConsumptionFcpLric,
            A51 => $chargeableCapacity,
            A52 => $importCapacity,
            A53 => $importCapacity,
            A3  => $daysInYear,
            A7  => $edcmDirect,
            A8  => $edcmRates,
            A9  => $demandScalingShortfall,
            A54 => $hoursInPurple,
            A55 => $tariffHoursInPurpleNot,
            A56 => $tariffDaysInYearNot,
            A31 => $daysInYear,
            A32 => $daysInYear,
        },
    );

    $model->{summaryInformationColumns}[7] = Arithmetic(
        name          => 'Demand scaling asset based (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=A1*MAX(A2,'
          . '0-(A21+IF(A22=0,0,(1-A55/A54)*A31/(A32-A56)*IF(A52=0,1,A51/A53)*A5))'
          . ')*A3*0.01*(1-(A8+A7)/A9)',
        arguments => {
            A1  => $importCapacity,
            A2  => $scalingChargeCapacity,
            A21 => $capacityChargeT,
            A22 => $activeCoincidence935,
            A5  => $demandConsumptionFcpLric,
            A51 => $chargeableCapacity,
            A52 => $importCapacity,
            A53 => $importCapacity,
            A3  => $daysInYear,
            A7  => $edcmDirect,
            A8  => $edcmRates,
            A9  => $demandScalingShortfall,
            A54 => $hoursInPurple,
            A55 => $tariffHoursInPurpleNot,
            A56 => $tariffDaysInYearNot,
            A31 => $daysInYear,
            A32 => $daysInYear,
        },
    );

    $importCapacityScaled =
      $scalingChargeCapacity
      ? Arithmetic(
        name          => 'Total import capacity charge p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=MAX(0-(A3*A31*A33/A32),A1+A2)',
        arguments     => {
            A1  => $capacityChargeT,
            A3  => $unitRateFcpLricNonDSM,
            A31 => $activeCoincidence,
            A32 => $daysInYear,
            A33 => $hoursInPurple,
            A2  => $scalingChargeCapacity,
        }
      )
      : Stack( sources => [$capacityChargeT] );

    $purpleRateFcpLric = Arithmetic(
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
            A51 => $demandConsumptionFcpLric,
            A7  => $daysInYear,
            A71 => $tariffDaysInYearNot,
            A8  => $hoursInPurple,
            A81 => $tariffHoursInPurpleNot,
        }
    ) if $unitRateFcpLricDSM;

    push @{ $model->{calc4Tables} },
      $importCapacityScaledSaved = $importCapacityScaled;

    $importCapacityScaled = Arithmetic(
        name       => 'Import capacity charge p/kVA/day',
        groupName  => 'Demand charges after scaling',
        arithmetic => '=IF(A3,MAX(0,A1),0)',
        arguments  => {
            A1 => $importCapacityScaled,
            A3 => $importEligible,
        },
        defaultFormat => '0.00softnz'
    );

    $importCapacityExceeded = Arithmetic(
        name          => 'Exceeded import capacity charge (p/kVA/day)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=A7+A2',
        defaultFormat => '0.00softnz',
        arguments     => {
            A3 => $fcpLricDemandCapacityChargeBig,
            A2 => $importCapacityExceededAdjustment,
            A4 => $chargeableCapacity,
            A5 => $importCapacity,
            A1 => $importCapacity,
            A7 => $importCapacityScaled,
        },
        defaultFormat => '0.00softnz'
    );

    $model->{summaryInformationColumns}[5] = Arithmetic(
        name          => 'FCP/LRIC charge (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=0.01*(A11*A9*A2+A1*A4*A8*(A6-A61)*(A91/(A92-A71)))',
        arguments     => {
            A1  => $importCapacity,
            A2  => $fcpLricDemandCapacityChargeBig,
            A3  => $capacityChargeT->{arguments}{A1},
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

    $fixedDchargeTrueRound = Arithmetic(
        name          => 'Import fixed charge (p/day)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=ROUND(A1,2)',
        arguments     => { A1 => $fixedDchargeTrue, },
    );

    $purpleRateFcpLricRound = Arithmetic(
        name          => "Import $model->{timebandName} unit rate (p/kWh)",
        defaultFormat => '0.000softnz',
        arithmetic    => '=ROUND(A1,3)',
        arguments     => { A1 => $purpleRateFcpLric, },
    );

    $importCapacityScaledRound = Arithmetic(
        name          => 'Import capacity rate (p/kVA/day)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=ROUND(A1,2)',
        arguments     => { A1 => $importCapacityScaled, },
    );

    $exportCapacityExceeded = Arithmetic(
        name          => 'Export exceeded capacity rate (p/kVA/day)',
        defaultFormat => '0.00copynz',
        arithmetic    => '=A1',
        arguments     => { A1 => $exportCapacityChargeRound, },
    );

    my $importCapacityExceededRound = Arithmetic(
        name          => 'Import exceeded capacity rate (p/kVA/day)',
        defaultFormat => '0.00softnz',
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
