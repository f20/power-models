package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2012 Energy Networks Association Limited and others.
Copyright 2013-2015 Franck Latrémolière, Reckon LLP and others.

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

sub fudge41 {

    my (
        $model,            $activeCoincidence,   $agreedCapacity,
        $indirect,         $direct,              $rates,
        $daysInYear,       $capacityChargeRef,   $shortfallRef,
        $indirectExposure, $reactiveCoincidence, $powerFactorInModel,
    ) = @_;

    my $adderAmount = Arithmetic(
        name          => 'Amount to be recovered from adders ex costs (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=IV1-IV7-IV91-IV92',
        arguments     => {
            IV1  => $$shortfallRef,
            IV7  => $indirect,
            IV91 => $direct,
            IV92 => $rates,
        },
    );
    $model->{transparency}{olFYI}{1259} = $adderAmount
      if $model->{transparency};

    my $ynonFudge = Constant(
        name => 'Factor for the allocation of capacity scaling',
        data => [0.5],
    );

    my $ynonFudge41 = Constant(
        name => 'Proportion of residual to go into fixed adder',
        data => [0.2],
    );

    $activeCoincidence = Arithmetic(
        name       => 'Peak-time capacity use per kVA of agreed capacity',
        arithmetic => '=SQRT(IV1*IV2+IV3*IV4)',
        arguments  => {
            IV1 => $activeCoincidence,
            IV2 => $activeCoincidence,
            IV3 => $reactiveCoincidence,
            IV4 => $reactiveCoincidence,
        }
    ) if $model->{dcp183};

    my $slope =
      $model->{dcp185}
      ? Arithmetic(
        name => 'Marginal revenue effect of demand'
          . ( $model->{dcp185} == 2 ? ' and indirect cost adders' : ' adder' ),
        defaultFormat => '0soft',
        arithmetic    => '=IV1*(IV4+IV5)*IF(IV6<0,1,IV7)',
        arguments     => {
            IV1 => $agreedCapacity,
            IV4 => $ynonFudge,
            IV5 => $activeCoincidence,
            IV6 => $adderAmount,
            IV7 => $indirectExposure,
        }
      )
      : Arithmetic(
        name          => 'Marginal revenue effect of demand adder',
        defaultFormat => '0soft',
        arithmetic    => '=IV1*(IV4+IV5)',
        arguments     => {
            IV1 => $agreedCapacity,
            IV4 => $ynonFudge,
            IV5 => $activeCoincidence,
        }
      );

    my $fudgeIndirect =
      $model->{dcp185} && $model->{dcp185} == 2
      ? Arithmetic(
        name       => 'Data for capacity-based allocation of indirect costs',
        groupName  => 'Allocation of indirect costs',
        arithmetic => '=IF(IV6<0,1,IV1)*(IV2+IV3)',
        arguments  => {
            IV2 => $ynonFudge,
            IV3 => $activeCoincidence,
            IV1 => $indirectExposure,
            IV6 => $adderAmount,
        }
      )
      : Arithmetic(
        name       => 'Data for capacity-based allocation of indirect costs',
        groupName  => 'Allocation of indirect costs',
        arithmetic => '=IV1*(IV2+IV3)',
        arguments  => {
            IV2 => $ynonFudge,
            IV3 => $activeCoincidence,
            IV1 => $indirectExposure,
        }
      );

    my $totalIndirectFudge =
      $model->{transparencyMasterFlag}
      ? Arithmetic(
        name          => 'Total marginal effect of indirect cost adder',
        defaultFormat => '0soft',
        arithmetic    => '=IF(IV123,0,IV1)+SUMPRODUCT(IV2_IV3,IV4_IV5,IV6_IV7)',
        arguments     => {
            IV123   => $model->{transparencyMasterFlag},
            IV1     => $model->{transparency}{ol119102},
            IV2_IV3 => $model->{transparency},
            IV4_IV5 => $fudgeIndirect,
            IV6_IV7 => $agreedCapacity,
        },
      )
      : SumProduct(
        name          => 'Total marginal effect of indirect cost adder',
        defaultFormat => '0soft',
        matrix        => $fudgeIndirect,
        vector        => $agreedCapacity
      );

    $model->{transparency}{olTabCol}{119102} = $totalIndirectFudge
      if $model->{transparency};

    my $indirectAppRate =
      $model->{legacy201}
      ? Arithmetic(
        name       => 'Indirect costs application rate',
        groupName  => 'EDCM demand adders',
        arithmetic => '=IF(IV2,IV3/SUMPRODUCT(IV4_IV5,IV64_IV65),0)',
        arguments  => {
            IV2       => $indirect,
            IV3       => $indirect,
            IV4_IV5   => $fudgeIndirect,
            IV64_IV65 => $agreedCapacity,
        },
        location => 'Charging rates',
      )
      : Arithmetic(
        name       => 'Indirect costs application rate',
        groupName  => 'EDCM demand adders',
        arithmetic => '=IF(IV2,IV3/IV4,0)',
        arguments  => {
            IV2 => $indirect,
            IV3 => $indirect,
            IV4 => $totalIndirectFudge,
        },
        location => 'Charging rates',
      );
    $model->{transparency}{olFYI}{1262} = $indirectAppRate
      if $model->{transparency};

    $$capacityChargeRef = Arithmetic(
        arithmetic => '=IV1+IV3*IV4*100/IV9',
        name => 'Capacity charge after applying indirect cost charge p/kVA/day',
        arguments => {
            IV1 => $$capacityChargeRef,
            IV3 => $indirectAppRate,
            IV4 => $fudgeIndirect,
            IV9 => $daysInYear,
        }
    );

    if ( $model->{matricesData} ) {
        push @{ $model->{matricesData}[0] },
          Arithmetic(
            arithmetic => '=IV1*IV3*100/IV9',
            rows       => $$capacityChargeRef->{rows},
            name       => 'Notional indirect cost contribution p/kWh',
            arguments  => {
                IV3 => $indirectAppRate,
                IV1 => $indirectExposure,
                IV9 => $model->{matricesData}[3],
            }
          );
        push @{ $model->{matricesData}[1] },
          Arithmetic(
            arithmetic => '=IV1*IV3*IV4*100/IV9',
            rows       => $$capacityChargeRef->{rows},
            name       => 'Notional indirect cost contribution p/kVA/day',
            arguments  => {
                IV3 => $indirectAppRate,
                IV4 => $ynonFudge,
                IV1 => $indirectExposure,
                IV9 => $daysInYear,
            }
          );
    }

    $model->{summaryInformationColumns}[3] = Arithmetic(
        name          => 'Indirect cost allocation (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=IV1*IV3*IV7',
        arguments     => {
            IV1 => $agreedCapacity,
            IV3 => $indirectAppRate,
            IV7 => $fudgeIndirect,
        },
    );

    my $totalSlope =
      $model->{transparencyMasterFlag}
      ? Arithmetic(
        name          => 'Total marginal revenue effect of demand adder',
        defaultFormat => '0soft',
        arithmetic    => '=IF(IV123,0,IV1)+SUMPRODUCT(IV2_IV3,IV4_IV5)',
        arguments     => {
            IV123   => $model->{transparencyMasterFlag},
            IV1     => $model->{transparency}{ol119103},
            IV2_IV3 => $model->{transparency},
            IV4_IV5 => $slope,
        },
      )
      : GroupBy(
        name          => 'Total marginal revenue effect of demand adder',
        defaultFormat => '0soft',
        source        => $slope
      );

    $model->{transparency}{olTabCol}{119103} = $totalSlope
      if $model->{transparency};

    my $fixedAdderRate =
      $model->{legacy201}
      ? Arithmetic(
        name       => 'Fixed adder ex indirects application rate',
        groupName  => 'EDCM demand adders',
        arithmetic => '=IF(IV9,IV1*IV2/SUM(IV4_IV5),0)',
        arguments  => {
            IV1     => $ynonFudge41,
            IV2     => $adderAmount,
            IV9     => $adderAmount,
            IV4_IV5 => $slope,
        },
        location => 'Charging rates',
      )
      : Arithmetic(
        name       => 'Fixed adder ex indirects application rate',
        groupName  => 'EDCM demand adders',
        arithmetic => '=IF(IV9,IV1*IV2/IV4,0)',
        arguments  => {
            IV1 => $ynonFudge41,
            IV2 => $adderAmount,
            IV9 => $adderAmount,
            IV4 => $totalSlope,
        },
        location => 'Charging rates',
      );
    $model->{transparency}{olFYI}{1261} = $fixedAdderRate
      if $model->{transparency};

    $$capacityChargeRef = Arithmetic(
        arithmetic => '=IV1+IV3*(IV7+IV4)*100/IV9'
          . ( $model->{dcp185} ? '*IF(IV6<0,1,IV8)' : '' ),
        name =>
          'Capacity charge after applying fixed adder ex indirects p/kVA/day',
        arguments => {
            IV1 => $$capacityChargeRef,
            IV3 => $fixedAdderRate,
            IV4 => $activeCoincidence,
            IV7 => $ynonFudge,
            IV9 => $daysInYear,
            $model->{dcp185}
            ? (
                IV6 => $adderAmount,
                IV8 => $indirectExposure,
              )
            : (),
        }
    );

    if ( $model->{matricesData} ) {
        push @{ $model->{matricesData}[0] },
          Arithmetic(
            arithmetic => '=IV3*100/IV9' . ( $model->{dcp185} ? '*IV8' : '' ),
            rows      => $$capacityChargeRef->{rows},
            name      => 'Notional fixed adder p/kWh',
            arguments => {
                IV3 => $fixedAdderRate,
                IV9 => $model->{matricesData}[3],
                $model->{dcp185} ? ( IV8 => $indirectExposure ) : (),
            }
          );
        push @{ $model->{matricesData}[1] },
          Arithmetic(
            arithmetic => '=IV3*IV4*100/IV9'
              . ( $model->{dcp185} ? '*IV8' : '' ),
            rows      => $$capacityChargeRef->{rows},
            name      => 'Notional fixed adder p/kVA/day',
            arguments => {
                IV3 => $fixedAdderRate,
                IV4 => $ynonFudge,
                IV9 => $daysInYear,
                $model->{dcp185} ? ( IV8 => $indirectExposure ) : (),
            }
          );
    }

    $model->{summaryInformationColumns}[6] = Arithmetic(
        name          => 'Demand scaling fixed adder (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=IV1*IV3*(IV71+IV72)'
          . ( $model->{dcp185} ? '*IF(IV6<0,1,IV8)' : '' ),
        arguments => {
            IV1  => $agreedCapacity,
            IV3  => $fixedAdderRate,
            IV71 => $ynonFudge,
            IV72 => $activeCoincidence,
            $model->{dcp185}
            ? (
                IV6 => $adderAmount,
                IV8 => $indirectExposure,
              )
            : (),
        },
    );

    $model->{indirectShareOfCapacityBasedAllocation} = Arithmetic(
        name          => 'Indirect cost share of capacity-based allocation',
        defaultFormat => '%soft',
        arithmetic    => '=IV4/(IV1*(IV2-IV7-IV91-IV92)+IV3)',
        arguments     => {
            IV4  => $indirect,
            IV1  => $ynonFudge41,
            IV2  => $$shortfallRef,
            IV3  => $indirect,
            IV7  => $indirect,
            IV91 => $direct,
            IV92 => $rates,
        }
    );

    $$shortfallRef = Arithmetic(
        name          => 'Residual residual (£/year)',
        groupName     => 'Residual EDCM demand revenue',
        defaultFormat => '0softnz',
        arithmetic    => '=(1-IV1)*(IV2-IV3-IV71-IV72)+IV81+IV82',
        arguments     => {
            IV1  => $ynonFudge41,
            IV2  => $$shortfallRef,
            IV3  => $indirect,
            IV71 => $direct,
            IV81 => $direct,
            IV72 => $rates,
            IV82 => $rates,
        },
    );
    $model->{transparency}{olFYI}{1257} = $$shortfallRef
      if $model->{transparency};

    0 and Columnset(
        name    => 'Calculation of revenue shortfall',
        columns => [$$shortfallRef]
    );

    $model->{directShareOfAssetBasedAllocation} = Arithmetic(
        name          => 'Direct cost share of asset-based allocation',
        defaultFormat => '%soft',
        arithmetic    => '=IV1/IV2',
        arguments     => {
            IV1 => $direct,
            IV2 => $$shortfallRef,
        }
    );

    $model->{ratesShareOfAssetBasedAllocation} = Arithmetic(
        name          => 'Network rates share of asset-based allocation',
        defaultFormat => '%soft',
        arithmetic    => '=IV1/IV2',
        arguments     => {
            IV1 => $rates,
            IV2 => $$shortfallRef,
        }
    );

}

sub demandScaling41 {

    my (
        $model,             $agreedCapacity, $shortfall,
        $daysInYear,        $assetsFixed,    $assetsCapacity,
        $assetsConsumption, $capacityCharge, $fixedCharge,
    ) = @_;

    my $slopeCapacity = Arithmetic(
        name          => 'Non sole use notional assets subject to matching (£)',
        groupName     => 'Demand scaling',
        defaultFormat => '0softnz',
        arithmetic    => '=(IV2+IV1)*IV4',
        arguments     => {
            IV2 => $assetsCapacity,
            IV1 => $assetsConsumption,
            IV4 => $agreedCapacity,
        }
    );

    my $totalSlopeCapacity =
      $model->{transparencyMasterFlag}
      ? Arithmetic(
        name => 'Total non sole use notional assets subject to matching (£)',
        defaultFormat => '0softnz',
        arithmetic    => '=IF(IV123,0,IV1)+SUMPRODUCT(IV21_IV22,IV51_IV52)',
        arguments     => {
            IV123     => $model->{transparencyMasterFlag},
            IV1       => $model->{transparency}{ol119305},
            IV21_IV22 => $model->{transparency},
            IV51_IV52 => $slopeCapacity,
        }
      )
      : GroupBy(
        name => 'Total non sole use notional assets subject to matching (£)',
        defaultFormat => '0softnz',
        source        => $slopeCapacity,
      );

    $model->{transparency}{olTabCol}{119305} = $totalSlopeCapacity
      if $model->{transparency};

    my $minCapacity = Arithmetic(
        name       => 'Threshold for asset percentage adder - capacity',
        arithmetic => '=IF(IV3,-0.01*IV1*IV4*IV5/IV2,0)',
        arguments  => {
            IV1 => $capacityCharge,
            IV2 => $slopeCapacity,
            IV3 => $slopeCapacity,
            IV4 => $daysInYear,
            IV5 => $agreedCapacity,
        }
    );

    my $demandScaling =
      $model->{legacy201}
      ? Arithmetic(
        name          => 'Annual charge on assets',
        groupName     => 'Demand scaling rate',
        defaultFormat => '%soft',
        arithmetic    => '=IF(IV4,IV1/SUM(IV2_IV3),0)',
        arguments =>
          { IV1 => $shortfall, IV4 => $shortfall, IV2_IV3 => $slopeCapacity, },
        location => 'Charging rates',
      )
      : Arithmetic(
        name          => 'Annual charge on assets',
        groupName     => 'Demand scaling rate',
        defaultFormat => '%soft',
        arithmetic    => '=IF(IV4,IV1/IV2,0)',
        arguments =>
          { IV1 => $shortfall, IV4 => $shortfall, IV2 => $totalSlopeCapacity, },
        location => 'Charging rates',
      );
    $model->{transparency}{olFYI}{1258} = $demandScaling
      if $model->{transparency};

    my $scalingChargeCapacity = Arithmetic(
        arithmetic => '=IV3*(IV62+IV63)*100/IV9',
        name       => 'Demand scaling p/kVA/day',
        arguments  => {
            IV3  => $demandScaling,
            IV62 => $assetsCapacity,
            IV63 => $assetsConsumption,
            IV9  => $daysInYear,
            IV1  => $capacityCharge,
        }
    );

    if ( $model->{matricesData} ) {
        push @{ $model->{matricesData}[0] },
          Arithmetic(
            arithmetic => '=IF(IV5,IV1*IV3*100/IV9/IV6,0)',
            name       => 'Asset-based contribution p/kWh',
            arguments  => {
                IV3 => $demandScaling,
                IV1 => $assetsConsumption,
                IV9 => $model->{matricesData}[3],
                IV5 => $model->{matricesData}[2],
                IV6 => $model->{matricesData}[2],
            }
          );
        push @{ $model->{matricesData}[1] },
          Arithmetic(
            arithmetic => '=IV1*IV3*100/IV9',
            name       => 'Asset-based contribution p/kVA/day',
            arguments  => {
                IV3 => $demandScaling,
                IV1 => $assetsCapacity,
                IV9 => $daysInYear,
            }
          );
    }

    $scalingChargeCapacity;

}

1;
