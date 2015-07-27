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
        arithmetic    => '=A1-A7-A91-A92',
        arguments     => {
            A1  => $$shortfallRef,
            A7  => $indirect,
            A91 => $direct,
            A92 => $rates,
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
        arithmetic => '=SQRT(A1*A2+A3*A4)',
        arguments  => {
            A1 => $activeCoincidence,
            A2 => $activeCoincidence,
            A3 => $reactiveCoincidence,
            A4 => $reactiveCoincidence,
        }
    ) if $model->{dcp183};

    my $slope =
      $model->{dcp185}
      ? Arithmetic(
        name => 'Marginal revenue effect of demand'
          . ( $model->{dcp185} == 2 ? ' and indirect cost adders' : ' adder' ),
        defaultFormat => '0soft',
        arithmetic    => '=A1*(A4+A5)*IF(A6<0,1,A7)',
        arguments     => {
            A1 => $agreedCapacity,
            A4 => $ynonFudge,
            A5 => $activeCoincidence,
            A6 => $adderAmount,
            A7 => $indirectExposure,
        }
      )
      : Arithmetic(
        name          => 'Marginal revenue effect of demand adder',
        defaultFormat => '0soft',
        arithmetic    => '=A1*(A4+A5)',
        arguments     => {
            A1 => $agreedCapacity,
            A4 => $ynonFudge,
            A5 => $activeCoincidence,
        }
      );

    my $fudgeIndirect =
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
            A123   => $model->{transparencyMasterFlag},
            A1     => $model->{transparency}{ol119102},
            A2_A3 => $model->{transparency},
            A4_A5 => $fudgeIndirect,
            A6_A7 => $agreedCapacity,
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
        arithmetic => '=IF(A2,A3/SUMPRODUCT(A4_A5,A64_A65),0)',
        arguments  => {
            A2       => $indirect,
            A3       => $indirect,
            A4_A5   => $fudgeIndirect,
            A64_A65 => $agreedCapacity,
        },
        location => 'Charging rates',
      )
      : Arithmetic(
        name       => 'Indirect costs application rate',
        groupName  => 'EDCM demand adders',
        arithmetic => '=IF(A2,A3/A4,0)',
        arguments  => {
            A2 => $indirect,
            A3 => $indirect,
            A4 => $totalIndirectFudge,
        },
        location => 'Charging rates',
      );
    $model->{transparency}{olFYI}{1262} = $indirectAppRate
      if $model->{transparency};

    $$capacityChargeRef = Arithmetic(
        arithmetic => '=A1+A3*A4*100/A9',
        name => 'Capacity charge after applying indirect cost charge p/kVA/day',
        arguments => {
            A1 => $$capacityChargeRef,
            A3 => $indirectAppRate,
            A4 => $fudgeIndirect,
            A9 => $daysInYear,
        }
    );

    if ( $model->{matricesData} ) {
        push @{ $model->{matricesData}[0] },
          Arithmetic(
            arithmetic => '=A1*A3*100/A9',
            rows       => $$capacityChargeRef->{rows},
            name       => 'Notional indirect cost contribution p/kWh',
            arguments  => {
                A3 => $indirectAppRate,
                A1 => $indirectExposure,
                A9 => $model->{matricesData}[3],
            }
          );
        push @{ $model->{matricesData}[1] },
          Arithmetic(
            arithmetic => '=A1*A3*A4*100/A9',
            rows       => $$capacityChargeRef->{rows},
            name       => 'Notional indirect cost contribution p/kVA/day',
            arguments  => {
                A3 => $indirectAppRate,
                A4 => $ynonFudge,
                A1 => $indirectExposure,
                A9 => $daysInYear,
            }
          );
    }

    $model->{summaryInformationColumns}[3] = Arithmetic(
        name          => 'Indirect cost allocation (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=A1*A3*A7',
        arguments     => {
            A1 => $agreedCapacity,
            A3 => $indirectAppRate,
            A7 => $fudgeIndirect,
        },
    );

    my $totalSlope =
      $model->{transparencyMasterFlag}
      ? Arithmetic(
        name          => 'Total marginal revenue effect of demand adder',
        defaultFormat => '0soft',
        arithmetic    => '=IF(A123,0,A1)+SUMPRODUCT(A2_A3,A4_A5)',
        arguments     => {
            A123   => $model->{transparencyMasterFlag},
            A1     => $model->{transparency}{ol119103},
            A2_A3 => $model->{transparency},
            A4_A5 => $slope,
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
        arithmetic => '=IF(A9,A1*A2/SUM(A4_A5),0)',
        arguments  => {
            A1     => $ynonFudge41,
            A2     => $adderAmount,
            A9     => $adderAmount,
            A4_A5 => $slope,
        },
        location => 'Charging rates',
      )
      : Arithmetic(
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
    $model->{transparency}{olFYI}{1261} = $fixedAdderRate
      if $model->{transparency};

    $$capacityChargeRef = Arithmetic(
        arithmetic => '=A1+A3*(A7+A4)*100/A9'
          . ( $model->{dcp185} ? '*IF(A6<0,1,A8)' : '' ),
        name =>
          'Capacity charge after applying fixed adder ex indirects p/kVA/day',
        arguments => {
            A1 => $$capacityChargeRef,
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
        }
    );

    if ( $model->{matricesData} ) {
        push @{ $model->{matricesData}[0] },
          Arithmetic(
            arithmetic => '=A3*100/A9' . ( $model->{dcp185} ? '*A8' : '' ),
            rows      => $$capacityChargeRef->{rows},
            name      => 'Notional fixed adder p/kWh',
            arguments => {
                A3 => $fixedAdderRate,
                A9 => $model->{matricesData}[3],
                $model->{dcp185} ? ( A8 => $indirectExposure ) : (),
            }
          );
        push @{ $model->{matricesData}[1] },
          Arithmetic(
            arithmetic => '=A3*A4*100/A9'
              . ( $model->{dcp185} ? '*A8' : '' ),
            rows      => $$capacityChargeRef->{rows},
            name      => 'Notional fixed adder p/kVA/day',
            arguments => {
                A3 => $fixedAdderRate,
                A4 => $ynonFudge,
                A9 => $daysInYear,
                $model->{dcp185} ? ( A8 => $indirectExposure ) : (),
            }
          );
    }

    $model->{summaryInformationColumns}[6] = Arithmetic(
        name          => 'Demand scaling fixed adder (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=A1*A3*(A71+A72)'
          . ( $model->{dcp185} ? '*IF(A6<0,1,A8)' : '' ),
        arguments => {
            A1  => $agreedCapacity,
            A3  => $fixedAdderRate,
            A71 => $ynonFudge,
            A72 => $activeCoincidence,
            $model->{dcp185}
            ? (
                A6 => $adderAmount,
                A8 => $indirectExposure,
              )
            : (),
        },
    );

    $model->{indirectShareOfCapacityBasedAllocation} = Arithmetic(
        name          => 'Indirect cost share of capacity-based allocation',
        defaultFormat => '%soft',
        arithmetic    => '=A4/(A1*(A2-A7-A91-A92)+A3)',
        arguments     => {
            A4  => $indirect,
            A1  => $ynonFudge41,
            A2  => $$shortfallRef,
            A3  => $indirect,
            A7  => $indirect,
            A91 => $direct,
            A92 => $rates,
        }
    );

    $$shortfallRef = Arithmetic(
        name          => 'Residual residual (£/year)',
        groupName     => 'Residual EDCM demand revenue',
        defaultFormat => '0softnz',
        arithmetic    => '=(1-A1)*(A2-A3-A71-A72)+A81+A82',
        arguments     => {
            A1  => $ynonFudge41,
            A2  => $$shortfallRef,
            A3  => $indirect,
            A71 => $direct,
            A81 => $direct,
            A72 => $rates,
            A82 => $rates,
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
        arithmetic    => '=A1/A2',
        arguments     => {
            A1 => $direct,
            A2 => $$shortfallRef,
        }
    );

    $model->{ratesShareOfAssetBasedAllocation} = Arithmetic(
        name          => 'Network rates share of asset-based allocation',
        defaultFormat => '%soft',
        arithmetic    => '=A1/A2',
        arguments     => {
            A1 => $rates,
            A2 => $$shortfallRef,
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
        arithmetic    => '=(A2+A1)*A4',
        arguments     => {
            A2 => $assetsCapacity,
            A1 => $assetsConsumption,
            A4 => $agreedCapacity,
        }
    );

    my $totalSlopeCapacity =
      $model->{transparencyMasterFlag}
      ? Arithmetic(
        name => 'Total non sole use notional assets subject to matching (£)',
        defaultFormat => '0softnz',
        arithmetic    => '=IF(A123,0,A1)+SUMPRODUCT(A21_A22,A51_A52)',
        arguments     => {
            A123     => $model->{transparencyMasterFlag},
            A1       => $model->{transparency}{ol119305},
            A21_A22 => $model->{transparency},
            A51_A52 => $slopeCapacity,
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
        name       => 'Threshold for asset percentage adder on capacity',
        arithmetic => '=IF(A3,-0.01*A1*A4*A5/A2,0)',
        arguments  => {
            A1 => $capacityCharge,
            A2 => $slopeCapacity,
            A3 => $slopeCapacity,
            A4 => $daysInYear,
            A5 => $agreedCapacity,
        }
    );

    my $demandScaling =
      $model->{legacy201}
      ? Arithmetic(
        name          => 'Annual charge on assets',
        groupName     => 'Demand scaling rate',
        defaultFormat => '%soft',
        arithmetic    => '=IF(A4,A1/SUM(A2_A3),0)',
        arguments =>
          { A1 => $shortfall, A4 => $shortfall, A2_A3 => $slopeCapacity, },
        location => 'Charging rates',
      )
      : Arithmetic(
        name          => 'Annual charge on assets',
        groupName     => 'Demand scaling rate',
        defaultFormat => '%soft',
        arithmetic    => '=IF(A4,A1/A2,0)',
        arguments =>
          { A1 => $shortfall, A4 => $shortfall, A2 => $totalSlopeCapacity, },
        location => 'Charging rates',
      );
    $model->{transparency}{olFYI}{1258} = $demandScaling
      if $model->{transparency};

    my $scalingChargeCapacity = Arithmetic(
        arithmetic => '=A3*(A62+A63)*100/A9',
        name       => 'Demand scaling p/kVA/day',
        arguments  => {
            A3  => $demandScaling,
            A62 => $assetsCapacity,
            A63 => $assetsConsumption,
            A9  => $daysInYear,
            A1  => $capacityCharge,
        }
    );

    if ( $model->{matricesData} ) {
        push @{ $model->{matricesData}[0] },
          Arithmetic(
            arithmetic => '=IF(A5,A1*A3*100/A9/A6,0)',
            name       => 'Asset-based contribution p/kWh',
            arguments  => {
                A3 => $demandScaling,
                A1 => $assetsConsumption,
                A9 => $model->{matricesData}[3],
                A5 => $model->{matricesData}[2],
                A6 => $model->{matricesData}[2],
            }
          );
        push @{ $model->{matricesData}[1] },
          Arithmetic(
            arithmetic => '=A1*A3*100/A9',
            name       => 'Asset-based contribution p/kVA/day',
            arguments  => {
                A3 => $demandScaling,
                A1 => $assetsCapacity,
                A9 => $daysInYear,
            }
          );
    }

    $scalingChargeCapacity;

}

1;
