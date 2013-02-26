package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2012 Energy Networks Association Limited and others.

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

1144

=cut

use warnings;
use strict;
use utf8;

use SpreadsheetModel::Shortcuts ':all';
use SpreadsheetModel::SegmentRoot;

sub fudge41 {

    my (
        $model,                      $activeCoincidence,
        $agreedCapacity,             $indirect,
        $direct,                     $rates,
        $daysInYear,                 $capacityChargeRef,
        $shortfallRef,               $indirectExposure,
        $assetsCapacityDoubleCooked, $assetsConsumptionDoubleCooked,
        $reactiveCoincidence,        $powerFactorInModel,
    ) = @_;

    my $ynonFudge = Constant(
        name => 'Factor for the allocation of capacity scaling',
        data => [0.5],
    );

    my $ynonFudge41 = Constant(
        name => 'Proportion of residual to go into fixed adder',
        data => [0.2],
    );

    my $slope = Arithmetic(
        name       => 'Marginal revenue effect of demand adder',
        arithmetic => '=IV1*(IV4+IV5)',
        arguments  => {
            IV1 => $agreedCapacity,
            IV4 => $ynonFudge,
            IV5 => $activeCoincidence,
        }
    );

    my $fudgeIndirect = Arithmetic(
        name       => 'Data for capacity-based allocation of indirect costs',
        arithmetic => '=IV1*(IV2+IV3)',
        arguments  => {
            IV2 => $ynonFudge,
            IV3 => $activeCoincidence,
            IV1 => $indirectExposure,
        }
    );

    my $indirectAppRate = Arithmetic(
        name       => 'Indirect costs application rate',
        arithmetic => '=IF(IV2,IV3/SUMPRODUCT(IV4_IV5,IV64_IV65),0)',
        arguments  => {
            IV2       => $indirect,
            IV3       => $indirect,
            IV4_IV5   => $fudgeIndirect,
            IV64_IV65 => $agreedCapacity,
        },
    );

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

    my $fixedAdderAmount = Arithmetic(
        name => 'Amount to be recovered from fixed adder ex indirects (£/year)',
        arithmetic => '=IV1-IV7-IV91-IV92',
        arguments  => {
            IV1  => $$shortfallRef,
            IV7  => $indirect,
            IV91 => $direct,
            IV92 => $rates,
        },
    );

    my $fixedAdderRate = Arithmetic(
        name       => 'Fixed adder ex indirects application rate',
        arithmetic => '=IF(IV9,IV1*IV2/SUM(IV4_IV5),0)',
        arguments  => {
            IV1     => $ynonFudge41,
            IV2     => $fixedAdderAmount,
            IV9     => $fixedAdderAmount,
            IV4_IV5 => $slope,
        },
    );

    $$capacityChargeRef = Arithmetic(
        arithmetic => '=IV1+IV3*(IV7+IV4)*100/IV9',
        name =>
          'Capacity charge after applying fixed adder ex indirects p/kVA/day',
        arguments => {
            IV1 => $$capacityChargeRef,
            IV3 => $fixedAdderRate,
            IV4 => $activeCoincidence,
            IV7 => $ynonFudge,
            IV9 => $daysInYear,
        }
    );

    $model->{summaryInformationColumns}[6] = Arithmetic(
        name          => 'Demand scaling fixed adder (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=IV1*IV3*(IV71+IV72)',
        arguments     => {
            IV1  => $agreedCapacity,
            IV3  => $fixedAdderRate,
            IV71 => $ynonFudge,
            IV72 => $activeCoincidence,
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
        defaultFormat => '0softnz',
        arithmetic    => '=(IV2+IV1)*IV4',
        arguments     => {
            IV2 => $assetsCapacity,
            IV1 => $assetsConsumption,
            IV4 => $agreedCapacity,
        }
    );

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

    my $demandScaling = Arithmetic(
        name          => 'Annual charge on assets',
        defaultFormat => '%soft',
        arithmetic    => '=IF(IV4,IV1/SUM(IV2_IV3),0)',
        arguments =>
          { IV1 => $shortfall, IV4 => $shortfall, IV2_IV3 => $slopeCapacity, }
    );

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

    $scalingChargeCapacity;

}

1;
