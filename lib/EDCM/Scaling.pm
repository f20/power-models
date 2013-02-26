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

1108
1009
1144

=cut

use warnings;
use strict;
use utf8;

use SpreadsheetModel::Shortcuts ':all';
use SpreadsheetModel::SegmentRoot;

sub cookedPot55565960 {

    die 'cookedPot55565960 has been removed';

}

sub fudge41 {

    my (
        $model,                      $included,
        $tariffDorG,                 $activeCoincidence,
        $agreedCapacity,             $indirect,
        $direct,                     $rates,
        $daysInYear,                 $capacityChargeRef,
        $shortfallRef,               $importForLdno,
        $assetsCapacityDoubleCooked, $assetsConsumptionDoubleCooked,
        $reactiveCoincidence,        $powerFactorInModel,
      )
      = @_;

    my $ynonFudge = Constant(
        name     => 'Factor for the allocation of capacity scaling',
        data     => [0.5],
        number   => 1108,
        dataset  => $model->{dataset},
        appendTo => $model->{inputTables}
    );

    my $ynonFudge41 = Constant(
        name     => 'Proportion of residual to go into fixed adder',
        data     => [0.2],
        number   => 1109,
        dataset  => $model->{dataset},
        appendTo => $model->{inputTables}
    );

    my $slope = 0 ? undef: Arithmetic(
        name       => 'Marginal revenue effect of demand adder',
        arithmetic => '=IF(AND(IV1,IV2="Demand"),IV3*(IV4+IV5),0)',
        arguments  => {
            IV1 => $included,
            IV2 => $tariffDorG,
            IV3 => $agreedCapacity,
            IV4 => $ynonFudge,
            IV5 => $activeCoincidence,
        }
    );

    my $fudgeIndirect =
      $model->{someStupidIssue3aThing}
      ? Arithmetic(
        name       => 'Data for capacity-based allocation of indirect costs',
        arithmetic => '=IF(AND(IV1,IV5="Demand"),IF(IV4,0.5,1)*(IV2+'
          . 'IF(IV35,SQRT(IV31^2+IV32^2),IV38)),0)',
        arguments => {
            IV1  => $included,
            IV5  => $tariffDorG,
            IV2  => $ynonFudge,
            IV31 => $activeCoincidence,
            IV38 => $activeCoincidence,
            IV32 => $reactiveCoincidence,
            IV35 => Constant(
                number        => 4999,
                name          => 'Issue 3a flag',
                defaultFormat => 'boolhard',
                data          => [ [ $model->{issue3a} ? 'TRUE' : 'FALSE' ] ]
            ),
            IV4 => $importForLdno,
        }
      )
      : Arithmetic(
        name       => 'Data for capacity-based allocation of indirect costs',
        arithmetic => '=IF(AND(IV1,IV5="Demand"),IF(IV4,0.5,1)*(IV2+IV3),0)',
        arguments  => {
            IV1 => $included,
            IV5 => $tariffDorG,
            IV2 => $ynonFudge,
            IV3 => $activeCoincidence,
            IV4 => $importForLdno,
        }
      );

    my $indirectAppRate = Arithmetic(
        name       => 'Indirect costs application rate',
        arithmetic => '=IV3/SUMPRODUCT(IV4_IV5,IV64_IV65)',
        arguments  => {
            IV3       => $indirect,
            IV4_IV5   => $fudgeIndirect,
            IV64_IV65 => $agreedCapacity,
        },
    );

    $$capacityChargeRef = Arithmetic(
        arithmetic => '=IF(AND(IV1,IV5="Demand"),IV2+IV3*IV4*100/IV9,0)',
        name => 'Capacity charge after applying indirect cost charge p/kVA/day',
        arguments => {
            IV1 => $included,
            IV5 => $tariffDorG,
            IV2 => $$capacityChargeRef,
            IV3 => $indirectAppRate,
            IV4 => $fudgeIndirect,
            IV9 => $daysInYear,
        }
    );

    $model->{summaryInformationColumns}[3] = Arithmetic(
        name          => 'Indirect cost allocation (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=IF(AND(IV4,IV5="Demand"),IV1*IV3*IV7,0)',
        arguments     => {
            IV1 => $agreedCapacity,
            IV3 => $indirectAppRate,
            IV7 => $fudgeIndirect,
            IV4 => $included,
            IV5 => $tariffDorG,
        },
    );

    $$capacityChargeRef = 0
      ? Arithmetic(
        arithmetic => '=IF(AND(IV1,IV5="Demand"),IV2+IV3*IV4*100/IV9,0)',
        name       =>
          'Capacity charge after applying fixed adder ex indirects p/kVA/day',
        arguments => {
            IV1 => $included,
            IV5 => $tariffDorG,
            IV2 => $$capacityChargeRef,
            IV3 => Arithmetic(
                name       => 'Fixed adder ex indirects application rate',
                arithmetic =>
                  '=IV1*(IV2-IV7-IV91-IV92)/SUMPRODUCT(IV4_IV5,IV64_IV65)',
                arguments => {
                    IV1       => $ynonFudge41,
                    IV2       => $$shortfallRef,
                    IV3       => $indirect,
                    IV7       => $indirect,
                    IV91      => $direct,
                    IV92      => $rates,
                    IV4_IV5   => $fudgeIndirect,
                    IV64_IV65 => $agreedCapacity,
                },
            ),
            IV4 => $fudgeIndirect,
            IV9 => $daysInYear,
        }
      )
      : Arithmetic(
        arithmetic => '=IF(AND(IV1,IV5="Demand"),IV2+IV3*(IV7+IV4)*100/IV9,0)',
        name       =>
          'Capacity charge after applying fixed adder ex indirects p/kVA/day',
        arguments => {
            IV1 => $included,
            IV5 => $tariffDorG,
            IV2 => $$capacityChargeRef,
            IV3 => (
                my $fixedAdderRate = Arithmetic(
                    name       => 'Fixed adder ex indirects application rate',
                    arithmetic => '=IV1*(IV2-IV7-IV91-IV92)/SUM(IV4_IV5)',
                    arguments  => {
                        IV1     => $ynonFudge41,
                        IV2     => $$shortfallRef,
                        IV3     => $indirect,
                        IV7     => $indirect,
                        IV91    => $direct,
                        IV92    => $rates,
                        IV4_IV5 => $slope,
                    },
                )
            ),
            IV4 => $activeCoincidence,
            IV7 => $ynonFudge,
            IV9 => $daysInYear,
        }
      );

    $model->{summaryInformationColumns}[6] = Arithmetic(
        name          => 'Demand scaling fixed adder (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=IF(AND(IV4,IV5="Demand"),IV1*IV3*(IV71+IV72),0)',
        arguments     => {
            IV1  => $agreedCapacity,
            IV3  => $fixedAdderRate,
            IV71 => $ynonFudge,
            IV72 => $activeCoincidence,
            IV4  => $included,
            IV5  => $tariffDorG,
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
        $model,          $included,       $tariffDorG,
        $agreedCapacity, $shortfall,      $daysInYear,
        $assetsFixed,    $assetsCapacity, $assetsConsumption,
        $capacityCharge, $fixedCharge,
      )
      = @_;

    my $slopeCapacity = Arithmetic(
        name => 'Non sole use notional assets subject to matching (£)',
        defaultFormat => '0softnz',
        arithmetic    => '=IF(AND(IV5,IV6="Demand"),(IV2+IV1)*IV4,0)',
        arguments     => {
            IV5 => $included,
            IV6 => $tariffDorG,
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

    my $slope45 = Arithmetic(
        name       => 'Marginal revenue effect of option 45 cap',
        arithmetic => '=IF(AND(IV1,IV2="Demand"),IF(IV4,0-IV3,0),0)',
        arguments  => {
            IV1 => $included,
            IV2 => $tariffDorG,
            IV3 => $slopeCapacity,
            IV4 => $model->{cap45},
        },
      )
      if $model->{vedcm} == 45;

    my $threshold45 = Arithmetic(
        name       => 'Threshold for option 45 cap',
        arithmetic =>
          '=IF(AND(IV1,IV2="Demand"),IF(IV4,IF(IV7,IV5/IV3,0),0),0)',
        arguments => {
            IV1 => $included,
            IV2 => $tariffDorG,
            IV3 => $slopeCapacity,
            IV7 => $slopeCapacity,
            IV4 => $model->{cap45},
            IV5 => $model->{cap45},
        },
      )
      if $model->{vedcm} == 45;

    # require SpreadsheetModel::SegmentRoot;

    my $demandScaling = Arithmetic(
        name          => 'Annual charge on assets',
        defaultFormat => '%soft',
        arithmetic    => '=IV1/SUM(IV2_IV3)',
        arguments     => { IV1 => $shortfall, IV2_IV3 => $slopeCapacity, }
    );

    0
      and push @{ $model->{demandChargeTables} },
      $demandScaling = new SpreadsheetModel::SegmentRoot(
        name          => 'Demand scaling adder (annual charge on assets)',
        defaultFormat => '%soft',
        target        => $model->{vedcm} == 45
        ? Arithmetic(
            name       => 'Option 45 magic',
            arithmetic => '=IV1+SUMPRODUCT(IV2_IV3,IV4_IV5)',
            arguments  => {
                IV1     => $shortfall,
                IV2_IV3 => $threshold45,
                IV4_IV5 => $slope45
            }
          )
        : $shortfall,
        slopes => Columnset(
            name    => 'Slopes',
            columns => [ $slopeCapacity, ]
        ),
        min => Columnset(
            name    => 'Thresholds',
            columns => [ $minCapacity, ]
        ),
      );

    my $scalingChargeCapacity =
      $model->{vedcm} == 45
      ? Arithmetic(
        arithmetic => '=MAX(0-IV81,IF(AND(IV1,IV5="Demand"),'
          . 'IF(IV71,MIN(IV72,IV31),IV32)*(IV62+IV63)*100/IV9,0))',
        name      => 'Demand scaling p/kVA/day',
        arguments => {
            IV1  => $included,
            IV5  => $tariffDorG,
            IV31 => $demandScaling,
            IV32 => $demandScaling,
            IV62 => $assetsCapacity,
            IV63 => $assetsConsumption,
            IV9  => $daysInYear,
            IV81 => $capacityCharge,
            IV71 => $threshold45,
            IV72 => $threshold45,
        }
      )
      : Arithmetic(
        arithmetic => '=MAX(0-IV81,IF(AND(IV1,IV5="Demand"),'
          . 'IV3*(IV62+IV63)*100/IV9,0))',
        name      => 'Demand scaling p/kVA/day',
        arguments => {
            IV1  => $included,
            IV5  => $tariffDorG,
            IV3  => $demandScaling,
            IV62 => $assetsCapacity,
            IV63 => $assetsConsumption,
            IV9  => $daysInYear,
            IV81 => $capacityCharge,
        }
      );

    $scalingChargeCapacity;

}

sub genScaling {

    my (
        $model,                   $llfcs,          $tariffs,
        $included,                $tariffDorG,     $agreedCapacity,
        $exceededCapacity,        $exportCapacity, $daysInYear,
        $generationChargesExExit, $gPot,
      )
      = @_;

    my $shortfall = Arithmetic(
        name          => 'Generation revenue shortfall (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=IV1-IV2',
        arguments     => {
            IV1 => $gPot,
            IV2 => GroupBy(
                name => 'Total revenue from generation charges (£/year)',
                defaultFormat => '0softnz',
                source        => $generationChargesExExit
            )
        }
    );

    my $slope = Arithmetic(
        name       => 'Marginal revenue effect of generation adder',
        arithmetic => $exceededCapacity
        ? '=IF(AND(IV1,IV2="Generation"),IV4+IV3,0)'
        : '=IF(AND(IV1,IV2="Generation"),IV4,0)',
        arguments => {
            IV1 => $included,
            IV2 => $tariffDorG,
            $exceededCapacity ? ( IV3 => $exceededCapacity ) : (),
            IV4 => $agreedCapacity,
        }
    );

    my $minAdder = Arithmetic(
        name       => 'Generation adder threshold',
        arithmetic => '=IF(AND(IV2,IV3="Generation"),-0.01*IV9*IV1,0)',
        arguments  => {
            IV1 => $exportCapacity,
            IV2 => $included,
            IV9 => $daysInYear,
            IV3 => $tariffDorG
        }
    );

    Columnset(
        name    => 'Prepare data for generation scaling',
        columns => [ $slope, $minAdder, $generationChargesExExit ]
    );

    # require SpreadsheetModel::SegmentRoot;
    my $genScaling = new SpreadsheetModel::SegmentRoot(
        name   => 'General generation adder rate (£/kVA/year)',
        slopes => $slope,
        target => $shortfall,
        min    => $minAdder,
    );

    Arithmetic(
        name       => 'Generation adder p/kVA/day',
        arithmetic => '=IF(AND(IV1,IV2="Generation"),100/IV9*MAX(IV4,IV5),0)',
        arguments  => {
            IV1 => $included,
            IV2 => $tariffDorG,
            IV4 => $genScaling,
            IV5 => $minAdder,
            IV9 => $daysInYear,
        }
    );

}

sub gPot {

    my ( $model, $included, $pre2005, $tariffDorG, $agreedCapacity ) = @_;

    my $actualGenerationIncentive = Dataset(
        name => 'Distributed generation incentive revenue'
          . ' under the price control (GI+GP+GO+GL) (£/year)',
        defaultFormat => '0hard',
        data          => ['#VALUE!'],
        dataset       => $model->{dataset},
    );

    my $rawom = Dataset(
        name => 'Distributed generation incentive O&M'
          . ' charging rate applied in the charging year (£/kW/year)',
        data    => ['#VALUE!'],
        dataset => $model->{dataset},
    );

    my $suf = Dataset(
        name    => 'Distributed generation incentive sole use factor',
        data    => ['#VALUE!'],
        dataset => $model->{dataset},
    );

    my $cdcmpre2005 = Dataset(
        name          => 'CDCM pre-2005 generation capacity (kW)',
        defaultFormat => '0hard',
        data          => ['#VALUE!'],
        dataset       => $model->{dataset},
    );

    my $cdcmpost2005 = Dataset(
        name          => 'CDCM post-2005 generation capacity (kW)',
        defaultFormat => '0hard',
        data          => ['#VALUE!'],
        dataset       => $model->{dataset},
    );

    Columnset(
        name    => 'Generation scaling parameters',
        columns => [
            $actualGenerationIncentive, $rawom,
            $suf,                       $cdcmpre2005,
            $cdcmpost2005,
        ],
        number  => 1144,
        dataset => $model->{dataset},
        $model->{noGen} ? () : ( appendTo => $model->{inputTables} )
    );

    my $pre2005cap = Arithmetic(
        name          => 'Pre-2005 generation capacity (kVA)',
        defaultFormat => '0softnz',
        arithmetic    => '=IF(AND(IV1,IV3="Generation"),IV4*IV2,0)',
        arguments     => {
            IV1 => $included,
            IV2 => $pre2005,
            IV3 => $tariffDorG,
            IV4 => $agreedCapacity
        }
    );

    my $cap = Arithmetic(
        name          => 'Generation capacity (kVA)',
        defaultFormat => '0softnz',
        arithmetic    => '=IF(AND(IV1,IV3="Generation"),IV4,0)',
        arguments     => {
            IV1 => $included,
            IV3 => $tariffDorG,
            IV4 => $agreedCapacity
        }
    );

    Columnset(
        name    => 'Generation capacity data',
        columns => [ $cap, $pre2005cap ]
    );

    my $totalpre2005 = GroupBy(
        name          => 'Total pre-2005 generation capacity (kW)',
        defaultFormat => '0softnz',
        source        => $pre2005cap
    );

    my $total = GroupBy(
        name          => 'Total generation capacity (kW)',
        defaultFormat => '0softnz',
        source        => $cap
    );

    Columnset(
        name    => 'Generation capacity totals',
        columns => [ $total, $totalpre2005 ]
    );

    Arithmetic(
        name          => 'EDCM generation charges revenue target (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=(IV1+IV2*(IV4+IV5))*IV8/(IV9+IV6+IV7)-IV3*IU2*IU8',
        arguments     => {
            IV1 => $actualGenerationIncentive,
            IV2 => $rawom,
            IU2 => $rawom,
            IV3 => $suf,
            IV4 => $totalpre2005,
            IV5 => $cdcmpre2005,
            IV6 => $cdcmpre2005,
            IV7 => $cdcmpost2005,
            IV8 => $total,
            IU8 => $total,
            IV9 => $total,
        }
    );

}

1;
