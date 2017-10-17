﻿package EDCM2::IdiotMitigation;

=head Copyright licence and disclaimer

Copyright 2017 Franck Latrémolière, Reckon LLP and others.

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

sub new {
    my ( $class, $model ) = @_;
    bless { model => $model }, $class;
}

sub notes {
    my ($self) = @_;
    Notes(
        name => 'DNO totals data (special version to migitate'
          . ' the consequences of being taken for an idiot by the DNO)',
        lines => [
            'This version of the model contains adjustments fed from data in'
              . ' tables 119xx to mitigate the impact of the DNO taking you for an'
              . ' idiot by refusing, on some spurious excuse, to provide data from'
              . ' the non-confidential summary sheets in its EDCM charging model.',
            'Documentation of these features does not exist yet.'
              . ' I plan to add a document to dcmf.co.uk/models at some point.',
        ],
    );
}

sub demandRevenuePotAdj {
    my (
        $self,         $totalRevenue3,
        $rateDirect,   $rateRates,
        $rateIndirect, $rateOther,
        $chargeOther,  $totalAssetsCapacity,
        $totalAssetsConsumption,
    ) = @_;
    my $demandRevenuePot = Dataset(
        name          => 'Demand revenue pot (£/year)',
        defaultFormat => '0hard',
        data          => [1e7],
        dataset       => $self->{model}{dataset},
        appendTo      => $self->{model}{inputTables},
        number        => 11951,
        lines         => [
            'Enter your estimate of the demand revenue pot here'
              . ' if your DNO takes you for an idiot and'
              . ' refuses to provide the data for tables 1191 and 1193.',
            'The DNO\'s Schedule 15 published on dcusa.co.uk'
              . ' might help you guess this number.',
            'This number is used in conjunction with data in tables 11962-11965'
              . ' to calculate the DNO\'s total demand EDCM shared assets'
              . ' and total demand EDCM sole use assets.'
        ],
    );
    $self->{constraintDemandRevenuePot} = [
        'not implemented: '
          . 'link between edcmDemandSharedAssets and edcmDemandAssets',
        'edcmDemandSharedAssets/(cdcmSharedAssets+edcmDemandSharedAssets)'
          . '*A1 + edcmDemandAssets*A2 + A3 = 0',
        $demandRevenuePot,
        $totalRevenue3,
        $rateDirect,
        $rateRates,
        $rateIndirect,
        $rateOther,
        $chargeOther,
        $totalAssetsCapacity,
        $totalAssetsConsumption,
        'need more variables, like tranmsission exit',
    ];
    Arithmetic(
        name          => 'Adjusted demand revenue pot (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=IF(ISERROR(A1),A2,A3)',
        arguments     => {
            A1 => $demandRevenuePot,
            A2 => $totalRevenue3,
            A3 => $demandRevenuePot,
        },
    );
}

sub gChargeAdj {
    my ( $self, $gCharge ) = @_;
    my $gChargeHard = Dataset(
        name          => 'Export capacity charge (p/kVA/day)',
        defaultFormat => '0.00hard',
        data          => [0.05],
        dataset       => $self->{model}{dataset},
        appendTo      => $self->{model}{inputTables},
        number        => 11954,
        lines         => [
            'Enter the DNO\'s standard EDCM export capacity charge here,'
              . ' in p/kVA/day,'
              . ' if your DNO takes you for an idiot and'
              . ' refuses to provide the data for table 1192.',
            'The DNO\'s charging statement published on the DNO\'s website'
              . ' might help you guess this number.',
            'This number is used to adjust the figures in table 1192 to fit.',
        ],
    );
    $self->{adjustmentExportCapacityChargeablePost2010} = Arithmetic(
        name => 'Adjustment to EDCM post-2010 non-exempt export capacity (kVA)',
        defaultFormat => '0softpm',
        arithmetic    => '=1/(1/A4-A7*0.01*(A1-A2)/A3/A6)-A5',
        arguments     => {
            A1 => $gChargeHard,
            A2 => $gCharge,
            A3 => $gCharge->{arguments}{A21},
            A4 => $gCharge->{arguments}{A23},
            A5 => $gCharge->{arguments}{A23},
            A6 => $gCharge->{arguments}{A1},
            A7 => $gCharge->{arguments}{A9},
        },
    );
    Arithmetic(
        name       => 'Adjusted export capacity charge (p/kVA/day)',
        arithmetic => '=IF(ISERROR(A1),A2,A3)',
        arguments =>
          { A1 => $gChargeHard, A2 => $gCharge, A3 => $gChargeHard, },
    );
}

sub exitChargeAdj {
    my ( $self, $rateExit, $cdcmPurpleUse, $edcmPurpleUse, $chargeExit,
        $purpleUseRate, $importCapacity, )
      = @_;
    my $tariffIndex = Dataset(
        name          => 'Tariff index',
        defaultFormat => '0hard',
        data          => [1],
    );
    my $charge = Dataset(
        name          => 'Transmission exit charge (£/year)',
        defaultFormat => '0hard',
        data          => [1e4],
    );
    Columnset(
        name => 'Demand transmission exit charge for a non-pathological site',
        columns  => [ $tariffIndex, $charge, ],
        dataset  => $self->{model}{dataset},
        appendTo => $self->{model}{inputTables},
        number   => 11960,
        lines    => [
            'Enter the transmission exit charge for a non-pathological site,'
              . ' in £/year, if your DNO takes you for an idiot and'
              . ' refuses to provide the data for tables 1191 and 1193.',
            'You will need to show authority for the site and to ask the DNO'
              . ' for the HSummary charge breakdown.',
            'This number is used to adjust the figure in the'
              . ' first column of table 1191.',
        ],
    );
    my $rateExitCalculated = Arithmetic(
        name       => 'Calculated transmission exit charging rate (£/kW/year)',
        arithmetic => '=A1/INDEX(A2_A21,A22)/INDEX(A3_A31,A32)',
        arguments  => {
            A1     => $charge,
            A2_A21 => $purpleUseRate,
            A3_A31 => $importCapacity,
            A22    => $tariffIndex,
            A32    => $tariffIndex,
        },
    );
    Arithmetic(
        name       => 'Adjusted transmission exit charging rate (£/kW/year)',
        arithmetic => '=IF(ISERROR(A1),A2,A3)',
        arguments  => {
            A1 => $rateExitCalculated,
            A2 => $rateExit,
            A3 => $rateExitCalculated,
        },
      ),
      $self->{adjustedEdcmPurpleUse} = Arithmetic(
        name          => 'Adjusted total EDCM peak-time consumption (kW)',
        defaultFormat => '0soft',
        arithmetic    => '=A1/A2-A3',
        arguments     => {
            A1 => $chargeExit,
            A2 => $rateExitCalculated,
            A3 => $cdcmPurpleUse,
        },
      );
}

sub fixedChargeAdj {
    my ( $self, $rateDirect, $rateRates, $rateIndirect, $demandSoleUseAsset,
        $chargeDirect, $chargeRates, $chargeIndirect, )
      = @_;
    my $tariffIndex = Dataset(
        name          => 'Tariff index',
        defaultFormat => '0hard',
        data          => [1],
    );
    my $fixedCharge = Dataset(
        name          => 'Demand fixed charge (£/year)',
        defaultFormat => '0hard',
        data          => [1e5],
    );
    Columnset(
        name     => 'Demand fixed charge for a site not affected by DCP 189',
        columns  => [ $tariffIndex, $fixedCharge, ],
        dataset  => $self->{model}{dataset},
        appendTo => $self->{model}{inputTables},
        number   => 11961,
        lines    => [
            'Enter the demand fixed charge for a non-pathological site,'
              . ' in £/year, if your DNO takes you for an idiot and'
              . ' refuses to provide the data for table 1193.',
            'You will need to show authority for the site and to ask the DNO'
              . ' for the HSummary charge breakdown.',
            'This number is used to adjust figures in table 1193'
              . ' to fit the total value of EDCM notional assets.',
        ],
    );
    my $actualRate = Arithmetic(
        name       => 'Actual charging rate for fixed charges',
        arithmetic => '=A1/INDEX(A2_A3,A4)',
        arguments  => {
            A1    => $fixedCharge,
            A2_A3 => $demandSoleUseAsset,
            A4    => $tariffIndex,
        },
    );
    my $f = Arithmetic(
        name       => 'Ratio of direct cost denominator to rates denominator',
        arithmetic => '=A1/A2/A3*A4',
        arguments  => {
            A1 => $chargeDirect,
            A2 => $rateDirect,
            A3 => $chargeRates,
            A4 => $rateRates,
        },
    );
    my $thing = Arithmetic(
        name       => 'Intermediate step in solving for total EDCM assets',
        arithmetic => '=0.5*(1+A1-(A2*A3+A4)/A5)',
        arguments  => {
            A1 => $f,
            A2 => $f,
            A3 => $rateDirect,
            A4 => $rateRates,
            A5 => $actualRate,
        },
    );
    $self->{edcmAssetAdjustment} = Arithmetic(
        name          => 'Adjustment to total EDCM assets',
        defaultFormat => '0softpm',
        arithmetic    => '=A1/A2*(SQRT(A3^2+A4*((A5+A6)/A7-1))-A8)',
        arguments     => {
            A1 => $chargeRates,
            A2 => $rateRates,
            A3 => $thing,
            A4 => $f,
            A5 => $rateDirect,
            A6 => $rateRates,
            A7 => $actualRate,
            A8 => $thing,
        },
    );
    my $rateDirectAdjusted = Arithmetic(
        name          => 'Adjusted direct charging rate',
        defaultFormat => '%soft',
        arithmetic    => '=IF(ISERROR(A11),A21,1/(1/A2+A1/A3))',
        arguments     => {
            A1  => $self->{edcmAssetAdjustment},
            A11 => $self->{edcmAssetAdjustment},
            A2  => $rateDirect,
            A3  => $chargeDirect,
            A21 => $rateDirect,
        },
    );
    my $rateRatesAdjusted = Arithmetic(
        name          => 'Adjusted rates charging rate',
        defaultFormat => '%soft',
        arithmetic    => '=IF(ISERROR(A11),A21,1/(1/A2+A1/A3))',
        arguments     => {
            A1  => $self->{edcmAssetAdjustment},
            A11 => $self->{edcmAssetAdjustment},
            A2  => $rateRates,
            A3  => $chargeRates,
            A21 => $rateRates,
        },
    );
    my $rateIndirectAdjusted = Arithmetic(
        name          => 'Adjusted indirect charging rate',
        defaultFormat => '%soft',
        arithmetic    => '=IF(ISERROR(A11),A21,1/(1/A2+A1/A3))',
        arguments     => {
            A1  => $self->{edcmAssetAdjustment},
            A11 => $self->{edcmAssetAdjustment},
            A2  => $rateIndirect,
            A3  => $chargeIndirect,
            A21 => $rateIndirect,
        },
    );
    $rateDirectAdjusted, $rateRatesAdjusted, $rateIndirectAdjusted;
}

sub indirectChargeAdj {
    my ( $self, $indirectChargingRate, $fudgeIndirect, $agreedCapacity,
        $indirect, )
      = @_;
    my $tariffIndex = Dataset(
        name          => 'Tariff index',
        defaultFormat => '0hard',
        data          => [1],
    );
    my $charge = Dataset(
        name          => 'Indirect cost charge (£/year)',
        defaultFormat => '0hard',
        data          => [1e4],
    );
    Columnset(
        name     => 'Demand indirect cost charge for a site',
        columns  => [ $tariffIndex, $charge, ],
        dataset  => $self->{model}{dataset},
        appendTo => $self->{model}{inputTables},
        number   => 11962,
        lines    => [
            'Enter the indirect cost charge for a non-pathological site,'
              . ' in £/year, if your DNO takes you for an idiot and'
              . ' refuses to provide the data for tables 1191 and 1193.',
            'You will need to show authority for the site and to ask the DNO'
              . ' for the HSummary charge breakdown.',
            'This number is used in conjunction with data from tables'
              . ' 11951, 11963, 11964 and 11965 to adjust the DNO totals to fit.',
        ],
    );
    my $indirectChargingRateCalculated = Arithmetic(
        name       => 'Calculated indirect costs application rate',
        arithmetic => '=A1/INDEX(A2_A3,A4)/INDEX(A5_A6,A7)',
        arguments  => {
            A1    => $charge,
            A2_A3 => $fudgeIndirect,
            A4    => $tariffIndex,
            A5_A6 => $agreedCapacity,
            A7    => $tariffIndex,
        },
    );
    $self->{constraintIndirectChargingRate} = [
        'not implemented: '
          . 'link between edcmDemandAssets and indirectDenominator',
        $indirectChargingRateCalculated, $indirectChargingRate,
        $fudgeIndirect,                  $agreedCapacity,
        $indirect,
    ];
    Arithmetic(
        name       => 'Adjusted indirect costs application rate',
        arithmetic => '=IF(ISERROR(A1),A2,A3)',
        arguments  => {
            A1 => $indirectChargingRateCalculated,
            A2 => $indirectChargingRate,
            A3 => $indirectChargingRateCalculated,
        },
    );
}

sub fixedAdderAdj {
    my ( $self, $fixedAdderChargingRate, $slope, $ynonFudge41, $adderAmount, )
      = @_;
    my $tariffIndex = Dataset(
        name          => 'Tariff index',
        defaultFormat => '0hard',
        data          => [1],
    );
    my $charge = Dataset(
        name          => 'Fixed adder charge (£/year)',
        defaultFormat => '0hard',
        data          => [1e4],
    );
    Columnset(
        name     => 'Demand fixed adder charge for a non-pathological site',
        columns  => [ $tariffIndex, $charge, ],
        dataset  => $self->{model}{dataset},
        appendTo => $self->{model}{inputTables},
        number   => 11963,
        lines    => [
            'Enter the fixed adder charge for a non-pathological site,'
              . ' in £/year, if your DNO takes you for an idiot and'
              . ' refuses to provide the data for tables 1191 and 1193.',
            'You will need to show authority for the site and to ask the DNO'
              . ' for the HSummary charge breakdown.',
            'This number is used in conjunction with data from tables'
              . ' 11951, 11962, 11964 and 11965 to adjust the DNO totals to fit.',
        ],
    );
    $self->{constraintFixedAdderRate} = [
            'not implemented: '
          . 'link between residualResidual and indirectDenominator',
    ];
    $fixedAdderChargingRate;
}

sub assetAdderAdj {
    my ( $self, $demandScaling, $totalSlopeCapacity,
        $slopeCapacity, $shortfall, )
      = @_;
    my $tariffIndex = Dataset(
        name          => 'Tariff index',
        defaultFormat => '0hard',
        data          => [1],
    );
    my $adderCharge = Dataset(
        name          => 'Asset adder charge (£/year)',
        defaultFormat => '0hard',
        data          => [1e4],
    );
    Columnset(
        name     => 'Demand asset adder for a non-pathological non-0000 site',
        columns  => [ $tariffIndex, $adderCharge, ],
        dataset  => $self->{model}{dataset},
        appendTo => $self->{model}{inputTables},
        number   => 11964,
        lines    => [
            'Enter the asset adder charge for a non-pathological non-0000 site,'
              . ' in £/year, if your DNO takes you for an idiot and'
              . ' refuses to provide the data for tables 1191 and 1193.',
            'You will need to show authority for the site and to ask the DNO'
              . ' for the HSummary charge breakdown.',
            'This number is used in conjunction with data from tables'
              . ' 11951, 11962, 11963 and 11965 to adjust the DNO totals to fit.',
        ],
    );
    $self->{cookedSharedAssetAdjustment} = Arithmetic(
        name          => 'Adjustment to shared assets subject to adder (£)',
        defaultFormat => '0softpm',
        arithmetic    => '=A1/(A2/INDEX(A3_A4,A5))-A6',
        arguments     => {
            A1    => $shortfall,
            A2    => $adderCharge,
            A3_A4 => $slopeCapacity,
            A5    => $tariffIndex,
            A6    => $totalSlopeCapacity,
        },
    );
    $self->{constraintAssetAdderRate} = [
            'not implemented: '
          . 'link between residualResidual and assetDenominator',
    ];
    Arithmetic(
        name          => 'Adjusted annual charge on assets',
        defaultFormat => '%soft',
        arithmetic    => '=IF(ISERROR(A11),A21,1/(1/A2+A1/A3))',
        arguments     => {
            A1  => $self->{cookedSharedAssetAdjustment},
            A11 => $self->{cookedSharedAssetAdjustment},
            A2  => $demandScaling,
            A3  => $shortfall,
            A21 => $demandScaling,
        },
    );
}

sub directCostAdj {
    my ( $self, $direct, $rates, ) = @_;
    my $tariffIndex = Dataset(
        name          => 'Tariff index',
        defaultFormat => '0hard',
        data          => [1],
    );
    my $charge = Dataset(
        name          => 'Direct cost charge (£/year)',
        defaultFormat => '0hard',
        data          => [1e4],
    );
    Columnset(
        name =>
          'Demand direct cost charge for a non-pathological non-0000 site',
        columns  => [ $tariffIndex, $charge, ],
        dataset  => $self->{model}{dataset},
        appendTo => $self->{model}{inputTables},
        number   => 11965,
        lines    => [
            'Enter the direct cost charge for a non-pathological non-0000 site,'
              . ' in £/year, if your DNO takes you for an idiot and'
              . ' refuses to provide the data for tables 1191 and 1193.',
            'You will need to show authority for the site and to ask the DNO'
              . ' for the HSummary charge breakdown.',
            'This number is used in conjunction with data from tables'
              . ' 11951, 11962, 11963 and 11964 to adjust the DNO totals to fit.',
        ],
    );
    $self->{constraintDirectCost} = [
            'not implemented: '
          . 'link between edcmDemandSharedAssets and assetDenominator',
    ];
    $direct, $rates;
}

sub calcTables {

    my ($self) = @_;

    # do the calculations here, except the final DNO total adjustment step
    # use:
    # $self->{constraintAssetAdderRate}
    # $self->{constraintDemandRevenuePot}
    # $self->{constraintDirectCost}
    # $self->{constraintFixedAdderRate}
    # $self->{constraintIndirectChargingRate}

    [
        grep { $_; } @{$self}{
            qw(
              adjustedEdcmPurpleUse
              adjustmentExportCapacityChargeablePost2010
              )
        }
    ];

}

sub adjustDnoTotals {

    my ( $self, $model, $hashedArrays, ) = @_;

    if ( $self->{adjustedEdcmPurpleUse} ) {
        $hashedArrays->{1191}[0] = Arithmetic(
            name          => 'Adjusted total EDCM peak time consumption (kW)',
            defaultFormat => '0soft',
            arithmetic    => '=IF(ISERROR(A1),A2,A3)',
            arguments     => {
                A1 => $self->{adjustedEdcmPurpleUse},
                A2 => $hashedArrays->{1191}[0]{sources}[0],
                A3 => $self->{adjustedEdcmPurpleUse},
            },
        );
    }

    if ( $self->{adjustmentExportCapacityChargeablePost2010} ) {
        $hashedArrays->{1192}[0] = Arithmetic(
            name => 'Adjusted chargeable export capacity baseline (kVA)',
            defaultFormat => '0soft',
            arithmetic    => '=A1+IF(ISERROR(A2),0,A3)',
            arguments     => {
                A1 => $hashedArrays->{1192}[0]{sources}[0],
                A2 => $self->{adjustmentExportCapacityChargeablePost2010},
                A3 => $self->{adjustmentExportCapacityChargeablePost2010},
            },
        );
        $hashedArrays->{1192}[2] = Arithmetic(
            name =>
              'Adjusted non-exempt post-2010 export capacity baseline (kVA)',
            defaultFormat => '0soft',
            arithmetic    => '=A1+IF(ISERROR(A2),0,A3)',
            arguments     => {
                A1 => $hashedArrays->{1192}[2]{sources}[0],
                A2 => $self->{adjustmentExportCapacityChargeablePost2010},
                A3 => $self->{adjustmentExportCapacityChargeablePost2010},
            },
        );
    }

}

1;
