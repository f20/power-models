package EDCM2::SecrecyMitigation;

# Copyright 2017-2020 Franck Latrémolière and others.
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

sub new {
    my ( $class, $model ) = @_;
    bless { model => $model }, $class;
}

sub additionalLines {
    my ($self) = @_;
    'This workbook contains experimental features that use'
      . ' data provided in tables 119xx to replace the data'
      . ' in tables 1191-1193 with improved estimates. '
      . 'To get a good fit, you need to get from'
      . ' the DNO the contents of tables 935 and'
      . ' demand capacity for a non-pathological EDCM tariff.';
}

sub fitTotalAssets {
    my ($self) = @_;
    my $proportionDemandAssetsWhichAreShared = Arithmetic(
        name => 'Proportion of EDCM demand assets which are not sole use',
        defaultFormat => '%soft',
        arithmetic    => '=A1/A2*A3/A41*(1-A42)/A5/A6',
        arguments     => {
            A1  => $self->{rateIndirect},
            A2  => $self->{indirectChargingRate},
            A3  => $self->{fixedAdderChargingRate},
            A41 => $self->{fudge41param},
            A42 => $self->{fudge41param},
            A5  => $self->{assetAdderChargingRate},
            A6  => $self->{assetCookingRatio},
        },
    );
    my $D = Arithmetic(
        name =>
          'Pot correction required if no change to shared demand assets (£)',
        defaultFormat => '0softpm',
        arithmetic    => '=A1-A2+(A3+A4+A5)*(A9-(A7+A8)*(1/A6-1))',
        arguments     => {
            A1 => $self->{demandRevenuePot},
            A2 => $self->{totalRevenue3},
            A3 => $self->{rateDirect},
            A4 => $self->{rateRates},
            A5 => $self->{rateIndirect},
            A6 => $proportionDemandAssetsWhichAreShared,
            A7 => $self->{totalAssetsCapacity},
            A8 => $self->{totalAssetsConsumption},
            A9 => $self->{totalAssetsDemandSoleUse},
        },
    );
    my $F = Arithmetic(
        name => 'Cost-based pot contribution of demand assets'
          . ' if no change to shared demand assets (£)',
        defaultFormat => '0soft',
        arithmetic    => '=(A3+A4+A5)*(A7+A8)/A6',
        arguments     => {
            A3 => $self->{rateDirect},
            A4 => $self->{rateRates},
            A5 => $self->{rateIndirect},
            A6 => $proportionDemandAssetsWhichAreShared,
            A7 => $self->{totalAssetsCapacity},
            A8 => $self->{totalAssetsConsumption},
        },
    );
    my $e = Arithmetic(
        name => 'Intermediate step in solving quadratic equation'
          . ' for EDCM demand shared assets',
        arithmetic =>
          '=0.5*(A11/A12/(A13+A14)+A15/A16-A17*(A18+A19)/A20-A21/A22)',
        arguments => {
            A11 => $self->{chargeOther},
            A12 => $self->{rateOther},
            A13 => $self->{totalAssetsCapacity},
            A14 => $self->{totalAssetsConsumption},
            A15 => $self->{chargeOther},
            A16 => $F,
            A17 => $self->{rateOther},
            A18 => $self->{totalAssetsCapacity},
            A19 => $self->{totalAssetsConsumption},
            A20 => $F,
            A21 => $D,
            A22 => $F,
        },
    );
    $self->{scalerSharedDemandAssets} = Arithmetic(
        name          => 'Percentage adjustment to EDCM demand shared assets',
        defaultFormat => '%softpm',
        arithmetic    => '=SQRT(A1*A2+A4/A5/(A7+A8)*A6/A9)-A3',
        arguments     => {
            A1 => $e,
            A2 => $e,
            A3 => $e,
            A4 => $self->{chargeOther},
            A5 => $self->{rateOther},
            A6 => $D,
            A9 => $F,
            A7 => $self->{totalAssetsCapacity},
            A8 => $self->{totalAssetsConsumption},
        },
    );
    $self->{adjustmentTotalDemandAssets} = Arithmetic(
        name          => 'Adjustment to total EDCM demand assets (£)',
        defaultFormat => '0softpm',
        arithmetic    => '=(1+A2)*(A4+A5)/A1-A3-A41-A51',
        arguments     => {
            A1  => $proportionDemandAssetsWhichAreShared,
            A2  => $self->{scalerSharedDemandAssets},
            A3  => $self->{totalAssetsDemandSoleUse},
            A4  => $self->{totalAssetsCapacity},
            A41 => $self->{totalAssetsCapacity},
            A5  => $self->{totalAssetsConsumption},
            A51 => $self->{totalAssetsConsumption},
        },
    );
    $self->{adjustedFixedIndirectDenominator} = Arithmetic(
        name => 'Fitted charging base for fixed adder and indirects (kVA)',
        defaultFormat => '0soft',
        arithmetic    => '=(A1+A2+A3+A4)*A5/A6',
        arguments     => {
            A1 => $self->{adjustmentTotalDemandAssets},
            A2 => $self->{totalAssetsDemandSoleUse},
            A3 => $self->{totalAssetsCapacity},
            A4 => $self->{totalAssetsConsumption},
            A5 => $self->{rateIndirect},
            A6 => $self->{indirectChargingRate},
        },
    );
    my $costBasedRevenue = Arithmetic(
        name          => 'Revenue from cost charges on EDCM demand (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=(A1+A2+A3+A4)*(A5+A6+A7)+A8*A9',
        arguments     => {
            A1 => $self->{adjustmentTotalDemandAssets},
            A2 => $self->{totalAssetsDemandSoleUse},
            A3 => $self->{totalAssetsCapacity},
            A4 => $self->{totalAssetsConsumption},
            A5 => $self->{rateDirect},
            A6 => $self->{rateRates},
            A7 => $self->{rateIndirect},
            A8 => $self->{rateExitCalculated},
            A9 => $self->{adjustedEdcmPurpleUse},
        },
    );
    my $fixedAdderRevenue = Arithmetic(
        name          => 'Revenue from fixed adder (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=A1*A2',
        arguments     => {
            A1 => $self->{adjustedFixedIndirectDenominator},
            A2 => $self->{fixedAdderChargingRate},
        },
    );
    my $assetAdderRevenue = Arithmetic(
        name          => 'Revenue from asset adder (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=A1*(1+A2)*(A3+A4)*A5',
        arguments     => {
            A1 => $self->{assetCookingRatio},
            A2 => $self->{scalerSharedDemandAssets},
            A3 => $self->{totalAssetsCapacity},
            A4 => $self->{totalAssetsConsumption},
            A5 => $self->{assetAdderChargingRate},
        },
    );
    $self->{adjustedRevenueFromCharge1} = Arithmetic(
        name          => 'Fitted revenue from charge 1 (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=A1-A2-A3-A4',
        arguments     => {
            A1 => $self->{demandRevenuePot},
            A2 => $costBasedRevenue,
            A3 => $assetAdderRevenue,
            A4 => $fixedAdderRevenue,
        },
    );
}

sub calcTables {
    my ($self) = @_;
    eval { $self->fitTotalAssets; };
    warn "Could not fit total DNO assets: $@" if $@;
    [
        grep { $_; } @{$self}{
            qw(adjustedEdcmPurpleUse adjustmentExportCapacityChargeablePost2010)
        }
    ],
      [ grep { $_; } @{$self}{qw(adjustmentTotalEdcmAssets)} ],
      [ grep { $_; }
          @{$self}{qw(interimRecookEdcmDirect interimRecookEdcmRates)} ], [
        grep { $_; } @{$self}{
            qw(adjustmentTotalDemandAssets scalerSharedDemandAssets
              adjustedFixedIndirectDenominator
              adjustedRevenueFromCharge1
              assetCookingRatio)
        }
          ];
}

sub adjustDnoTotals {

    my ( $self, $model, $hashedArrays, ) = @_;

    $hashedArrays->{1191}[0] = Arithmetic(
        name          => 'Adjusted total EDCM peak time consumption (kW)',
        defaultFormat => '0soft',
        arithmetic    => '=IF(ISERROR(A1),A2,A3)',
        arguments     => {
            A1 => $self->{adjustedEdcmPurpleUse},
            A2 => $hashedArrays->{1191}[0]{sources}[0],
            A3 => $self->{adjustedEdcmPurpleUse},
        },
    ) if $self->{adjustedEdcmPurpleUse};

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

    if ( $self->{adjustedFixedIndirectDenominator} ) {
        $hashedArrays->{1191}[1] = Arithmetic(
            name =>
              'Adjusted total marginal effect of indirect cost adder (kVA)',
            defaultFormat => '0soft',
            arithmetic    => '=IF(ISERROR(A1),A2,A3)',
            arguments     => {
                A1 => $self->{adjustedFixedIndirectDenominator},
                A2 => $hashedArrays->{1191}[1]{sources}[0],
                A3 => $self->{adjustedFixedIndirectDenominator},
            },
        );
        $hashedArrays->{1191}[2] = Arithmetic(
            name =>
              'Adjusted total marginal effect of demand fixed adder (kVA)',
            defaultFormat => '0soft',
            arithmetic    => '=IF(ISERROR(A1),A2,A3)',
            arguments     => {
                A1 => $self->{adjustedFixedIndirectDenominator},
                A2 => $hashedArrays->{1191}[2]{sources}[0],
                A3 => $self->{adjustedFixedIndirectDenominator},
            },
        );
    }

    $hashedArrays->{1191}[3] = Arithmetic(
        name          => 'Adjusted total revenue from charge 1 (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=IF(ISERROR(A1),A2,A3)',
        arguments     => {
            A1 => $self->{adjustedRevenueFromCharge1},
            A2 => $hashedArrays->{1191}[3]{sources}[0],
            A3 => $self->{adjustedRevenueFromCharge1},
        },
    ) if $self->{adjustedRevenueFromCharge1};

    $hashedArrays->{1193}[0] = Arithmetic(
        name          => 'Adjusted demand sole use assets (£)',
        defaultFormat => '0soft',
        arithmetic => '=A1+IF(ISERROR(A2),0,A3)-IF(ISERROR(A4),0,A5*(A6+A7))',
        arguments  => {
            A1 => $hashedArrays->{1193}[0]{sources}[0],
            A2 => $self->{adjustmentTotalDemandAssets},
            A3 => $self->{adjustmentTotalDemandAssets},
            A4 => $self->{scalerSharedDemandAssets},
            A5 => $self->{scalerSharedDemandAssets},
            A6 => $hashedArrays->{1193}[2]{sources}[0],
            A7 => $hashedArrays->{1193}[3]{sources}[0],
        },
      )
      if $self->{adjustmentTotalDemandAssets}
      && $self->{scalerSharedDemandAssets};

    $hashedArrays->{1193}[1] = Arithmetic(
        name          => 'Adjusted generation sole use assets (£)',
        defaultFormat => '0soft',
        arithmetic    => '=A1+IF(ISERROR(A2),0,A3)-IF(ISERROR(A4),0,A5)',
        arguments     => {
            A1 => $hashedArrays->{1193}[1]{sources}[0],
            A2 => $self->{adjustmentTotalEdcmAssets},
            A3 => $self->{adjustmentTotalEdcmAssets},
            A4 => $self->{adjustmentTotalDemandAssets},
            A5 => $self->{adjustmentTotalDemandAssets},
        },
      )
      if $self->{adjustmentTotalEdcmAssets}
      && $self->{adjustmentTotalDemandAssets};

    $hashedArrays->{1193}[2] = Arithmetic(
        name          => 'Adjusted demand capacity assets (£)',
        defaultFormat => '0soft',
        arithmetic    => '=A1*(1+IF(ISERROR(A2),0,A3))',
        arguments     => {
            A1 => $hashedArrays->{1193}[2]{sources}[0],
            A2 => $self->{scalerSharedDemandAssets},
            A3 => $self->{scalerSharedDemandAssets},
        },
    ) if $self->{scalerSharedDemandAssets};

    $hashedArrays->{1193}[3] = Arithmetic(
        name          => 'Adjusted demand consumption assets (£)',
        defaultFormat => '0soft',
        arithmetic    => '=A1*(1+IF(ISERROR(A2),0,A3))',
        arguments     => {
            A1 => $hashedArrays->{1193}[3]{sources}[0],
            A2 => $self->{scalerSharedDemandAssets},
            A3 => $self->{scalerSharedDemandAssets},
        },
    ) if $self->{scalerSharedDemandAssets};

    $hashedArrays->{1193}[4] = Arithmetic(
        name          => 'Adjusted total assets subject to asset adder (£)',
        defaultFormat => '0soft',
        arithmetic    => '=IF(ISERROR(A1),A2,A3*(A4+A5))',
        arguments     => {
            A1 => $self->{assetCookingRatio},
            A2 => $hashedArrays->{1193}[4]{sources}[0],
            A3 => $self->{assetCookingRatio},
            A4 => $hashedArrays->{1193}[2],
            A5 => $hashedArrays->{1193}[3],
        },
    ) if $self->{assetCookingRatio};

}

sub interimRecookTotals {
    my ( $self, $demandScalingShortfall, $edcmDirect, $edcmRates, ) = @_;
    $self->{interimRecookEdcmDirect} = Arithmetic(
        name => 'Interim re-estimate of direct costs'
          . ' charged on EDCM demand shared assets (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=IF(ISERROR(A1/A2),A3,A4*A5/A6)',
        arguments     => {
            A1 => $self->{directCostChargingRate},
            A2 => $self->{overallAssetChargingRate},
            A3 => $edcmDirect,
            A4 => $demandScalingShortfall,
            A5 => $self->{directCostChargingRate},
            A6 => $self->{overallAssetChargingRate},
        },
      ),
      $self->{interimRecookEdcmRates} = Arithmetic(
        name => 'Interim re-estimate of rates'
          . ' charged on EDCM demand shared assets (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=IF(ISERROR(A1/A2),A3,A4*A5/A6)',
        arguments     => {
            A1 => $self->{ratesChargingRate},
            A2 => $self->{overallAssetChargingRate},
            A3 => $edcmRates,
            A4 => $demandScalingShortfall,
            A5 => $self->{ratesChargingRate},
            A6 => $self->{overallAssetChargingRate},
        },
      );
}

sub fudge41param {
    my ( $self, $fudge41param ) = @_;
    $self->{fudge41param} = $fudge41param;
}

sub demandRevenuePotAdj {
    my (
        $self,                $totalRevenue3,
        $rateDirect,          $rateRates,
        $rateIndirect,        $rateOther,
        $chargeOther,         $totalAssetsDemandSoleUse,
        $totalAssetsCapacity, $totalAssetsConsumption,
    ) = @_;
    $self->{demandRevenuePot} = Dataset(
        name          => 'Demand revenue pot (£/year)',
        defaultFormat => '0hard',
        data          => [1e7],
        dataset       => $self->{model}{dataset},
        appendTo      => $self->{model}{inputTables},
        number        => 11951,
        lines         => [
            'Enter your estimate of the demand revenue pot here'
              . ' if the DNO refuses to provide the data for tables 1191 and 1193.',
            'The DNO\'s Schedule 15 published on dcusa.co.uk'
              . ' might help you guess this number.',
            'This number is used in conjunction with data in tables 1196X'
              . ' to calculate the DNO\'s total demand EDCM shared assets'
              . ' and total demand EDCM sole use assets.'
        ],
    );
    $self->{rateDirect}               = $rateDirect;
    $self->{rateRates}                = $rateRates;
    $self->{rateIndirect}             = $rateIndirect;
    $self->{rateOther}                = $rateOther;
    $self->{chargeOther}              = $chargeOther;
    $self->{totalAssetsDemandSoleUse} = $totalAssetsDemandSoleUse;
    $self->{totalAssetsCapacity}      = $totalAssetsCapacity;
    $self->{totalAssetsConsumption}   = $totalAssetsConsumption;
    $self->{totalRevenue3}            = $totalRevenue3;
    Arithmetic(
        name          => 'Adjusted demand revenue pot (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=IF(ISERROR(A1),A2,A3)',
        arguments     => {
            A1 => $self->{demandRevenuePot},
            A2 => $totalRevenue3,
            A3 => $self->{demandRevenuePot},
        },
    );
}

sub directAndRatesChargeAdj {
    my ( $self, $rateDirect, $rateRates, $assetsCapacity, $assetsConsumption,
        $importCapacity, )
      = @_;
    my $tariffIndex = Dataset(
        name          => 'Tariff index',
        defaultFormat => '0hard',
        data          => [1],
    );
    my $charge = Dataset(
        name          => 'Direct cost and rates charge (£/year)',
        defaultFormat => '0hard',
        data          => [1e4],
    );
    Columnset(
        name => 'Demand capacity direct cost and rates'
          . ' charge for a non-pathological non-0000 site',
        columns  => [ $tariffIndex, $charge, ],
        dataset  => $self->{model}{dataset},
        appendTo => $self->{model}{inputTables},
        number   => 11966,
        lines    => [
            'Enter the sum of the direct cost and rates charge (£/year) for a'
              . ' non-pathological non-0000 site, if the DNO refuses'
              . ' to provide the DNO totals data for tables 1191 and 1193.',
            'You will need to show authority for the site and to ask the DNO'
              . ' for the table 935 input data and the demand capacity charge breakdown.',
            'This number is used in conjunction with data from tables'
              . ' 1195X to adjust the DNO totals to fit.',
        ],
    );
    $self->{directCostChargingRate} = Arithmetic(
        name          => 'Charging rate for direct cost element of asset adder',
        defaultFormat => '%soft',
        arithmetic =>
'=A1/((INDEX(A2_A22,A92)+INDEX(A3_A33,A93))*INDEX(A4_A44,A94))/(1+A5/A6)',
        arguments => {
            A1     => $charge,
            A2_A22 => $assetsCapacity,
            A3_A33 => $assetsConsumption,
            A4_A44 => $importCapacity,
            A92    => $tariffIndex,
            A93    => $tariffIndex,
            A94    => $tariffIndex,
            A5     => $rateRates,
            A6     => $rateDirect,
        },
    );
    $self->{ratesChargingRate} = Arithmetic(
        name          => 'Charging rate for rates element of asset adder',
        defaultFormat => '%soft',
        arithmetic    => '=A1/A2*A3',
        arguments     => {
            A1 => $self->{directCostChargingRate},
            A2 => $rateDirect,
            A3 => $rateRates,
        },
    );
    $self->{assetCookingRatio} = Arithmetic(
        name => 'Ratio of asset adder denominator'
          . ' to total EDCM demand shared assets',
        arithmetic => '=A5/A1',
        arguments  => {
            A1 => $self->{directCostChargingRate},
            A5 => $rateDirect,
        },
    );
}

sub assetAdderAdj {
    my ( $self, $demandScaling, $slope, ) = @_;
    my $tariffIndex = Dataset(
        name          => 'Tariff index',
        defaultFormat => '0hard',
        data          => [1],
    );
    my $charge = Dataset(
        name          => 'Asset adder charge (£/year)',
        defaultFormat => '0hard',
        data          => [1e4],
    );
    Columnset(
        name     => 'Demand asset adder for a non-pathological non-0000 site',
        columns  => [ $tariffIndex, $charge, ],
        dataset  => $self->{model}{dataset},
        appendTo => $self->{model}{inputTables},
        number   => 11964,
        lines    => [
            'Enter the asset adder charge (£/year) for a'
              . ' non-pathological non-0000 site, if the DNO refuses'
              . ' to provide the DNO totals data for tables 1191 and 1193.',
            'You will need to show authority for the site and to ask the DNO'
              . ' for the table 935 input data and the demand capacity charge breakdown.',
            'This number is used in conjunction with data from tables'
              . ' 1195X to adjust the DNO totals to fit.',
        ],
    );
    $self->{assetAdderChargingRate} = Arithmetic(
        name          => 'Charging rate for asset adder',
        defaultFormat => '%soft',
        arithmetic    => '=A1/INDEX(A2_A3,A4)',
        arguments     => {
            A1    => $charge,
            A2_A3 => $slope,
            A4    => $tariffIndex,
        },
    );
    $self->{overallAssetChargingRate} = Arithmetic(
        name          => 'Overall charge on (cooked) assets',
        defaultFormat => '%soft',
        arithmetic    => '=A1+A2+A3',
        arguments     => {
            A1 => $self->{assetAdderChargingRate},
            A2 => $self->{directCostChargingRate},
            A3 => $self->{ratesChargingRate},
        },
    );
    Arithmetic(
        name          => 'Adjusted charge on assets',
        defaultFormat => '%soft',
        arithmetic    => '=IF(ISERROR(A1),A2,A3)',
        arguments     => {
            A1 => $self->{overallAssetChargingRate},
            A2 => $demandScaling,
            A3 => $self->{overallAssetChargingRate},
        },
    );
}

sub fixedAdderAdj {
    my ( $self, $fixedAdderChargingRate, $slope ) = @_;
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
            'Enter the fixed adder charge (£/year) for a'
              . ' non-pathological site, if the DNO refuses'
              . ' to provide the DNO totals data for tables 1191 and 1193.',
            'You will need to show authority for the site and to ask the DNO'
              . ' for the table 935 input data and the demand capacity charge breakdown.',
            'This number is used in conjunction with data from tables'
              . ' 1195X to adjust the DNO totals to fit.',
        ],
    );
    $self->{fixedAdderChargingRate} = Arithmetic(
        name       => 'Charging rate for fixed adder',
        arithmetic => '=A1/INDEX(A2_A3,A4)',
        arguments  => {
            A1    => $charge,
            A2_A3 => $slope,
            A4    => $tariffIndex,
        },
    );
    Arithmetic(
        name       => 'Adjusted fixed adder charging rate',
        arithmetic => '=IF(ISERROR(A1),A2,A3)',
        arguments  => {
            A1 => $self->{fixedAdderChargingRate},
            A2 => $fixedAdderChargingRate,
            A3 => $self->{fixedAdderChargingRate},
        },
    );
}

sub indirectChargeAdj {
    my ( $self, $indirectChargingRate, $fudgeIndirect, $agreedCapacity,
        $edcmIndirect )
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
            'Enter the indirect cost charge (£/year) for a'
              . ' non-pathological site, if the DNO refuses'
              . ' to provide the DNO totals data for tables 1191 and 1193.',
            'You will need to show authority for the site and to ask the DNO'
              . ' for the table 935 input data and the demand capacity charge breakdown.',
            'This number is used in conjunction with data from tables'
              . ' 1195X to adjust the DNO totals to fit.',
        ],
    );
    $self->{edcmIndirect}         = $edcmIndirect;
    $self->{indirectChargingRate} = Arithmetic(
        name       => 'Charging rate for indirect costs',
        arithmetic => '=A1/INDEX(A2_A3,A4)/INDEX(A5_A6,A7)',
        arguments  => {
            A1    => $charge,
            A2_A3 => $fudgeIndirect,
            A4    => $tariffIndex,
            A5_A6 => $agreedCapacity,
            A7    => $tariffIndex,
        },
    );
    Arithmetic(
        name       => 'Adjusted indirect costs charging rate',
        arithmetic => '=IF(ISERROR(A1),A2,A3)',
        arguments  => {
            A1 => $self->{indirectChargingRate},
            A2 => $indirectChargingRate,
            A3 => $self->{indirectChargingRate},
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
              . ' if the DNO refuses to provide the data for table 1192.',
            'The DNO\'s charging statement published on the DNO\'s website'
              . ' might help you find this number.',
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
            'Enter the transmission exit charge (£/year) for a'
              . ' non-pathological site, if the DNO refuses'
              . ' to provide the DNO totals data for table 1191.',
            'You will need to show authority for the site and to ask the DNO'
              . ' for the table 935 input data and the demand capacity charge breakdown.',
            'This number is used to adjust the figure in the'
              . ' first column of table 1191.',
        ],
    );
    $self->{rateExitCalculated} = Arithmetic(
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
    my $rateExitToUse = Arithmetic(
        name       => 'Adjusted transmission exit charging rate (£/kW/year)',
        arithmetic => '=IF(ISERROR(A1),A2,A3)',
        arguments  => {
            A1 => $self->{rateExitCalculated},
            A2 => $rateExit,
            A3 => $self->{rateExitCalculated},
        },
    );
    $rateExitToUse,
      $self->{adjustedEdcmPurpleUse} = Arithmetic(
        name          => 'Adjusted total EDCM peak-time consumption (kW)',
        defaultFormat => '0soft',
        arithmetic    => '=A1/A2-A3',
        arguments     => {
            A1 => $chargeExit,
            A2 => $rateExitToUse,
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
            'Enter the demand fixed charge (£/year) for a'
              . ' non-pathological site, if the DNO refuses'
              . ' to provide the DNO totals data for table 1193.',
            'You will need to show authority for the site and to ask the DNO'
              . ' for the table 935 input data and the demand capacity charge breakdown.',
            'This number is used to adjust figures in table 1193'
              . ' to fit the total value of EDCM notional assets.',
        ],
    );
    my $actualRate = Arithmetic(
        name          => 'Actual charging rate for fixed charges',
        defaultFormat => '%soft',
        arithmetic    => '=A1/INDEX(A2_A3,A4)',
        arguments     => {
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
        name => 'Intermediate step in solving quadratic'
          . ' equation for total EDCM assets',
        arithmetic => '=0.5*(1+A1-(A2*A3+A4)/A5)',
        arguments  => {
            A1 => $f,
            A2 => $f,
            A3 => $rateDirect,
            A4 => $rateRates,
            A5 => $actualRate,
        },
    );
    $self->{adjustmentTotalEdcmAssets} = Arithmetic(
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
            A1  => $self->{adjustmentTotalEdcmAssets},
            A11 => $self->{adjustmentTotalEdcmAssets},
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
            A1  => $self->{adjustmentTotalEdcmAssets},
            A11 => $self->{adjustmentTotalEdcmAssets},
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
            A1  => $self->{adjustmentTotalEdcmAssets},
            A11 => $self->{adjustmentTotalEdcmAssets},
            A2  => $rateIndirect,
            A3  => $chargeIndirect,
            A21 => $rateIndirect,
        },
    );
    $rateDirectAdjusted, $rateRatesAdjusted, $rateIndirectAdjusted;
}

1;
