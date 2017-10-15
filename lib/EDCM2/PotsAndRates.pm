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

sub demandPot {

    my (
        $model,               $allowedRevenue,
        $chargeDirect,        $chargeRates,
        $chargeIndirect,      $chargeExit,
        $rateDirect,          $rateRates,
        $rateIndirect,        $rateExit,
        $generationRevenue,   $totalDcp189DiscountedAssets,
        $totalAssetsCapacity, $totalAssetsConsumption,
        $cdcmEhvAssets,       $cdcmHvLvShared,
        $demandSoleUseAsset,  $edcmPurpleUse,
        $totalAssetsFixed,    $daysInYear,
        $assetsCapacity,      $assetsConsumption,
        $purpleUseRate,       $importCapacity,
    ) = @_;

    my $chargeOther = Arithmetic(
        name => 'Revenue less costs and '
          . (
            !$totalDcp189DiscountedAssets
              || $model->{dcp189} =~ /preservePot/i
            ? 'net forecast EDCM generation revenue'
            : 'adjustments'
          )
          . ' (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=A1-A2-A3-A4-A5'
          . ( $model->{table1101} ? '-A9' : '' )
          . (
            !$totalDcp189DiscountedAssets
              || $model->{dcp189} =~ /preservePot/i ? ''
            : '+A31*A32'
          ),
        arguments => {
            A1 => $allowedRevenue,
            $model->{table1101} ? ( A9 => $chargeExit ) : (),
            A2 => $chargeDirect,
            A3 => $chargeIndirect,
            A4 => $chargeRates,
            A5 => $generationRevenue,
            $totalDcp189DiscountedAssets
            ? (
                A31 => $rateDirect,
                A32 => $totalDcp189DiscountedAssets,
              )
            : (),
        }
    );
    $model->{transparency}{dnoTotalItem}{1248} = $chargeOther
      if $model->{transparency};

    my $rateOther = Arithmetic(
        name          => 'Other revenue charging rate',
        groupName     => 'Other revenue charging rate',
        arithmetic    => '=A1/(A21+A22+A3+A4)',
        defaultFormat => '%soft',
        arguments     => {
            A1  => $chargeOther,
            A21 => $totalAssetsCapacity,
            A22 => $totalAssetsConsumption,
            A3  => $cdcmEhvAssets,
            A4  => $cdcmHvLvShared,
        },
        location => 'Charging rates',
    );
    $model->{transparency}{dnoTotalItem}{1249} = $rateOther
      if $model->{transparency};

    my $totalRevenue3;

    if ( $model->{legacy201} && !$model->{dcp189} ) {

        my $fixed3contribution = Arithmetic(
            name          => 'Demand fixed pot contribution p/day',
            defaultFormat => '0.00softnz',
            arithmetic    => '=100/A2*A1*(A6+A7+A88)',
            arguments     => {
                A1  => $demandSoleUseAsset,
                A6  => $rateDirect,
                A7  => $rateIndirect,
                A88 => $rateRates,
                A2  => $daysInYear,
            }
        );

        my $capacity3 = Arithmetic(
            name          => 'Capacity pot contribution p/kVA/day',
            defaultFormat => '0.00softnz',
            arithmetic    => '=100/A3*((A1+A53)*(A6+A7+A8+A9)+A41*A42)',
            arguments     => {
                A3  => $daysInYear,
                A1  => $assetsCapacity,
                A53 => $assetsConsumption,
                A41 => $rateExit,
                A42 => ref $purpleUseRate eq 'ARRAY'
                ? $purpleUseRate->[0]
                : $purpleUseRate,
                A6 => $rateDirect,
                A7 => $rateIndirect,
                A8 => $rateRates,
                A9 => $rateOther,
            }
        );

        my $revenue3 = Arithmetic(
            name          => 'Pot contribution £/year',
            defaultFormat => '0softnz',
            arithmetic    => '=A9*0.01*(A1+A2*A3)',
            arguments     => {
                A1 => $fixed3contribution,
                A2 => $capacity3,
                A3 => $importCapacity,
                A9 => $daysInYear,
            }
        );

        $totalRevenue3 = GroupBy(
            name          => 'Pot £/year',
            defaultFormat => '0softnz',
            source        => $revenue3,
        );

    }

    else {

        $totalRevenue3 = Arithmetic(
            name          => 'Demand revenue target pot (£/year)',
            newBlock      => 1,
            defaultFormat => '0softnz',
            arithmetic    => '=A5*A6'
              . '+(A11+A12+A13)*(A21+A22+A23)'
              . '+(A14+A15)*A24'
              . ( $totalDcp189DiscountedAssets ? '-A31*A32' : '' ),
            arguments => {
                A5  => $rateExit,
                A6  => $edcmPurpleUse,
                A11 => $totalAssetsFixed,
                A12 => $totalAssetsCapacity,
                A13 => $totalAssetsConsumption,
                A14 => $totalAssetsCapacity,
                A15 => $totalAssetsConsumption,
                A21 => $rateDirect,
                A22 => $rateRates,
                A23 => $rateIndirect,
                A24 => $rateOther,
                $totalDcp189DiscountedAssets
                ? (
                    A31 => $rateDirect,
                    A32 => $totalDcp189DiscountedAssets,
                  )
                : (),
            },
        );

        $totalRevenue3 =
          $model->{takenForAnIdiot}->demandRevenuePot($totalRevenue3)
          if $model->{takenForAnIdiot};

    }

    $model->{transparency}{dnoTotalItem}{1201} = $totalRevenue3
      if $model->{transparency};

    push @{ $model->{calc3Tables} }, $totalRevenue3;

    $totalRevenue3, $rateOther;

}

sub chargingRates {

    my (
        $model,          $chargeDirect,    $chargeRates,
        $chargeIndirect, $totalEdcmAssets, $cdcmEhvAssets,
        $cdcmHvLvShared, $cdcmHvLvService, $ehvIntensity,
    ) = @_;

    my $rateDirect = Arithmetic(
        name          => 'Direct cost charging rate',
        groupName     => 'Expenditure charging rates',
        arithmetic    => '=A1/(A2+A3+(A4+A5)/A6)',
        defaultFormat => '%soft',
        arguments     => {
            A1 => $chargeDirect,
            A2 => $totalEdcmAssets,
            A3 => $cdcmEhvAssets,
            A4 => $cdcmHvLvShared,
            A5 => $cdcmHvLvService,
            A6 => $ehvIntensity,
        },
        location => 'Charging rates',
    );
    $model->{transparency}{dnoTotalItem}{1245} = $rateDirect
      if $model->{transparency};

    my $rateRates = Arithmetic(
        name          => 'Network rates charging rate',
        groupName     => 'Expenditure charging rates',
        arithmetic    => '=A1/(A2+A3+A4+A5)',
        defaultFormat => '%soft',
        arguments     => {
            A1 => $chargeRates,
            A2 => $totalEdcmAssets,
            A3 => $cdcmEhvAssets,
            A4 => $cdcmHvLvShared,
            A5 => $cdcmHvLvService,
        },
        location => 'Charging rates',
    );
    $model->{transparency}{dnoTotalItem}{1246} = $rateRates
      if $model->{transparency};

    my $rateIndirect = Arithmetic(
        name          => 'Indirect cost charging rate',
        groupName     => 'Expenditure charging rates',
        arithmetic    => '=A1/(A20+A3+(A4+A5)/A6)',
        defaultFormat => '%soft',
        arguments     => {
            A1  => $chargeIndirect,
            A3  => $cdcmEhvAssets,
            A4  => $cdcmHvLvShared,
            A6  => $ehvIntensity,
            A20 => $totalEdcmAssets,
            A5  => $cdcmHvLvService,
        },
        location => 'Charging rates',
    );
    $model->{transparency}{dnoTotalItem}{1250} = $rateIndirect
      if $model->{transparency};

    $rateDirect, $rateRates, $rateIndirect;

}

sub fixedCharges {

    my (
        $model,                      $rateDirect,
        $rateRates,                  $daysInYear,
        $demandSoleUseAsset,         $dcp189Input,
        $demandSoleUseAssetUnscaled, $importEligible,
        $generationSoleUseAsset,     $generationSoleUseAssetUnscaled,
    ) = @_;

    my $fixedDcharge =
      !$model->{dcp189} ? Arithmetic(
        name          => 'Demand fixed charge p/day (scaled for part year)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=100/A2*A1*(A6+A88)',
        arguments     => {
            A1  => $demandSoleUseAsset,
            A6  => $rateDirect,
            A88 => $rateRates,
            A2  => $daysInYear,
        }
      )
      : $model->{dcp189} =~ /proportion/i ? Arithmetic(
        name          => 'Demand fixed charge p/day (scaled for part year)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=100/A2*A1*((1-A4)*A6+A88)',
        arguments     => {
            A1  => $demandSoleUseAsset,
            A4  => $dcp189Input,
            A6  => $rateDirect,
            A88 => $rateRates,
            A2  => $daysInYear,
        }
      )
      : Arithmetic(
        name          => 'Demand fixed charge p/day (scaled for part year)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=100/A2*A1*(IF(A4="Y",0,A6)+A88)',
        arguments     => {
            A1  => $demandSoleUseAsset,
            A4  => $dcp189Input,
            A6  => $rateDirect,
            A88 => $rateRates,
            A2  => $daysInYear,
        }
      );

    my $fixedDchargeTrue =
      !$model->{dcp189} ? Arithmetic(
        name          => 'Demand fixed charge p/day',
        groupName     => 'Fixed charges',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(A3,(100/A2*A1*(A6+A88)),0)',
        arguments     => {
            A1  => $demandSoleUseAssetUnscaled,
            A6  => $rateDirect,
            A88 => $rateRates,
            A2  => $daysInYear,
            A3  => $importEligible,
        }
      )
      : $model->{dcp189} =~ /proportion/i ? Arithmetic(
        name          => 'Demand fixed charge p/day',
        groupName     => 'Fixed charges',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(A3,(100/A2*A1*((1-A4)*A6+A88)),0)',
        arguments     => {
            A1  => $demandSoleUseAssetUnscaled,
            A4  => $dcp189Input,
            A6  => $rateDirect,
            A88 => $rateRates,
            A2  => $daysInYear,
            A3  => $importEligible,
        }
      )
      : Arithmetic(
        name          => 'Demand fixed charge p/day',
        groupName     => 'Fixed charges',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(A3,(100/A2*A1*(IF(A4="Y",0,A6)+A88)),0)',
        arguments     => {
            A1  => $demandSoleUseAssetUnscaled,
            A4  => $dcp189Input,
            A6  => $rateDirect,
            A88 => $rateRates,
            A2  => $daysInYear,
            A3  => $importEligible,
        }
      );

    my $fixedGcharge = Arithmetic(
        name          => 'Generation fixed charge p/day (scaled for part year)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=100/A2*A1*(A6+A88)',
        arguments     => {
            A1  => $generationSoleUseAsset,
            A6  => $rateDirect,
            A88 => $rateRates,
            A2  => $daysInYear,
        }
    );

    my $fixedGchargeUnround = Arithmetic(
        name          => 'Export fixed charge (unrounded) p/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=100/A2*A1*(A6+A88)',
        arguments     => {
            A1  => $generationSoleUseAssetUnscaled,
            A6  => $rateDirect,
            A88 => $rateRates,
            A2  => $daysInYear,
        }
    );

    my $fixedGchargeTrue = Arithmetic(
        name          => 'Export fixed charge p/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=ROUND(A1,2)',
        arguments     => { A1 => $fixedGchargeUnround, }
    );

    $fixedDcharge, $fixedDchargeTrue, $fixedGcharge, $fixedGchargeUnround,
      $fixedGchargeTrue;

}

sub exitChargingRate {

    my ( $model, $cdcmUse, $purpleUseRate, $importCapacity, $chargeExit, ) = @_;

    my $cdcmPurpleUse = Stack(
        cols => Labelset( list => [ $cdcmUse->{cols}{list}[0] ] ),
        name    => 'Total CDCM peak time consumption (kW)',
        sources => [$cdcmUse]
    );
    $model->{transparency}{dnoTotalItem}{1237} = $cdcmPurpleUse
      if $model->{transparency};

    my $edcmPurpleUse =
      $model->{transparencyMasterFlag}
      ? Arithmetic(
        name          => 'Total EDCM peak time consumption (kW)',
        defaultFormat => '0softnz',
        arithmetic    => '=IF(A123,0,A1)+SUMPRODUCT(A21_A22,A51_A52,A53_A54)',
        arguments     => {
            A123    => $model->{transparencyMasterFlag},
            A1      => $model->{transparency}{baselineItem}{119101},
            A21_A22 => $model->{transparency},
            A51_A52 => ref $purpleUseRate eq 'ARRAY'
            ? $purpleUseRate->[0]
            : $purpleUseRate,
            A53_A54 => $importCapacity,
        }
      )
      : SumProduct(
        name   => 'Total EDCM peak time consumption (kW)',
        vector => ref $purpleUseRate eq 'ARRAY'
        ? $purpleUseRate->[0]
        : $purpleUseRate,
        matrix        => $importCapacity,
        defaultFormat => '0softnz'
      );

    $model->{transparency}{dnoTotalItem}{119101} = $edcmPurpleUse
      if $model->{transparency};

    my $overallPurpleUse = Arithmetic(
        name          => 'Estimated total peak-time consumption (kW)',
        defaultFormat => '0softnz',
        arithmetic    => '=A1+A2',
        arguments     => { A1 => $cdcmPurpleUse, A2 => $edcmPurpleUse }
    );
    $model->{transparency}{dnoTotalItem}{1238} = $overallPurpleUse
      if $model->{transparency};

    my $rateExit = Arithmetic(
        name       => 'Transmission exit charging rate (£/kW/year)',
        arithmetic => '=A1/A2',
        arguments  => { A1 => $chargeExit, A2 => $overallPurpleUse },
        location   => 'Charging rates',
    );
    $model->{transparency}{dnoTotalItem}{1239} = $rateExit
      if $model->{transparency};

    $rateExit, $edcmPurpleUse;

}

1;
