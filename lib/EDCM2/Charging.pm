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
