package EDCM2::IdiotMitigation;

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

sub demandRevenuePotAdj {
    my ( $self, $totalRevenue3, $rateDirect, $rateRates, $rateIndirect, ) = @_;
    my $demandRevenuePot = Dataset(
        name          => 'Demand revenue pot (£/year)',
        defaultFormat => '0hard',
        data          => [1e7],
        dataset       => $self->{model}{dataset},
        appendTo      => $self->{model}{inputTables},
        number        => 11951,
    );
    $self->{demandAssetAdjustment} = Arithmetic(
        name          => 'Adjustment to EDCM demand assets (£)',
        defaultFormat => '0softpm',
        arithmetic    => '=(A1-A2)/(A21+A22+A23)',
        arguments     => {
            A1  => $demandRevenuePot,
            A2  => $totalRevenue3,
            A21 => $rateDirect,
            A22 => $rateRates,
            A23 => $rateIndirect,
        },
    );
    $demandRevenuePot;
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
    );
    $self->{exportCapacityChargeable20052010adj} = Arithmetic(
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
    $gChargeHard;
}

sub exitChargeAdj {
    my ( $self, ) = @_;
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
        name     => 'Demand transmission exit charge for a site',
        columns  => [ $tariffIndex, $charge, ],
        dataset  => $self->{model}{dataset},
        appendTo => $self->{model}{inputTables},
        number   => 11960,
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
        name       => 'Thing',
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
        arithmetic    => '=IF(ISNUMBER(A11),1/(1/A2+A1/A3),A21)',
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
        arithmetic    => '=IF(ISNUMBER(A11),1/(1/A2+A1/A3),A21)',
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
        arithmetic    => '=IF(ISNUMBER(A11),1/(1/A2+A1/A3),A21)',
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
    my ( $self, ) = @_;
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
    );
}

sub fixedAdderAdj {
    my ( $self, ) = @_;
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
    );
}

sub assetAdderAdj {
    my ( $self, $demandScaling, $totalSlopeCapacity,
        $slopeCapacity, $shortfall, $direct, $rates )
      = @_;
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
        name     => 'Demand asset adder charge for a non-pathological site',
        columns  => [ $tariffIndex, $charge, ],
        dataset  => $self->{model}{dataset},
        appendTo => $self->{model}{inputTables},
        number   => 11964,
    );
    $self->{cookedSharedAssetAdjustment} = Arithmetic(
        name          => 'Adjustment to shared assets subject to adder (£)',
        defaultFormat => '0softpm',
        arithmetic    => '=(A1-A12-A13)/A2*INDEX(A3_A4,A5)-A6',
        arguments     => {
            A1    => $shortfall,
            A12   => $direct,
            A13   => $rates,
            A2    => $charge,
            A3_A4 => $slopeCapacity,
            A5    => $tariffIndex,
            A6    => $totalSlopeCapacity,
        },
    );
    Arithmetic(
        name          => 'Adjusted asset adder rate',
        defaultFormat => '%soft',
        arithmetic    => '=IF(ISNUMBER(A11),1/(1/A2+A1/A3),A21)',
        arguments     => {
            A1  => $self->{cookedSharedAssetAdjustment},
            A11 => $self->{cookedSharedAssetAdjustment},
            A2  => $demandScaling,
            A3  => $shortfall,
            A21 => $demandScaling,
        },
    );
}

sub directChargeAdj {
    my ( $self, ) = @_;
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
        name     => 'Demand direct cost charge for a non-pathological site',
        columns  => [ $tariffIndex, $charge, ],
        dataset  => $self->{model}{dataset},
        appendTo => $self->{model}{inputTables},
        number   => 11965,
    );
}

sub calcTables {
    my ($self) = @_;
    [
        grep { $_; } @{$self}{
            qw(
              edcmPurpleUseAdjustment
              edcmAssetAdjustment
              demandAssetAdjustment
              demandSharedAssetAdjustment
              exportCapacityChargeable20052010adj
              )
        }
    ];
}

sub adjustDnoTotals {

    my ( $self, $model, $hashedArrays, ) = @_;

    if ( $self->{exportCapacityChargeable20052010adj} ) {
        $hashedArrays->{1192}[0] = Arithmetic(
            name => 'Adjusted chargeable export capacity baseline (kVA)',
            defaultFormat => '0soft',
            arithmetic    => '=A1+IF(ISNUMBER(A2),A3,0)',
            arguments     => {
                A1 => $hashedArrays->{1192}[0]{sources}[0],
                A2 => $self->{exportCapacityChargeable20052010adj},
                A3 => $self->{exportCapacityChargeable20052010adj},
            },
        );
        $hashedArrays->{1192}[2] = Arithmetic(
            name =>
              'Adjusted non-exempt post-2010 export capacity baseline (kVA)',
            defaultFormat => '0soft',
            arithmetic    => '=A1+IF(ISNUMBER(A2),A3,0)',
            arguments     => {
                A1 => $hashedArrays->{1192}[2]{sources}[0],
                A2 => $self->{exportCapacityChargeable20052010adj},
                A3 => $self->{exportCapacityChargeable20052010adj},
            },
        );
    }

    if ( $self->{demandAssetAdjustment} && $self->{edcmAssetAdjustment} ) {
        if ( $self->{demandSharedAssetAdjustment} ) {
            $hashedArrays->{1193}[2] = Arithmetic(
                name          => 'Adjusted demand sole use assets baseline (£)',
                defaultFormat => '0soft',
                arithmetic    => '=A1+0.5*IF(ISNUMBER(A2),A3,0)',
                arguments     => {
                    A1 => $hashedArrays->{1193}[0]{sources}[0],
                    A2 => $self->{demandSharedAssetAdjustment},
                    A3 => $self->{demandSharedAssetAdjustment},
                },
            );
            $hashedArrays->{1193}[3] = Arithmetic(
                name          => 'Adjusted demand sole use assets baseline (£)',
                defaultFormat => '0soft',
                arithmetic    => '=A1+0.5*IF(ISNUMBER(A2),A3,0)',
                arguments     => {
                    A1 => $hashedArrays->{1193}[0]{sources}[0],
                    A2 => $self->{demandSharedAssetAdjustment},
                    A3 => $self->{demandSharedAssetAdjustment},
                },
            );
            $hashedArrays->{1193}[0] = Arithmetic(
                name          => 'Adjusted demand sole use assets baseline (£)',
                defaultFormat => '0soft',
                arithmetic => '=A1+IF(ISNUMBER(A2),A3,0)-IF(ISNUMBER(A4),A5,0)',
                arguments  => {
                    A1 => $hashedArrays->{1193}[0]{sources}[0],
                    A2 => $self->{demandAssetAdjustment},
                    A3 => $self->{demandAssetAdjustment},
                    A4 => $self->{demandSharedAssetAdjustment},
                    A5 => $self->{demandSharedAssetAdjustment},
                },
            );
        }
        else {
            $hashedArrays->{1193}[0] = Arithmetic(
                name          => 'Adjusted demand sole use assets baseline (£)',
                defaultFormat => '0soft',
                arithmetic    => '=A1+IF(ISNUMBER(A2),A3,0)',
                arguments     => {
                    A1 => $hashedArrays->{1193}[0]{sources}[0],
                    A2 => $self->{demandAssetAdjustment},
                    A3 => $self->{demandAssetAdjustment},
                },
            );
        }
        $hashedArrays->{1193}[1] = Arithmetic(
            name          => 'Adjusted generation sole use assets baseline (£)',
            defaultFormat => '0soft',
            arithmetic    => '=A1+IF(ISNUMBER(A2),A3,0)-IF(ISNUMBER(A4),A5,0)',
            arguments     => {
                A1 => $hashedArrays->{1193}[1]{sources}[0],
                A2 => $self->{edcmAssetAdjustment},
                A3 => $self->{edcmAssetAdjustment},
                A4 => $self->{demandAssetAdjustment},
                A5 => $self->{demandAssetAdjustment},
            },
        ) if $self->{edcmAssetAdjustment};
    }
    elsif ( $self->{edcmAssetAdjustment} )
    { # No pot data: put the whole EDCM asset adjustment on generation sole use assets
        $hashedArrays->{1193}[1] = Arithmetic(
            name          => 'Adjusted generation sole use assets baseline (£)',
            defaultFormat => '0soft',
            arithmetic    => '=A1+IF(ISNUMBER(A2),A3,0)',
            arguments     => {
                A1 => $hashedArrays->{1193}[1]{sources}[0],
                A2 => $self->{edcmAssetAdjustment},
                A3 => $self->{edcmAssetAdjustment},
            },
        );
    }

    if ( $self->{cookedSharedAssetAdjustment} ) {
        $hashedArrays->{1193}[4] = Arithmetic(
            name => 'Adjusted chargeable export capacity baseline (kVA)',
            defaultFormat => '0soft',
            arithmetic    => '=A1+IF(ISNUMBER(A2),A3,0)',
            arguments     => {
                A1 => $hashedArrays->{1193}[4]{sources}[0],
                A2 => $self->{cookedSharedAssetAdjustment},
                A3 => $self->{cookedSharedAssetAdjustment},
            },
        );
    }

}

1;
