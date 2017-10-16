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

sub fixedChargeAdj {
    my ( $self, $rateDirect, $rateRates, $demandSoleUseAsset, $chargeDirect,
        $chargeRates, )
      = @_;
    my $tariffIndex = Dataset(
        name          => 'Tariff index (avoiding DCP 189 sites)',
        defaultFormat => '0hard',
        data          => [1],
    );
    my $fixedCharge = Dataset(
        name          => 'Demand fixed charge (£/year)',
        defaultFormat => '0hard',
        data          => [1e5],
    );
    Columnset(
        name     => 'Demand fixed charge for one tariff',
        columns  => [ $tariffIndex, $fixedCharge, ],
        dataset  => $self->{model}{dataset},
        appendTo => $self->{model}{inputTables},
        number   => 11953,
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
    $rateDirectAdjusted, $rateRatesAdjusted;
}

sub demandRevenuePot {
    my ( $self, $totalRevenue3 ) = @_;
    return $totalRevenue3;
    $self->{demandRevenuePot} = Dataset(
        name          => 'Demand revenue pot (£/year)',
        defaultFormat => '0hard',
        data          => [1e7],
        dataset       => $self->{model}{dataset},
        appendTo      => $self->{model}{inputTables},
        number        => 11951,
    );
    push @{ $self->{feedbackTables} },
      Arithmetic(
        name          => 'Adjustment to demand revenue pot (£/year)',
        defaultFormat => '0softpm',
        arithmetic    => '=A1-A2',
        arguments     => {
            A1 => $self->{demandRevenuePot},
            A2 => $totalRevenue3,
        },
      );
    $self->{demandRevenuePot};
}

sub adjustDnoTotals {
    my ( $self, $hashedArrays, ) = @_;
    if ( $self->{demandRevenuePot} ) {
        warn 'Not implemented';
    }
    else {
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
}

sub feedbackTables {
    my ($self) = @_;
    return unless $self->{feedbackTables};
    @{ $self->{feedbackTables} };
}

1;
