package EDCM2;

# Copyright 2009-2012 Energy Networks Association Limited and others.
# Copyright 2015-2017 Franck Latrémolière, Reckon LLP and others.
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

sub exportCharges {

    my (
        $model,                    $gCharge,
        $daysInYear,               $exportEligible,
        $genCredit,                $genCreditCapacity,
        $exportCapacityChargeable, $activeUnits,
        $fixedGcharge,
    ) = @_;

    my $exportCapacityCharge = Arithmetic(
        name          => 'Export capacity charge (unrounded) p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(A1,A4,0)',
        arguments     => {
            A1 => $exportEligible,
            A4 => $gCharge,
        },
    );

    my $genCreditRound = Arithmetic(
        name          => "Export $model->{timebandName} unit rate (p/kWh)",
        defaultFormat => '0.000softnz',
        arithmetic    => '=ROUND(A1,3)',
        arguments     => { A1 => $genCredit }
    );

    my $exportCapacityChargeRound = Arithmetic(
        name          => 'Export capacity charge (p/kVA/day)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=ROUND(A1,2)',
        arguments     => { A1 => $exportCapacityCharge }
    );

    my $netexportCapacityChargeUnRound = Arithmetic(
        name =>
          'Net export capacity charge (or credit) (unrounded) (p/kVA/day)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(A21,(A1+A2),0)',
        arguments     => {
            A1  => $exportCapacityCharge,
            A2  => $genCreditCapacity,
            A21 => $exportEligible
        }
    );

    my $netexportCapacityChargeRound = Arithmetic(
        name          => 'Export capacity rate (p/kVA/day)',
        groupName     => 'Export capacity rate',
        defaultFormat => '0.00softnz',
        arithmetic    => '=ROUND(A1,2)',
        arguments     => { A1 => $netexportCapacityChargeUnRound, }
    );

    my $generationRevenue =
      $model->{transparencyMasterFlag}
      ? Arithmetic(
        name          => 'Net forecast EDCM generation revenue (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=IF(A123,0,A1)'
          . '+SUMPRODUCT(A21_A22,A51_A52,A53_A54)/100'
          . '+SUMPRODUCT(A31_A32,A71_A72,A73_A74)*A75/100+SUMPRODUCT(A41_A42,A83_A84)*A85/100',
        arguments => {
            A123    => $model->{transparencyMasterFlag},
            A1      => $model->{transparency}{baselineItem}{119105},
            A21_A22 => $model->{transparency},
            A31_A32 => $model->{transparency},
            A41_A42 => $model->{transparency},
            A51_A52 => $genCreditRound,
            A53_A54 => $activeUnits,
            A71_A72 => $netexportCapacityChargeRound,
            A73_A74 => $exportCapacityChargeable,
            A75     => $daysInYear,
            A83_A84 => $fixedGcharge,
            A85     => $daysInYear,
        }
      )
      : Arithmetic(
        name          => 'Net forecast EDCM generation revenue (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=SUMPRODUCT(A51_A52,A53_A54)/100+'
          . 'SUMPRODUCT(A71_A72,A73_A74)*A75/100+SUM(A83_A84)*A85/100',
        arguments => {
            A51_A52 => $genCreditRound,
            A53_A54 => $activeUnits,
            A71_A72 => $netexportCapacityChargeRound,
            A73_A74 => $exportCapacityChargeable,
            A75     => $daysInYear,
            A83_A84 => $fixedGcharge,
            A85     => $daysInYear,
        }
      );

    $model->{transparency}{dnoTotalItem}{119105} = $generationRevenue
      if $model->{transparency};

    $exportCapacityCharge, $genCreditRound, $exportCapacityChargeRound,
      $netexportCapacityChargeRound, $generationRevenue;

}

sub gCharge {

    my (
        $model,                            $genPot20p,
        $genPotGP,                         $genPotGL,
        $genPotCdcmCap20052010,            $genPotCdcmCapPost2010,
        $exportCapacityChargeable,         $exportCapacityChargeable20052010,
        $exportCapacityChargeablePost2010, $daysInYear,
    ) = @_;

    if ( $model->{transparencyMasterFlag} ) {
        ${ $_->[0] } = Arithmetic(
            name          => ${ $_->[0] }->objectShortName . ' (total)',
            defaultFormat => '0softnz',
            arithmetic    => '=IF(A123,0,A1)+SUMPRODUCT(A2_A3,A4_A5)',
            arguments     => {
                A123  => $model->{transparencyMasterFlag},
                A1    => $model->{transparency}{baselineItem}{ $_->[1] },
                A2_A3 => ${ $_->[0] },
                A4_A5 => $model->{transparency},
            }
          )
          foreach [ \$exportCapacityChargeable, 119201 ],
          [ \$exportCapacityChargeable20052010, 119202 ],
          [ \$exportCapacityChargeablePost2010, 119203 ];
    }
    else {
        $_ = GroupBy(
            name          => $_->objectShortName . ' (total)',
            defaultFormat => '0softnz',
            source        => $_
          )
          foreach $exportCapacityChargeable, $exportCapacityChargeable20052010,
          $exportCapacityChargeablePost2010;
    }

    if ( $model->{transparency} ) {
        $model->{transparency}{dnoTotalItem}{ $_->[1] } = $_->[0]
          foreach [ $exportCapacityChargeable, 119201 ],
          [ $exportCapacityChargeable20052010, 119202 ],
          [ $exportCapacityChargeablePost2010, 119203 ];
    }

    my $exportCapacityCharge = Arithmetic(
        name          => 'Export capacity charge p/kVA/day',
        groupName     => 'Generic export capacity charge',
        defaultFormat => '0.00softnz',
        arithmetic    => '=('
          . 'A1*(1-IF(A211,A21/A23,0))'
          . '+IF(A31,A3*A22/(A24+A52)/A231,0)'
          . '+IF(A41,A4*A25/(A26+A51)/A232,0)'
          . ')*100/A9',
        arguments => {
            A1   => $genPot20p,
            A21  => $exportCapacityChargeable20052010,
            A211 => $exportCapacityChargeable20052010,
            A22  => $exportCapacityChargeablePost2010,
            A23  => $exportCapacityChargeable,
            A231 => $exportCapacityChargeable,
            A232 => $exportCapacityChargeable,
            A24  => $exportCapacityChargeablePost2010,
            A25  => $exportCapacityChargeable20052010,
            A26  => $exportCapacityChargeable20052010,
            A3   => $genPotGP,
            A31  => $genPotGP,
            A4   => $genPotGL,
            A41  => $genPotGL,
            A51  => $genPotCdcmCap20052010,
            A52  => $genPotCdcmCapPost2010,
            A9   => $daysInYear,
        }
    );
    $model->{transparency}{dnoTotalItem}{1243} = $exportCapacityCharge
      if $model->{transparency};

    $exportCapacityCharge;

}

1;
