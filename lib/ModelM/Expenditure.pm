package ModelM;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others.
Copyright 2012-2013 Franck Latrémolière, Reckon LLP and others.

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

sub expenditureAlloc {

    my (
        $model,           $allocLevelset, $allocationRules, $capitalised,
        $directIndicator, $expenditure,   $meavPercentages,
    ) = @_;

    my ( $networkLengthPercentages, ) =
      $model->networkLengthPercentages($allocLevelset);

    my ( $preAllocated, ) =
      $model->allocated( $allocLevelset, $allocationRules->{rows} );

    ($preAllocated) = $model->ajust117( $meavPercentages, $preAllocated, )
      if $model->{dcp117};

    push @{ $model->{calcTables} }, $allocationRules,
      my $allocatedTotal = GroupBy(
        name          => 'Amounts already allocated',
        source        => $preAllocated,
        rows          => $preAllocated->{rows},
        defaultFormat => '0softnz'
      );

    my $lvOnly = Constant(
        name          => 'LV only',
        cols          => $preAllocated->{cols},
        data          => [qw(1 0 0 0)],
        defaultFormat => '0connz'
    );

    my $allAllocationPercentages = Arithmetic(
        name          => 'All allocation percentages',
        defaultFormat => '%softnz',
        rows          => $preAllocated->{rows},
        cols          => $preAllocated->{cols},
        arithmetic    => '=IF(IV44="60%MEAV",0.4*IV71+IV51,'
          . 'IF(IV45="MEAV",IV52,'
          . 'IF(IV46="EHV only",IV6,'
          . 'IF(IV47="LV only",IV72,'
          . 'IF(IV48="Network length",IV8,'
          . ( $model->{dcp097} ? 'IF(IV49="Customer numbers",IV9,0)' : '0' )
          . ')))))',
        arguments => {
            IV44 => $allocationRules,
            IV45 => $allocationRules,
            IV46 => $allocationRules,
            IV47 => $allocationRules,
            IV48 => $allocationRules,
            IV51 => $meavPercentages,
            IV52 => $meavPercentages,
            IV6  => Constant(
                name          => 'EHV only',
                cols          => $preAllocated->{cols},
                data          => [qw(0 0 0 1)],
                defaultFormat => '0connz'
            ),
            IV71 => $lvOnly,
            IV72 => $lvOnly,
            IV8  => $networkLengthPercentages,
            $model->{dcp097}
            ? (
                IV49 => $allocationRules,
                IV9  => $model->customerNumbersPercentages($allocLevelset)
              )
            : (),
        }
    );

    my $afterAllocation = Arithmetic(
        name          => 'Complete allocation',
        defaultFormat => '0softnz',
        arithmetic    => '=IF(IV44="Kill",0,IV1+(IV2-IV3)*IV5)',
        arguments     => {
            IV1  => $preAllocated,
            IV2  => $expenditure,
            IV3  => $allocatedTotal,
            IV44 => $allocationRules,
            IV5  => $allAllocationPercentages,
        }
    );

    my $omittingNegatives = $model->{allowNeg} ? $afterAllocation : Arithmetic(
        name          => 'Complete allocation, zeroing out negative numbers',
        defaultFormat => '0softnz',
        arithmetic    => '=MAX(0,IV1)',
        arguments     => { IV1 => $afterAllocation }
    );

    my $totalDirect = SumProduct(
        name          => 'Direct costs',
        defaultFormat => '0softnz',
        matrix        => $omittingNegatives,
        vector        => $directIndicator
    );

    my $total = GroupBy(
        name          => 'Total costs',
        defaultFormat => '0softnz',
        source        => $omittingNegatives,
        cols          => $omittingNegatives->{cols},
    );

    my $direct = Arithmetic(
        name          => 'Direct cost proportion',
        defaultFormat => '%soft',
        arithmetic    => '=IV1/IV2',
        arguments     => { IV1 => $totalDirect, IV2 => $total }
    );

    $afterAllocation, $direct;

}

1;
