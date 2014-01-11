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

sub allocation {

    my (
        $model,             $afterAllocation,     $allocLevelset,
        $allocationRules,   $capitalised,         $expenditure,
        $incentive,         $netCapexPercentages, $revenue,
        $totalDepreciation, $totalOperating,      $totalReturn,
        $units,
    ) = @_;

    push @{ $model->{calcTables} },
      my $toDeduct = Arithmetic(
        name => 'To be deducted from revenue and treated as "upstream" cost',
        defaultFormat => '0soft',
        arithmetic    => '=SUMIF(IV1_IV2,"Deduct from revenue",IV3_IV4)',
        arguments => { IV1_IV2 => $allocationRules, IV3_IV4 => $expenditure }
      );

    push @{ $model->{calcTables} },
      my $toDeductDown = Arithmetic(
        name => 'To be deducted from revenue and treated as "downstream" cost',
        arithmetic =>
'=SUMIF(IV1_IV2,"Deduct from revenue and treat as downstream",IV3_IV4)',
        arguments => { IV1_IV2 => $allocationRules, IV3_IV4 => $expenditure }
      ) if $model->{deductDownstream};    # not implemented (I think)

    my $expensed = Arithmetic(
        name => 'Complete allocation, adjusted for regulatory capitalisation',
        defaultFormat => '0softnz',
        arithmetic    => '=IV1*(1-IV2)',
        arguments     => { IV1 => $afterAllocation, IV2 => $capitalised, }
    );

    my $expensedTotals = GroupBy(
        name          => 'Total expensed for each level',
        defaultFormat => '0softnz',
        cols          => $allocLevelset,
        source        => $expensed
    );

    my $expensedPercentages = Arithmetic(
        name          => 'Expensed proportions',
        defaultFormat => '%soft',
        arithmetic    => '=IV1/SUM(IV2_IV3)',
        arguments     => { IV1 => $expensedTotals, IV2_IV3 => $expensedTotals }
    );

    my $ppusingle = Arithmetic(
        name => 'p/kWh split (single-step calculation)',
        arithmetic =>
'=100*((IV31+IV41)*IV2+IV51*IV1)/(IV32+IV42+IV52)*(IV9-IV81-IV82)/IV7',
        arguments => {
            IV1  => $expensedPercentages,
            IV2  => $netCapexPercentages,
            IV31 => $totalReturn,
            IV32 => $totalReturn,
            IV41 => $totalDepreciation,
            IV42 => $totalDepreciation,
            IV51 => $totalOperating,
            IV52 => $totalOperating,
            IV7  => $units,
            IV81 => $incentive,
            IV82 => $toDeduct,
            IV9  => $revenue,
        }
    );

    my $propOp = Arithmetic(
        name       => 'Proportion of price control revenue attributed to opex',
        arithmetic => '=IV1/(IV32+IV42+IV52)',
        arguments  => {
            IV32 => $totalReturn,
            IV42 => $totalDepreciation,
            IV1  => $totalOperating,
            IV52 => $totalOperating,
        },
        defaultFormat => '%soft',
    );

    my $revenueToBeAllocated = Arithmetic(
        name => 'Revenue to be allocated between network levels (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=IV1-IV81-IV82',
        arguments     => {
            IV81 => $incentive,
            IV82 => $toDeduct,
            IV1  => $revenue,
        },
    );

    push @{ $model->{calcTables} },
      my $ppu = Arithmetic(
        name       => 'p/kWh split',
        arithmetic => '=((1-IV52)*IV2+IV51*IV1)*IV6/IV7*100',
        arguments  => {
            IV1  => $expensedPercentages,
            IV2  => $netCapexPercentages,
            IV7  => $units,
            IV51 => $propOp,
            IV52 => $propOp,
            IV6  => $revenueToBeAllocated,
        }
      );

    my $ppuNotSplit = Arithmetic(
        name => 'p/kWh not split',
        cols => Labelset(
            list => [ $allocLevelset->{list}[ $#{ $allocLevelset->{list} } ] ]
        ),
        arithmetic => '=100*(IV81+IV82)/IV7',
        arguments  => {
            IV7  => $units,
            IV81 => $incentive,
            IV82 => $toDeduct,
            IV9  => $revenue,
        }
    );

    my $alloc = Arithmetic(
        name          => 'Allocated proportion',
        defaultFormat => '%soft',
        arithmetic    => '=IV1/(SUM(IV2_IV3)+IV4)',
        arguments     => { IV1 => $ppu, IV2_IV3 => $ppu, IV4 => $ppuNotSplit, }
    );

    $alloc;

}

1;
