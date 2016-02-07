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

    my $toDeduct =
      $model->{objects}{toDeduct}{ 0 + $allocationRules }{ 0 + $expenditure }
      ||= Arithmetic(
        name => 'To be deducted from revenue and treated as "upstream" cost',
        defaultFormat => '0soft',
        arithmetic    => '=SUMIF(A1_A2,"Deduct from revenue",A3_A4)',
        arguments     => { A1_A2 => $allocationRules, A3_A4 => $expenditure }
      );

    my $expensed = Arithmetic(
        name => 'Complete allocation, adjusted for regulatory capitalisation',
        defaultFormat => '0softnz',
        arithmetic    => '=A1*(1-A2)',
        arguments     => { A1 => $afterAllocation, A2 => $capitalised, }
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
        arithmetic    => '=A1/SUM(A2_A3)',
        arguments     => { A1 => $expensedTotals, A2_A3 => $expensedTotals }
    );

    my $ppuSingleNotUsed = Arithmetic(
        name => 'p/kWh split (single-step calculation)',
        arithmetic =>
          '=100*((A31+A41)*A2+A51*A1)/(A32+A42+A52)*(A9-A81-A82)/A7',
        arguments => {
            A1  => $expensedPercentages,
            A2  => $netCapexPercentages,
            A31 => $totalReturn,
            A32 => $totalReturn,
            A41 => $totalDepreciation,
            A42 => $totalDepreciation,
            A51 => $totalOperating,
            A52 => $totalOperating,
            A7  => $units,
            A81 => $incentive,
            A82 => $toDeduct,
            A9  => $revenue,
        }
    );

    my $propOp =
      $model->{objects}{propOp}{ 0 + $totalReturn }{ 0 + $totalDepreciation }
      { 0 + $totalOperating } ||= Arithmetic(
        name       => 'Proportion of price control revenue attributed to opex',
        arithmetic => '=A1/(A32+A42+A52)',
        arguments  => {
            A32 => $totalReturn,
            A42 => $totalDepreciation,
            A1  => $totalOperating,
            A52 => $totalOperating,
        },
        defaultFormat => '%soft',
      );

    my $revenueToBeAllocated =
      $model->{objects}{revenueToBeAllocated}{ 0 + $incentive }
      { 0 + $toDeduct }{ 0 + $revenue } ||= Arithmetic(
        name => 'Revenue to be allocated between network levels (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=A1-A81-A82',
        arguments     => {
            A81 => $incentive,
            A82 => $toDeduct,
            A1  => $revenue,
        },
      );

    my $ppu = Arithmetic(
        name       => 'p/kWh split',
        arithmetic => '=((1-A52)*A2+A51*A1)*A6/A7*100',
        arguments  => {
            A1  => $expensedPercentages,
            A2  => $netCapexPercentages,
            A7  => $units,
            A51 => $propOp,
            A52 => $propOp,
            A6  => $revenueToBeAllocated,
        }
    );

    if ( $model->{dcp117} && $model->{dcp117} =~ /2014/ ) {
        my $incomeForConnectionsIndirect =
          $model->{objects}{incomeForConnectionsIndirect} ||= Dataset(
            name  => 'Income for connections indirects (£)',
            lines => 'In a legacy Method M workbook, this item is '
              . 'on sheet Calc-Allocation, possibly cell G70',
            defaultFormat => '0hard',
            data          => [''],
            number        => 1328,
            dataset       => $model->{dataset},
            appendTo      => $model->{objects}{inputTables},
          );
        if ( $model->{dcp118} && $model->{dcp117dcp118interaction} ) {
            $model->{objects}{ehvRevenueInput} ||= Dataset(
                name => 'EHV connected customers component of revenue (£/year)',
                lines => 'In a legacy Method M workbook, this item is '
                  . 'on sheet Calc-Allocation, possibly cell G68.',
                defaultFormat => '0hard',
                data          => [''],
                number        => 1316,
                dataset       => $model->{dataset},
                appendTo      => $model->{objects}{inputTables},
            );
            $incomeForConnectionsIndirect =
              $model->{objects}{incomeForConnectionsIndirectCooked} ||=
              Arithmetic(
                name => 'Income for connections indirects '
                  . 'scaled by total/non-EHV revenue ratio (£)',
                defaultFormat => '0soft',
                arithmetic    => '=A1/(1-A2/A3)',
                arguments     => {
                    A1 => $incomeForConnectionsIndirect,
                    A2 => $model->{objects}{ehvRevenueInput},
                    A3 => $model->{objects}{oneYearDpcr}{columns}[0],
                },
              );
        }
        my $incomeForConnectionsIndirectPercentages = $expensedPercentages;
        if ( $model->{dcp117weirdness} ) {
            my $lvLevelset =
              Labelset( list => [ @{ $ppu->{cols}{list} }[ 0, 1 ] ] );
            my $ppuLv = Stack(
                name => 'Pre-DCP 117 p/kWh split for LV mains and LV services',
                cols => $lvLevelset,
                sources => [$ppu],
            );
            my $expensedLv = Stack(
                name    => 'Expensed proportions for LV mains and LV services',
                cols    => $lvLevelset,
                sources => [$expensedPercentages],
            );
            $incomeForConnectionsIndirectPercentages = Stack(
                name          => 'Allocation for connections indirects income',
                defaultFormat => '%copy',
                cols          => $expensedPercentages->{cols},
                sources       => [
                    Arithmetic(
                        name =>
                          'LV allocations for connections indirects income',
                        defaultFormat => '%soft',
                        arithmetic    => '=A1/SUM(A2_A3)*SUM(A4_A5)',
                        arguments     => {
                            A1    => $ppuLv,
                            A2_A3 => $ppuLv,
                            A4_A5 => $expensedLv,
                        },
                    ),
                    $expensedPercentages,
                ],
            );
        }
        $ppu = Arithmetic(
            name       => 'p/kWh split (DCP 117 modified)',
            arithmetic => '=(((1-A52)*A2+A51*A1)*A6+A8*A18)/A7*100',
            arguments  => {
                A1  => $expensedPercentages,
                A2  => $netCapexPercentages,
                A7  => $units,
                A51 => $propOp,
                A52 => $propOp,
                A6  => $revenueToBeAllocated,
                A1  => $expensedPercentages,
                A18 => $incomeForConnectionsIndirectPercentages,
                A8  => $incomeForConnectionsIndirect,
            }
        );
    }

    my $ppuNotSplit =
      $model->{objects}{ppuNotSplit}{ 0 + $units }{ 0 + $incentive }
      { 0 + $toDeduct }{ 0 + $revenue } ||= Arithmetic(
        name => 'p/kWh not split',
        cols => Labelset(
            list => [ $allocLevelset->{list}[ $#{ $allocLevelset->{list} } ] ]
        ),
        arithmetic => '=100*(A81+A82)/A7',
        arguments  => {
            A7  => $units,
            A81 => $incentive,
            A82 => $toDeduct,
            A9  => $revenue,
        }
      );

    my $alloc = Arithmetic(
        name          => 'Allocated proportion',
        defaultFormat => '%soft',
        arithmetic    => '=A1/(SUM(A2_A3)+A4)',
        arguments     => { A1 => $ppu, A2_A3 => $ppu, A4 => $ppuNotSplit, }
    );

    $alloc;

}

1;
