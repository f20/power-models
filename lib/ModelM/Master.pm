package ModelM;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others.

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

=head Development note

This file is becoming bloated.  Should perhaps spin all the guts out:
* Before the DCP095 test.
* Back end with DCP095.
* Back end without DCP095.

=cut

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Shortcuts ':all';
use ModelM::Inputs;
use ModelM::Options;
use ModelM::Sheets;

sub new {
    my $class = shift;
    my $model = bless { inputTables => [], @_ }, $class;

    my $allocLevelset = Labelset(
        name => 'Allocation levels',
        list => [ split /\n/, <<END_OF_LIST ] );
LV
HV/LV
HV
EHV&132
END_OF_LIST
    my ( $lvSplit, $hvSplit ) = $model->splits;
    my ( $totalReturn, $totalDepreciation, $totalOperating, ) =
      $model->totalDpcr;
    my ( $revenue, $incentive, $pension, ) = $model->oneYearDpcr;
    my ( $units, ) = $model->units($allocLevelset);
    my ( $allocationRules, $capitalised, $directIndicator, ) =
      @{ $model->allocationRules };
    my $expenditureSet = $allocationRules->{rows};
    my ( $preAllocated, ) =
      $model->allocated( $allocLevelset, $expenditureSet );
    my ( $expenditure, )         = $model->expenditure($expenditureSet);
    my ( $meavPercentages, )     = $model->meavPercentages($allocLevelset);
    my ( $netCapexPercentages, ) = $model->netCapexPercentages($allocLevelset);
    my ( $networkLengthPercentages, ) =
      $model->networkLengthPercentages($allocLevelset);

    if ( $model->{DCP117} ) {

        if ( $model->{DCP117} =~ /half[ -]?baked/i ) {
            $preAllocated = Stack(
                name => 'Table 1330 allocated costs, after DCP117 adjustments',
                rows => $preAllocated->{rows},
                cols => $preAllocated->{cols},
                sources => [
                    Constant(
                        name => 'DCP117: remove negative number',
                        rows => Labelset(
                            list => [
'Load related new connections & reinforcement (net of contributions)'
                            ]
                        ),
                        cols => Labelset( list => ['LV'] ),
                        data => [ [0] ],
                        defaultFormat => '0connz',
                    ),
                    $preAllocated
                ],
            );
        }
        else {
            my $dcp117negative = Stack(
                name => 'DCP117: negative number being removed',
                rows => Labelset(
                    list => [
                            'Load related new connections & '
                          . 'reinforcement (net of contributions)'
                    ]
                ),
                cols          => Labelset( list => ['LV'] ),
                sources       => [$preAllocated],
                defaultFormat => '0copynz',
            );
            my $dcp117 = new SpreadsheetModel::Custom(
                name          => 'DCP117: shares of reallocation',
                cols          => $preAllocated->{cols},
                defaultFormat => '%softnz',
                custom        => [
                    '=IV11/SUM(IV21:IV22)',
                    $model->{DCP117} =~ /A/i
                    ? ()
                    : '=SUM(IV11:IV12)/SUM(IV21:IV22)',
                ],
                arguments => {
                    IV11 => $meavPercentages,
                    IV21 => $meavPercentages,
                },
                wsPrepare => sub {
                    my ( $self, $wb, $ws, $format, $formula, $pha, $rowh,
                        $colh ) = @_;
                    my $unavailable = $wb->getFormat('unavailable');
                    $model->{DCP117} =~ /A/i
                      ? sub {
                        my ( $x, $y ) = @_;
                        return -1, $format if $x == 0;
                        return '', $format, $formula->[0],
                          IV11 =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{IV11}, $colh->{IV11} + 1,
                            0, 1 ),
                          IV21 =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{IV11}, $colh->{IV11} + 1,
                            0, 1 ),
                          IV22 =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{IV11}, $colh->{IV11} + 3,
                            0, 1 ),
                          if $x == 1;
                        return '', $format, $formula->[0],
                          IV11 =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{IV11}, $colh->{IV11} + 2,
                            0, 1 ),
                          IV21 =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{IV11}, $colh->{IV11} + 1,
                            0, 1 ),
                          IV22 =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{IV11}, $colh->{IV11} + 3,
                            0, 1 ),
                          if $x == 2;
                        return '', $format, $formula->[0],
                          IV11 =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{IV11}, $colh->{IV11} + 3,
                            0, 1 ),
                          IV21 =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{IV11}, $colh->{IV11} + 1,
                            0, 1 ),
                          IV22 =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{IV11}, $colh->{IV11} + 3,
                            0, 1 ),
                          if $x == 3;
                        return '', $unavailable;
                      }
                      : sub {
                        my ( $x, $y ) = @_;
                        return -1, $format if $x == 0;
                        return 0,  $format if $x == 1;
                        return '', $format, $formula->[1],
                          IV11 =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{IV11}, $colh->{IV11} + 1,
                            0, 1 ),
                          IV12 =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{IV11}, $colh->{IV11} + 2,
                            0, 1 ),
                          IV21 =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{IV11}, $colh->{IV11} + 1,
                            0, 1 ),
                          IV22 =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{IV11}, $colh->{IV11} + 3,
                            0, 1 ),
                          if $x == 2;
                        return '', $format, $formula->[0],
                          IV11 =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{IV11}, $colh->{IV11} + 3,
                            0, 1 ),
                          IV21 =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{IV11}, $colh->{IV11} + 1,
                            0, 1 ),
                          IV22 =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{IV11}, $colh->{IV11} + 3,
                            0, 1 ),
                          if $x == 3;
                        return '', $unavailable;
                      };
                },
            );
            $preAllocated = Stack(
                name => 'Table 1330 allocated costs, after DCP117 adjustments',
                rows => $preAllocated->{rows},
                cols => $preAllocated->{cols},
                sources => [
                    Arithmetic(
                        name => 'Load related new connections & '
                          . 'reinforcement (net of contributions)'
                          . ' after DCP117',
                        rows          => $dcp117negative->{rows},
                        cols          => $preAllocated->{cols},
                        defaultFormat => '0softnz',
                        arithmetic    => '=IV1+IV2*IV3',
                        arguments     => {
                            IV1 => $preAllocated,
                            IV2 => $dcp117,
                            IV3 => $dcp117negative,
                        },
                    ),
                    $preAllocated
                ],
            );

        }

    }

    push @{ $model->{calcTables} }, $allocationRules,
      my $allocatedTotal = GroupBy(
        name          => 'Amounts already allocated',
        source        => $preAllocated,
        rows          => $preAllocated->{rows},
        defaultFormat => '0softnz'
      );

    my $allAllocationPercentages = Arithmetic(
        name          => 'All allocation percentages',
        defaultFormat => '%softnz',
        rows          => $preAllocated->{rows},
        cols          => $preAllocated->{cols},
        arithmetic =>
'=IF(IV45="MEAV",IV5,IF(IV46="EHV only",IV6,IF(IV47="LV only",IV7,IF(IV48="Network length",IV8,'
          . ( $model->{DCP097} ? 'IF(IV49="Customer numbers",IV9,0)' : '0' )
          . '))))',
        arguments => {
            IV45 => $allocationRules,
            IV5  => $meavPercentages,
            IV46 => $allocationRules,
            IV6  => Constant(
                name          => 'EHV only',
                cols          => $preAllocated->{cols},
                data          => [qw(0 0 0 1)],
                defaultFormat => '0connz'
            ),
            IV47 => $allocationRules,
            IV7  => Constant(
                name          => 'LV only',
                cols          => $preAllocated->{cols},
                data          => [qw(1 0 0 0)],
                defaultFormat => '0connz'
            ),
            IV48 => $allocationRules,
            IV8  => $networkLengthPercentages,
            $model->{DCP097}
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

    if ( $model->{DCP095} ) {

        my $allocLevelset95 = Labelset(
            name => 'Allocation levels DCP095',
            list => [ split /\n/, <<END_OF_LIST ] );
LV service
LV main
HV/LV
HV
EHV&132
END_OF_LIST

        my $lvOnly        = Labelset( list => ['LV'] );
        my $lvServiceOnly = Labelset( list => ['LV service'] );
        my $lvMainOnly    = Labelset( list => ['LV main'] );

        my $meavLvServProp =
          $model->meavPercentageServiceLV( $lvOnly, $lvServiceOnly );

        my $networkLengthLvServProp =
          $model->networkLengthPercentageServiceLV( $lvOnly, $lvServiceOnly );

        my $netCapexLvServProp =
          $model->netCapexPercentageServiceLV( $lvOnly, $lvServiceOnly );

        my $rrpLvServProp = Arithmetic(
            name          => 'Allocation to LV mains',
            defaultFormat => '%soft',
            arithmetic =>
'=IF(IV1="MEAV",IV5,IF(IV46="EHV only",0,IF(IV47="LV only",1,IF(IV48="Network length",IV8,0))))',
            arguments => {
                IV1  => $allocationRules,
                IV5  => $meavLvServProp,
                IV46 => $allocationRules,
                IV47 => $allocationRules,
                IV48 => $allocationRules,
                IV8  => $networkLengthLvServProp,
            },
        );

        my $afterAllocationLv = Stack(
            name    => 'Complete allocation: LV total share',
            cols    => $lvOnly,
            rows    => $allocationRules->{rows},
            sources => [$afterAllocation],
        );

        $afterAllocation = Stack(
            name => 'Complete allocation split between LV mains and services',
            rows => $allocationRules->{rows},
            cols => $allocLevelset95,
            sources => [
                Arithmetic(
                    name       => 'Allocation to LV service',
                    cols       => $lvServiceOnly,
                    arithmetic => '=IV1*IV2',
                    arguments =>
                      { IV1 => $afterAllocationLv, IV2 => $rrpLvServProp, },
                ),
                Arithmetic(
                    name       => 'Allocation to LV main',
                    cols       => $lvMainOnly,
                    arithmetic => '=IV1*(1-IV2)',
                    arguments =>
                      { IV1 => $afterAllocationLv, IV2 => $rrpLvServProp, },
                ),
                $afterAllocation,
            ]
        );

        my $netCapexLv = Stack(
            name    => 'Net capex: LV total share',
            cols    => $lvOnly,
            sources => [$netCapexPercentages],
        );

        $netCapexPercentages = Stack(
            name => 'Net capex allocation split between LV mains and services',
            cols => $allocLevelset95,
            sources => [
                Arithmetic(
                    name       => 'Allocation to LV service',
                    cols       => $lvServiceOnly,
                    arithmetic => '=IV1*IV2',
                    arguments =>
                      { IV1 => $netCapexLv, IV2 => $netCapexLvServProp, },
                ),
                Arithmetic(
                    name       => 'Allocation to LV main',
                    cols       => $lvMainOnly,
                    arithmetic => '=IV1*(1-IV2)',
                    arguments =>
                      { IV1 => $netCapexLv, IV2 => $netCapexLvServProp, },
                ),
                $netCapexPercentages,
            ]
        );

        my $unitsLv = Stack(
            name    => 'Units at LV',
            cols    => $lvOnly,
            sources => [$units],
        );

        $units = Stack(
            name    => 'Units',
            cols    => $allocLevelset95,
            sources => [
                Arithmetic(
                    name       => 'Allocation to LV service',
                    cols       => $lvServiceOnly,
                    arithmetic => '=IV1',
                    arguments  => { IV1 => $unitsLv, },
                ),
                Arithmetic(
                    name       => 'Allocation to LV main',
                    cols       => $lvMainOnly,
                    arithmetic => '=IV1',
                    arguments  => { IV1 => $unitsLv, },
                ),
                $units,
            ]
        );

        $allocLevelset = $allocLevelset95;

    }    # end if DCP095

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

    push @{ $model->{calcTables} }, $alloc, $direct;

    my $dcp071 = $model->{DCP071} || $model->{DCP071A};

    if ( $model->{DCP095} ) {

        my $lvServiceAllocation = Stack(
            name    => 'LV service allocation',
            cols    => Labelset( list => ['LV service'] ),
            sources => [$alloc]
        );
        my $lvMainAllocation = Stack(
            name    => 'LV main allocation',
            cols    => Labelset( list => ['LV main'] ),
            sources => [$alloc]
        );
        my $hvLvAllocation = Stack(
            name    => 'HV/LV allocation',
            cols    => Labelset( list => ['HV/LV'] ),
            sources => [$alloc]
        );
        my $hvAllocation = Stack(
            name    => 'HV allocation',
            cols    => Labelset( list => ['HV'] ),
            sources => [$alloc]
        );
        push @{ $model->{impactTables} },
          Columnset(
            name    => 'Allocations to network levels',
            columns => [
                $lvServiceAllocation, $lvMainAllocation,
                $hvLvAllocation,      $hvAllocation,
            ]
          );

        my $lvDirect = Stack(
            name    => 'LV direct proportion',
            cols    => Labelset( list => ['LV'] ),
            sources => [$direct]
        );
        my $hvDirect = Stack(
            name    => 'HV direct proportion',
            cols    => Labelset( list => ['HV'] ),
            sources => [$direct]
        );
        push @{ $model->{impactTables} },
          Columnset(
            name    => 'Direct cost proportions',
            columns => [ $lvDirect, $hvDirect ]
          );

        my $discount = Columnset(
            name    => 'LDNO discounts',
            columns => [
                Arithmetic(
                    name       => 'LDNO LV: LV user',
                    arithmetic => '=IV4+IV1*(1-IV2*IV3)',
                    arguments  => {
                        IV4 => $lvServiceAllocation,
                        IV1 => $lvMainAllocation,
                        IV2 => $lvSplit,
                        IV3 => $lvDirect,
                    },
                    defaultFormat => '%soft',
                ),
                Arithmetic(
                    name       => 'LDNO HV: LV user',
                    arithmetic => $dcp071 ? '=IV9+IV1+IV2+IV3*(1-IV4*IV5)'
                    : '=IV1+IV2',
                    arguments => {
                        IV9 => $lvServiceAllocation,
                        IV1 => $lvMainAllocation,
                        IV2 => $hvLvAllocation,
                        $dcp071
                        ? (
                            IV3 => $hvAllocation,
                            IV4 => $hvSplit,
                            IV5 => $hvDirect
                          )
                        : (),
                    },
                    defaultFormat => '%soft',
                ),
                Arithmetic(
                    name       => 'LDNO HV: LV Sub user',
                    arithmetic => $dcp071
                    ? '=(IV2+IV3*(1-IV4*IV5))/(1-IV1-IV9)'
                    : '=IV2/(1-IV1-IV9)',
                    arguments => {
                        IV1 => $lvMainAllocation,
                        IV9 => $lvServiceAllocation,
                        IV2 => $hvLvAllocation,
                        $dcp071
                        ? (
                            IV3 => $hvAllocation,
                            IV4 => $hvSplit,
                            IV5 => $hvDirect
                          )
                        : (),
                    },
                    defaultFormat => '%soft',
                ),
                Arithmetic(
                    name       => 'LDNO HV: HV user',
                    arithmetic => '=IV1*(1-IV2*IV3)/(1-IV9-IV4-IV5)',
                    arguments  => {
                        IV1 => $hvAllocation,
                        IV2 => $hvSplit,
                        IV3 => $hvDirect,
                        IV4 => $lvMainAllocation,
                        IV9 => $lvServiceAllocation,
                        IV5 => $hvLvAllocation,
                    },
                    defaultFormat => '%soft',
                ),
            ]
        );

        push @{ $model->{impactTables} }, $discount;

        my ($discountCurrent) = $model->checks($allocLevelset);

        push @{ $model->{impactTables} }, Columnset(
            name    => 'Change from current discounts',
            columns => [
                map {
                    Arithmetic(
                        arithmetic => '=IV1-IV2',
                        name       => $discountCurrent->{columns}[$_]{name},
                        arguments  => {
                            IV1 => $discount->{columns}[$_],
                            IV2 => $discountCurrent->{columns}[$_]
                        },
                        defaultFormat => [
                            base => '%softpm',
                            num_format =>
                              '[Blue]+??0.000%;[Red]-??0.000%;[Green]=',
                        ],
                      )
                } 0 .. $#{ $discountCurrent->{columns} }
            ]
        );

    }

    else {    # not DCP095

        my $lvAllocation = Stack(
            name    => 'LV allocation',
            cols    => Labelset( list => ['LV'] ),
            sources => [$alloc]
        );
        my $hvLvAllocation = Stack(
            name    => 'HV/LV allocation',
            cols    => Labelset( list => ['HV/LV'] ),
            sources => [$alloc]
        );
        my $hvAllocation = Stack(
            name    => 'HV allocation',
            cols    => Labelset( list => ['HV'] ),
            sources => [$alloc]
        );
        push @{ $model->{impactTables} },
          Columnset(
            name    => 'Allocations to network levels',
            columns => [ $lvAllocation, $hvLvAllocation, $hvAllocation ]
          );

        my $lvDirect = Stack(
            name    => 'LV direct proportion',
            cols    => Labelset( list => ['LV'] ),
            sources => [$direct]
        );
        my $hvDirect = Stack(
            name    => 'HV direct proportion',
            cols    => Labelset( list => ['HV'] ),
            sources => [$direct]
        );
        push @{ $model->{impactTables} },
          Columnset(
            name    => 'Direct cost proportions',
            columns => [ $lvDirect, $hvDirect ]
          );

        my $discount = Columnset(
            name    => 'LDNO discounts',
            columns => [
                Arithmetic(
                    name       => 'LDNO LV: LV user',
                    arithmetic => '=IV1*(1-IV2*IV3)',
                    arguments  => {
                        IV1 => $lvAllocation,
                        IV2 => $lvSplit,
                        IV3 => $lvDirect
                    },
                    defaultFormat => '%soft',
                ),
                Arithmetic(
                    name       => 'LDNO HV: LV user',
                    arithmetic => $dcp071 ? '=IV1+IV2+IV3*(1-IV4*IV5)'
                    : '=IV1+IV2',
                    arguments => {
                        IV1 => $lvAllocation,
                        IV2 => $hvLvAllocation,
                        $dcp071
                        ? (
                            IV3 => $hvAllocation,
                            IV4 => $hvSplit,
                            IV5 => $hvDirect
                          )
                        : (),
                    },
                    defaultFormat => '%soft',
                ),
                Arithmetic(
                    name       => 'LDNO HV: LV Sub user',
                    arithmetic => $dcp071 ? '=(IV2+IV3*(1-IV4*IV5))/(1-IV1)'
                    : '=IV2/(1-IV1)',
                    arguments => {
                        IV1 => $lvAllocation,
                        IV2 => $hvLvAllocation,
                        $dcp071
                        ? (
                            IV3 => $hvAllocation,
                            IV4 => $hvSplit,
                            IV5 => $hvDirect
                          )
                        : (),
                    },
                    defaultFormat => '%soft',
                ),
                Arithmetic(
                    name       => 'LDNO HV: HV user',
                    arithmetic => '=IV1*(1-IV2*IV3)/(1-IV4-IV5)',
                    arguments  => {
                        IV1 => $hvAllocation,
                        IV2 => $hvSplit,
                        IV3 => $hvDirect,
                        IV4 => $lvAllocation,
                        IV5 => $hvLvAllocation,
                    },
                    defaultFormat => '%soft',
                ),
            ]
        );

        push @{ $model->{impactTables} }, $discount;

        my ($discountCurrent) = $model->checks($allocLevelset);

        push @{ $model->{impactTables} }, Columnset(
            name    => 'Change from current discounts',
            columns => [
                map {
                    Arithmetic(
                        arithmetic => '=IV1-IV2',
                        name       => $discountCurrent->{columns}[$_]{name},
                        arguments  => {
                            IV1 => $discount->{columns}[$_],
                            IV2 => $discountCurrent->{columns}[$_]
                        },
                        defaultFormat => '%softpm'
                      )
                } 0 .. $#{ $discountCurrent->{columns} }
            ]
        );

    }    # end if not DCP095

    $model;
}

1;
