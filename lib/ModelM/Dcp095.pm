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

sub realloc95 {

    my ( $model, $afterAllocation, $allocationRules, $netCapexPercentages,
        $units, )
      = @_;

    my $allocLevelset95 = Labelset(
        name => 'Allocation levels DCP 095',
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
        name    => 'Complete allocation split between LV mains and services',
        rows    => $allocationRules->{rows},
        cols    => $allocLevelset95,
        sources => [
            Arithmetic(
                name       => 'Allocation to LV service',
                cols       => $lvServiceOnly,
                arithmetic => '=IV1*IV2',
                arguments =>
                  { IV1 => $afterAllocationLv, IV2 => $rrpLvServProp, },
                defaultFormat => '0soft',
            ),
            Arithmetic(
                name       => 'Allocation to LV main',
                cols       => $lvMainOnly,
                arithmetic => '=IV1*(1-IV2)',
                arguments =>
                  { IV1 => $afterAllocationLv, IV2 => $rrpLvServProp, },
                defaultFormat => '0soft',
            ),
            $afterAllocation,
        ],
        defaultFormat => '0copy',
    );

    my $netCapexLv = Stack(
        name    => 'Net capex: LV total share',
        cols    => $lvOnly,
        sources => [$netCapexPercentages],
    );

    $netCapexPercentages = Stack(
        name    => 'Net capex allocation split between LV mains and services',
        cols    => $allocLevelset95,
        sources => [
            Arithmetic(
                name       => 'Allocation to LV service',
                cols       => $lvServiceOnly,
                arithmetic => '=IV1*IV2',
                arguments =>
                  { IV1 => $netCapexLv, IV2 => $netCapexLvServProp, },
                defaultFormat => '%soft',
            ),
            Arithmetic(
                name       => 'Allocation to LV main',
                cols       => $lvMainOnly,
                arithmetic => '=IV1*(1-IV2)',
                arguments =>
                  { IV1 => $netCapexLv, IV2 => $netCapexLvServProp, },
                defaultFormat => '%soft',
            ),
            $netCapexPercentages,
        ],
        defaultFormat => '%copy',
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
                name          => 'Allocation to LV service',
                cols          => $lvServiceOnly,
                arithmetic    => '=IV1',
                arguments     => { IV1 => $unitsLv, },
                defaultFormat => '0copy',
            ),
            Arithmetic(
                name          => 'Allocation to LV main',
                cols          => $lvMainOnly,
                arithmetic    => '=IV1',
                arguments     => { IV1 => $unitsLv, },
                defaultFormat => '0copy',
            ),
            $units,
        ]
    );

    $afterAllocation, $allocLevelset95, $netCapexPercentages, $units;

}

sub discounts95 {

    my ( $model, $alloc, $allocLevelset, $dcp071, $direct, $hvSplit, $lvSplit, )
      = @_;

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
                arithmetic => $dcp071 ? '=(IV2+IV3*(1-IV4*IV5))/(1-IV1-IV9)'
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
                        base       => '%softpm',
                        num_format => '[Blue]+??0.000%;[Red]-??0.000%;[Green]=',
                    ],
                  )
            } 0 .. $#{ $discountCurrent->{columns} }
        ]
    );

}

1;
