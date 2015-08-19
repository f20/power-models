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
LV services
LV mains
HV/LV
HV
EHV&132
END_OF_LIST

    my $lvOnly = $model->{objects}{lvOnly} ||= Labelset( list => ['LV'] );
    my $lvServiceOnly = $model->{objects}{lvServiceOnly} ||=
      Labelset( list => ['LV services'] );
    my $lvMainOnly = $model->{objects}{lvMainOnly} ||=
      Labelset( list => ['LV mains'] );

    my $meavLvServProp =
      $model->meavPercentageServiceLV( $lvOnly, $lvServiceOnly );

    my $networkLengthLvServProp;
    $networkLengthLvServProp =
      $model->networkLengthPercentageServiceLV( $lvOnly, $lvServiceOnly )
      if $model->{allowNetworkLength};

    my $netCapexLvServProp =
      $model->netCapexPercentageServiceLV( $lvOnly, $lvServiceOnly );

    my $rrpLvServProp = Arithmetic(
        name          => 'Allocation of LV to LV services',
        defaultFormat => '%soft',
        arithmetic    => '=IF(A1="60%MEAV",0.4+0.6*A51,'
          . 'IF(A45="MEAV",A52,'
          . 'IF(A46="EHV only",0,'
          . 'IF(A47="LV only",1,'
          . ( $networkLengthLvServProp ? 'IF(A48="Network length",A8,0)' : '0' )
          . '))))',
        arguments => {
            A1  => $allocationRules,
            A45 => $allocationRules,
            A46 => $allocationRules,
            A47 => $allocationRules,
            A48 => $allocationRules,
            A51 => $meavLvServProp,
            A52 => $meavLvServProp,
            $networkLengthLvServProp ? ( A8 => $networkLengthLvServProp ) : (),
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
                name       => 'Allocation to LV services',
                cols       => $lvServiceOnly,
                arithmetic => '=A1*A2',
                arguments =>
                  { A1 => $afterAllocationLv, A2 => $rrpLvServProp, },
                defaultFormat => '0soft',
            ),
            Arithmetic(
                name       => 'Allocation to LV mains',
                cols       => $lvMainOnly,
                arithmetic => '=A1*(1-A2)',
                arguments =>
                  { A1 => $afterAllocationLv, A2 => $rrpLvServProp, },
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
                name       => 'Allocation to LV services',
                cols       => $lvServiceOnly,
                arithmetic => '=A1*A2',
                arguments  => { A1 => $netCapexLv, A2 => $netCapexLvServProp, },
                defaultFormat => '%soft',
            ),
            Arithmetic(
                name       => 'Allocation to LV mains',
                cols       => $lvMainOnly,
                arithmetic => '=A1*(1-A2)',
                arguments  => { A1 => $netCapexLv, A2 => $netCapexLvServProp, },
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
        name          => 'Units',
        defaultFormat => '0copy',
        cols          => $allocLevelset95,
        sources       => [
            Arithmetic(
                name          => 'Allocation to LV services',
                cols          => $lvServiceOnly,
                arithmetic    => '=A1',
                arguments     => { A1 => $unitsLv, },
                defaultFormat => '0copy',
            ),
            Arithmetic(
                name          => 'Allocation to LV mains',
                cols          => $lvMainOnly,
                arithmetic    => '=A1',
                arguments     => { A1 => $unitsLv, },
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
        name    => 'LV services allocation',
        cols    => Labelset( list => ['LV services'] ),
        sources => [$alloc]
    );
    my $lvMainAllocation = Stack(
        name    => 'LV mains allocation',
        cols    => Labelset( list => ['LV mains'] ),
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

    Columnset(
        name    => 'Direct cost proportions',
        columns => [ $lvDirect, $hvDirect ]
    );

    my @columns = (
        Arithmetic(
            name       => 'LDNO LV: LV user',
            arithmetic => $model->{fixedIndirectPercentage}
            ? '=A4+A1*(1-A2)'
            : '=A4+A1*(1-A2*A3)',
            arguments => {
                A4 => $lvServiceAllocation,
                A1 => $lvMainAllocation,
                A2 => $lvSplit,
                $model->{fixedIndirectPercentage}
                ? ()
                : ( A3 => $lvDirect ),
            },
            defaultFormat => '%soft',
        ),
        Arithmetic(
            name       => 'LDNO HV: LV user',
            arithmetic => $dcp071
            ? (
                $model->{fixedIndirectPercentage}
                ? '=A9+A1+A2+A3*(1-A4)'
                : '=A9+A1+A2+A3*(1-A4*A5)'
              )
            : '=A9+A1+A2',
            arguments => {
                A9 => $lvServiceAllocation,
                A1 => $lvMainAllocation,
                A2 => $hvLvAllocation,
                $dcp071
                ? (
                    A3 => $hvAllocation,
                    A4 => $hvSplit,
                    $model->{fixedIndirectPercentage}
                    ? ()
                    : ( A5 => $hvDirect ),
                  )
                : (),
            },
            defaultFormat => '%soft',
        ),
        Arithmetic(
            name       => 'LDNO HV: LV Sub user',
            arithmetic => $dcp071
            ? (
                $model->{fixedIndirectPercentage}
                ? '=(A2+A3*(1-A4))/(1-A1-A9)'
                : '=(A2+A3*(1-A4*A5))/(1-A1-A9)'
              )
            : '=A2/(1-A1-A9)',
            arguments => {
                A1 => $lvMainAllocation,
                A9 => $lvServiceAllocation,
                A2 => $hvLvAllocation,
                $dcp071
                ? (
                    A3 => $hvAllocation,
                    A4 => $hvSplit,
                    $model->{fixedIndirectPercentage}
                    ? ()
                    : ( A5 => $hvDirect ),
                  )
                : (),
            },
            defaultFormat => '%soft',
        ),
        Arithmetic(
            name       => 'LDNO HV: HV user',
            arithmetic => $model->{fixedIndirectPercentage}
            ? '=A1*(1-A2)/(1-A9-A4-A5)'
            : '=A1*(1-A2*A3)/(1-A9-A4-A5)',
            arguments => {
                A1 => $hvAllocation,
                A2 => $hvSplit,
                $model->{fixedIndirectPercentage}
                ? ()
                : ( A3 => $hvDirect ),
                A4 => $lvMainAllocation,
                A9 => $lvServiceAllocation,
                A5 => $hvLvAllocation,
            },
            defaultFormat => '%soft',
        ),
    );

    push @{ $model->{objects}{calcSheets} },
      [
        $model->{suffix},
        map { @{ $_->{arguments} }{ sort keys %{ $_->{arguments} } }; }
          @columns
      ];

    push @columns, map {
        my $digits = /([0-9])/ ? $1 : 6;
        SpreadsheetModel::Checksum->new(
            name => $_,
            /recursive|model/i ? ( recursive => 1 ) : (),
            digits  => $digits,
            columns => [@columns],
            factors => [ map { 1000 } @columns ]
        );
      } split /;\s*/, $model->{checksums}
      if $model->{checksums};

    my $discount = Columnset(
        name    => 'LDNO discounts (CDCM)',
        columns => \@columns,
    );

    push @{ $model->{objects}{resultsTables} }, $discount;

    if ( $model->{table1399} ) {
        my ($discountCurrent) = $model->checks($allocLevelset);
        push @{ $model->{objects}{resultsTables} }, Columnset(
            name    => 'Change from current discounts',
            columns => [
                map {
                    Arithmetic(
                        arithmetic => '=A1-A2',
                        name       => $discountCurrent->{columns}[$_]{name},
                        arguments  => {
                            A1 => $discount->{columns}[$_],
                            A2 => $discountCurrent->{columns}[$_]
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

}

1;
