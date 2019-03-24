package ModelM;

# Copyright 2011 The Competitive Networks Association and others.
# Copyright 2012-2017 Franck Latrémolière, Reckon LLP and others.
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

sub discounts {    # Not used if DCP 095

    my ( $model, $alloc, $allocLevelset, $dcp071, $direct, $hvSplit, $lvSplit, )
      = @_;

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

    Columnset(
        name    => 'Allocations to network levels',
        columns => [ $lvAllocation, $hvLvAllocation, $hvAllocation ]
    );

    my $lvDirect = Stack(
        name    => 'LV direct proportion',
        cols    => Labelset( list => ['LV'] ),
        sources => [$direct]
    ) unless $model->{fixedIndirectPercentage};
    my $hvDirect = Stack(
        name    => 'HV direct proportion',
        cols    => Labelset( list => ['HV'] ),
        sources => [$direct]
    ) unless $model->{fixedIndirectPercentage};

    Columnset(
        name    => 'HV and LV direct cost proportions',
        columns => [ $lvDirect, $hvDirect ]
    ) unless $model->{fixedIndirectPercentage};

    my @columns = (
        Arithmetic(
            name => $model->{qno} . ' LV: LV user',
            cols => Labelset( list => [ $model->{qno} . ' LV: LV user' ] ),
            arithmetic => $model->{fixedIndirectPercentage}
            ? '=A1*(1-A2)'
            : '=A1*(1-A2*A3)',
            arguments => {
                A1 => $lvAllocation,
                A2 => $lvSplit,
                $model->{fixedIndirectPercentage}
                ? ()
                : ( A3 => $lvDirect ),
            },
            defaultFormat => '%soft',
        ),
        Arithmetic(
            name => $model->{qno} . ' HV: LV user',
            cols => Labelset( list => [ $model->{qno} . ' HV: LV user' ] ),
            arithmetic => $dcp071
            ? (
                $model->{fixedIndirectPercentage}
                ? '=A1+A2+A3*(1-A4)'
                : '=A1+A2+A3*(1-A4*A5)'
              )
            : '=A1+A2',
            arguments => {
                A1 => $lvAllocation,
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
            name => $model->{qno} . ' HV: LV Sub user',
            cols => Labelset( list => [ $model->{qno} . ' HV: LV Sub user' ] ),
            arithmetic => $dcp071
            ? (
                $model->{fixedIndirectPercentage}
                ? '=(A2+A3*(1-A4))/(1-A1)'
                : '=(A2+A3*(1-A4*A5))/(1-A1)'
              )
            : '=A2/(1-A1)',
            arguments => {
                A1 => $lvAllocation,
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
            name => $model->{qno} . ' HV: HV user',
            cols => Labelset( list => [ $model->{qno} . ' HV: HV user' ] ),
            arithmetic => $model->{fixedIndirectPercentage}
            ? '=A1*(1-A2)/(1-A4-A5)'
            : '=A1*(1-A2*A3)/(1-A4-A5)',
            arguments => {
                A1 => $hvAllocation,
                A2 => $hvSplit,
                $model->{fixedIndirectPercentage}
                ? ()
                : ( A3 => $hvDirect ),
                A4 => $lvAllocation,
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

    push @{ $model->{objects}{table1037sources} },
      grep { $_->{cols}; } @columns;

    push @columns, map {
        my $digits = /([0-9])/ ? $1 : 6;
        SpreadsheetModel::Checksum->new(
            name => $_,
            /table|recursive|model/i ? ( recursive => 1 ) : (),
            digits  => $digits,
            columns => [@columns],
            factors => [ map { 10000 } @columns ]
        );
      } split /;\s*/, $model->{checksums}
      if $model->{checksums};

    my $discount = Columnset(
        name => $model->{qno} . ' discounts (CDCM) ⇒1037. For CDCM model',
        singleRowName => $model->{qno} . ' discount',
        columns       => \@columns,
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
                        defaultFormat => '%softpm'
                      )
                } 0 .. $#{ $discountCurrent->{columns} }
            ]
        );
    }

    $discount;

}

1;
