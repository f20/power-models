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

sub discounts { # Not used if DCP 095

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
    push @{ $model->{impactTables} },
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
    push @{ $model->{calcTables} },
      Columnset(
        name    => 'Direct cost proportions',
        columns => [ $lvDirect, $hvDirect ]
      ) unless $model->{fixedIndirectPercentage};

    my @columns = (
        Arithmetic(
            name       => 'LDNO LV: LV user',
            arithmetic => $model->{fixedIndirectPercentage}
            ? '=IV1*(1-IV2)'
            : '=IV1*(1-IV2*IV3)',
            arguments => {
                IV1 => $lvAllocation,
                IV2 => $lvSplit,
                $model->{fixedIndirectPercentage}
                ? ()
                : ( IV3 => $lvDirect ),
            },
            defaultFormat => '%soft',
        ),
        Arithmetic(
            name       => 'LDNO HV: LV user',
            arithmetic => $dcp071
            ? (
                $model->{fixedIndirectPercentage}
                ? '=IV1+IV2+IV3*(1-IV4)'
                : '=IV1+IV2+IV3*(1-IV4*IV5)'
              )
            : '=IV1+IV2',
            arguments => {
                IV1 => $lvAllocation,
                IV2 => $hvLvAllocation,
                $dcp071
                ? (
                    IV3 => $hvAllocation,
                    IV4 => $hvSplit,
                    $model->{fixedIndirectPercentage}
                    ? ()
                    : ( IV5 => $hvDirect ),
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
                ? '=(IV2+IV3*(1-IV4))/(1-IV1)'
                : '=(IV2+IV3*(1-IV4*IV5))/(1-IV1)'
              )
            : '=IV2/(1-IV1)',
            arguments => {
                IV1 => $lvAllocation,
                IV2 => $hvLvAllocation,
                $dcp071
                ? (
                    IV3 => $hvAllocation,
                    IV4 => $hvSplit,
                    $model->{fixedIndirectPercentage}
                    ? ()
                    : ( IV5 => $hvDirect ),
                  )
                : (),
            },
            defaultFormat => '%soft',
        ),
        Arithmetic(
            name       => 'LDNO HV: HV user',
            arithmetic => $model->{fixedIndirectPercentage}
            ? '=IV1*(1-IV2)/(1-IV4-IV5)'
            : '=IV1*(1-IV2*IV3)/(1-IV4-IV5)',
            arguments => {
                IV1 => $hvAllocation,
                IV2 => $hvSplit,
                $model->{fixedIndirectPercentage}
                ? ()
                : ( IV3 => $hvDirect ),
                IV4 => $lvAllocation,
                IV5 => $hvLvAllocation,
            },
            defaultFormat => '%soft',
        ),
    );

    push @columns, map {
        SpreadsheetModel::Checksum->new(
            name => $_,
            /recursive|model/i ? ( recursive => 1 ) : (),
            digits => /([0-9])/ ? $1 : 6,
            columns => [@columns],
            factors => [ map { 1000 } @columns ]
        );
      } split /;\s*/, $model->{checksums}
      if $model->{checksums};

    my $discount = Columnset(
        name    => 'LDNO discounts',
        columns => \@columns,
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

}

1;
