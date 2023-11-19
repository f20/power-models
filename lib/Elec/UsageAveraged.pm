package Elec::UsageAveraged;

# Copyright 2023 Franck Latrémolière and others.
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
use base 'Elec::Usage';
use SpreadsheetModel::Custom;

sub usageRates {
    my ($self) = @_;
    return $self->{usageRates} if $self->{usageRates};
    my ( $model, $setup, $customers, $usage, $uset ) =
      @{$self}{qw(model setup customers usage uset)};
    my $usageRates1 = $usage->usageRates;
    my $customers1  = $customers->{customers};
    my $tariffSet1  = $customers1->tariffSet;
    my $tariffSet   = $customers->tariffSet;
    my @usageRates;

    for ( my $c = 0 ; $c < @$usageRates1 ; ++$c ) {
        push @usageRates, SpreadsheetModel::Custom->new(
            name => $usageRates1->[$c]->objectShortName
              . ( $self->{suffix} || '' ),
            rows       => $tariffSet,
            cols       => $usageRates1->[$c]{cols},
            arithmetic => '=IF(A41,SUMPRODUCT(A1*A2*A3)/A4,0)',
            custom     => [ '=IF(A41,SUMPRODUCT(A1:A10*A2:A20*A3:A30)/A4,0)', ],
            arguments  => {
                A1  => $customers->matrix,
                A10 => $customers->matrix,
                A2  => $customers1->totalDemand($uset)->[$c],
                A20 => $customers1->totalDemand($uset)->[$c],
                A3  => $usageRates1->[$c],
                A30 => $usageRates1->[$c],
                A4  => $customers->totalDemand($uset)->[$c],
                A41 => $customers->totalDemand($uset)->[$c],
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    '', $format, $formula->[0], map {
                        qr/\b$_\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{$_} + (
                                  /A4/  ? $y
                                : /A.0/ ? $customers->matrix->lastRow
                                : 0
                            ),
                            $colh->{$_} + (
                                /A4/ ? $x
                                : $y
                            ),
                            /A4/ ? () : ( 1, 1 ),
                          )
                    } @$pha;
                };
            },
        );
    }
    $self->{usageRates} = \@usageRates;
}

1;
