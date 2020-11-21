package CDCM;

# Copyright 2020 Franck Latrémolière and others.
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

sub domesticTcr {

    my ( $model, $components, $days, $volumes, $matchingTables ) = @_;

    # Each element in @$matchingTables is a hashref,
    #   with tariff component names as keys and datasets as values.
    # This method changes the contents of $matchingTables and returns nothing.

    my %domesticFlagByTariffset;
    my @componentsToRemove = grep { /kWh/; } @$components;
    my ($componentToUseInstead) = grep { /MPAN/; } @$components;
    foreach my $matchingTable (@$matchingTables) {
        my @componentsPresentToRemove =
          grep { $matchingTable->{$_}; } @componentsToRemove;
        next unless @componentsPresentToRemove;
        my $tariffSet = $matchingTable->{ $componentsPresentToRemove[0] }{rows};
        my $cols      = $matchingTable->{ $componentsPresentToRemove[0] }{cols};
        unless ( $domesticFlagByTariffset{ 0 + $tariffSet } ) {
            my @domesticIndicator =
              map { /domestic/i && !/non.domestic/i ? 1 : 0; }
              @{ $tariffSet->{list} };
            next unless grep { $_; } @domesticIndicator;
            $domesticFlagByTariffset{ 0 + $tariffSet } ||= Constant(
                name          => 'Domestic customer indicator',
                defaultFormat => '0con',
                rows          => $tariffSet,
                data          => \@domesticIndicator,
            );
        }
        my $fixedChargeAdder = Arithmetic(
            name => 'Fixed charge adder (p/MPAN/day) '
              . 'to offset removed volume adders',
            defaultFormat => '0.00soft',
            arithmetic    => '=1000*('
              . join(
                '+',
                map { "SUMPRODUCT(A1${_}_A2${_}*A3${_}_A4${_}*A5${_}_A6${_})"; }
                  0 .. $#componentsPresentToRemove
              )
              . ')/SUMPRODUCT(A71_A72*A73_A74)/A8',
            arguments => {
                A71_A72 => $domesticFlagByTariffset{ 0 + $tariffSet },
                A73_A74 => $volumes->{$componentToUseInstead},
                A8      => $days,
                map {
                    (
                        "A1${_}_A2${_}" =>
                          $domesticFlagByTariffset{ 0 + $tariffSet },
                        "A3${_}_A4${_}" =>
                          $volumes->{ $componentsPresentToRemove[$_] },
                        "A5${_}_A6${_}" =>
                          $matchingTable->{ $componentsPresentToRemove[$_] },
                    );
                } 0 .. $#componentsPresentToRemove
            },
        );
        $matchingTable->{$_} = Arithmetic(
            name          => $_ . ' adder except for domestic customers',
            defaultFormat => $matchingTable->{$_}{defulatFormat},
            arithmetic    => '=(1-A2)*A1',
            arguments     => {
                A1 => $matchingTable->{$_},
                A2 => $domesticFlagByTariffset{ 0 + $tariffSet },
            },
        ) foreach @componentsPresentToRemove;
        $matchingTable->{$componentToUseInstead} = Arithmetic(
            name          => $_ . ' adjusted for domestic customers',
            defaultFormat => '0.00soft',
            cols          => $cols,
            arithmetic    => '=A1*A2'
              . ( $matchingTable->{$componentToUseInstead} ? '+A9' : '' ),
            arguments => {
                A1 => $domesticFlagByTariffset{ 0 + $tariffSet },
                A2 => $fixedChargeAdder,
                $matchingTable->{$componentToUseInstead}
                ? ( A9 => $matchingTable->{$componentToUseInstead} )
                : (),
            },
        );
    }
}

1;
