package Elec::Profits;

# Copyright 2022 Franck Latrémolière and others.
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
use SpreadsheetModel::CalcBlock;

sub stripPounds {
    ( local $_ ) = @_;
    s/[ (]*£\/year\)?/ /;
    $_;
}

sub new {
    my ( $class, $model, $energyMargin, $distributionRevenueExt, $tariffs,
        $volumes, @distributionCosts )
      = @_;
    $model->register( my $self = bless { model => $model, }, $class );

    my @units               = grep { $_->{name} =~ /kWh/; } @$volumes;
    my @energyChargePerUnit = grep { $_->{name} =~ /kWh/; }
      @{ $model->{buildupTables}[ $#{ $model->{buildupTables} } ]{columns} };
    push @distributionCosts,
      Arithmetic(
        name          => 'Cost of distribution losses £/year',
        defaultFormat => '0soft',
        arithmetic    => '=0.01*SUMPRODUCT(A1_A2,A3_A4)',
        arguments => { A1_A2 => $energyChargePerUnit[0], A3_A4 => $units[0], },
      ) if @units == 1 && @energyChargePerUnit == 1;

    my $distributionRevenueInt = Arithmetic(
        name          => 'Notional use of system revenue from internal users',
        defaultFormat => '0soft',
        arithmetic    => '=A1-A2',
        arguments     => {
            A1 => $tariffs->revenueCalculation($volumes),
            A2 => $distributionRevenueExt,
        },
    );
    push @{ $self->{revenueTables} }, CalcBlock(
        name  => 'Profitability summary £/year',
        items => [
            [
                GroupBy(
                    name => 'Use of system revenue from external users',
                    defaultFormat => '0soft',
                    source        => $distributionRevenueExt,
                ),
                $energyMargin,
                GroupBy(
                    name =>
                      'Notional use of system revenue from internal users',
                    defaultFormat => '0soft',
                    source        => $distributionRevenueInt,
                ),
                'Total income above'
            ],
            (
                map {
                    Arithmetic(
                        name => 'Less '
                          . lcfirst( stripPounds( $_->objectShortName ) ),
                        defaultFormat => '0soft',
                        $_->{cols}
                        ? (
                            arithmetic => '=0-SUM(A2_A4)',
                            arguments  => { A2_A4 => $_, },
                          )
                        : (
                            arithmetic => '=0-A1',
                            arguments  => { A1 => $_, },
                        ),
                    );
                } @distributionCosts
            ),
            'Net income',
        ]
    );
    $self;
}

sub finish {
    my ($self) = @_;
    push @{ $self->{model}{revenueTables} }, @{ $self->{revenueTables} };
}

1;
