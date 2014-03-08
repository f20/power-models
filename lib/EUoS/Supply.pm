package EUoS::Supply;

=head Copyright licence and disclaimer

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
use base 'EUoS::Tariffs';
use SpreadsheetModel::Shortcuts ':all';

sub new {
    my ( $class, $model, $setup, $tariffs, $basicEnergyCharge ) = @_;
    my $uosTariffs = $tariffs->tariffs;
    my $self       = bless {
        uos               => $tariffs,
        model             => $model,
        setup             => $setup,
        basicEnergyCharge => $basicEnergyCharge,
    }, $class;
    $self->{tariffs} = [
        Arithmetic(
            name       => 'Supply p/kWh',
            arithmetic => '=IV1+IV2',
            arguments  => {
                IV1 => Dataset(
                    name     => 'Competitive energy charging rate p/kWh',
                    rows     => $uosTariffs->[0]{rows},
                    number   => 1588,
                    appendTo => $self->{model}{inputTables},
                    dataset  => $self->{model}{dataset},
                    data     => [ map { 10 } @{ $uosTariffs->[0]{rows}{list} } ],
                ),
                IV2 => $uosTariffs->[0],
            },
        ),
        map {
            Stack(
                name    => $uosTariffs->[$_]{name},
                sources => [ $uosTariffs->[$_] ]
              )
        } 1 .. 2,
    ];
    $self;
}

sub tariffName { 'supply tariffs'; }

sub marginCalculation {
    my ( $self, $volumes, $labelTail ) = @_;
    $labelTail ||= '';
    my $tariffs = $self->{tariffs};
    my $uos     = $self->{uos};
    Arithmetic(
        name       => 'Energy supply margin £/year' . $labelTail,
        arithmetic => '=IV1*(IV11-IV12-IV13)/100',
        arguments  => {
            IV1  => $volumes->[0],
            IV11 => $tariffs->[0],
            IV12 => $uos->{tariffs}[0],
            IV13 => $self->{basicEnergyCharge},
        },
        defaultFormat => '0softnz',
    );
}

sub margin {
    my ( $self, $volumes ) = @_;
    my $labelTail =
      $volumes->[0]{usetName} ? " for $volumes->[0]{usetName}" : '';
    my $revenues = $self->marginCalculation( $volumes, $labelTail );
    push @{ $self->{revenueTables} },
      GroupBy(
        name          => 'Total energy supply margin £/year' . $labelTail,
        defaultFormat => '0softnz',
        source        => $revenues,
      );
}
1;
