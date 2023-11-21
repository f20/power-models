package Elec::CustomersGrouped;

# Copyright 2023 Franck Latrémolière, Reckon LLP and others.
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
use Elec::UsageAveraged;

sub new {
    my ( $class, $model, $setup, $customers ) = @_;
    $model->register(
        bless {
            model     => $model,
            setup     => $setup,
            customers => $customers,
        },
        $class
    );

}

sub tariffSet {
    my ($self) = @_;
    return $self->{tariffSet} if $self->{tariffSet};
    my $oldList = $self->{customers}->tariffSet->{list};
    my ( @list, %map, @data );
    for ( my $i = 0 ; $i < @$oldList ; ++$i ) {
        my $routedTariff   = $oldList->[$i];
        my $chargingTariff = $routedTariff;
        $chargingTariff =~ s/ via .*//;
        my $j = $map{$chargingTariff};
        unless ( defined $j ) {
            push @list, $chargingTariff;
            $map{$chargingTariff} = $j = $#list;
        }
        $data[$j][$i] = 1;
    }
    my $tariffSet = Labelset( name => 'Tariffs for charging', list => \@list );
    $self->{matrix} = Constant(
        name          => 'Routed tariff to charged tariff mapping',
        defaultFormat => '0con',
        rows          => $self->{customers}->tariffSet,
        cols          => $tariffSet,
        data          => \@data,
    );
    $self->{tariffSet} = $tariffSet;
}

sub matrix {
    my ($self) = @_;
    $self->tariffSet unless defined $self->{matrix};
    $self->{matrix};
}

sub totalDemand {
    my ( $self, $usetName ) = @_;
    return $self->{totalDemand}{$usetName} if $self->{totalDemand}{$usetName};
    my @columns =
      map {
        Arithmetic(
            name          => "Total $_->{name} transposed",
            rows          => $self->tariffSet,
            cols          => 0,
            defaultFormat => '0copy',
            arithmetic    => '=A1',
            arguments     => {
                A1 => SumProduct(
                    name          => "Total $_->{name}",
                    matrix        => $self->{matrix},
                    vector        => $_,
                    usetName      => $usetName,
                    defaultFormat => '0soft',
                ),
            },
        );
      } @{ $self->{customers}->totalDemand($usetName) };
    push @{ $self->{model}{volumeTables} },
      Columnset(
        name    => "Forecast grouped volumes for $usetName",
        columns => \@columns,
      );
    $self->{totalDemand}{$usetName} = \@columns;
}

sub numberOfRoutes {
    my ($self) = @_;
    return $self->{numberOfRoutes} ||=  GroupBy(
                name          => 'Number of routes for each tariff',
                rows          => 0,
                cols          => $self->{matrix}{cols},
                source        => $self->{matrix},
                defaultFormat => '0soft',
            );
}

sub finish { }

1;
