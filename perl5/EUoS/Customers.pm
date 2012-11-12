package EUoS::Customers;

=head Copyright licence and disclaimer

Copyright 2012 Reckon LLP and contributors. All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice,
this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation
and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY RECKON LLP AND CONTRIBUTORS "AS IS" AND ANY
EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL RECKON LLP OR CONTRIBUTORS BE LIABLE FOR ANY
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

sub new {
    my ( $class, $model, $setup ) = @_;
    bless { model => $model, setup => $setup }, $class;
}

sub defaultData {
    my ( $tariffSet, $volumeComponent ) = @_;
    [ map { 0 } @{ $tariffSet->{list} } ];
}

sub scenario {
    my ( $self, $filter, $name ) = @_;
    my $tariffSet = $self->tariffSet;
    unless ($filter) {
        return $self->{aggregatedVolumes} if $self->{aggregatedVolumes};
        $self->{aggregatedVolumes} = [
            map {
                Dataset(
                    rows => $tariffSet,
                    name => $_,
                    data => defaultData( $tariffSet, $_ )
                );
            } @{ $self->{setup}->volumeComponents }
        ];
        Columnset(
            name          => 'Aggregated forecast volumes',
            number        => 1513,
            appendTo      => $self->{model}{inputTables},
            dataset       => $self->{model}{dataset},
            columns       => $self->{aggregatedVolumes},
            defaultFormat => '0hard',
        );
        return $self->{aggregatedVolumes};
    }
    my $customerSet     = $self->customerSet;
    my $detailedVolumes = $self->detailedVolumes;
    push @{ $self->{coefficients} }, my $coefficients = Dataset(
        name          => "Proportion included in $name",
        rows          => $customerSet,
        defaultFormat => '%hardnz',
        data          => [ map { $filter->($_) } @{ $customerSet->{list} } ]
    );
    my $columns = [
        map {
            SumProduct(
                name          => $_->{name},
                matrix        => $coefficients,
                vector        => $_,
                rows          => $tariffSet,
                scenario      => $name,
                defaultFormat => '0softnz',
            );
        } @$detailedVolumes
    ];
    push @{ $self->{model}{volumeTables} },
      Columnset(
        name    => "Forecast volume for $name",
        columns => $columns,
      );
    $columns;
}

sub customerSet {
    my ($self) = @_;
    $self->{customerSet} ||= Labelset(
        name   => 'Detailed list of customers',
        groups => [
            map { Labelset( name => keys %$_, list => values %$_ ); }
              @{ $self->{model}{customerList} }
        ]
    );
}

sub detailedVolumes {
    my ($self) = @_;
    return $self->{detailedVolumes} if $self->{detailedVolumes};
    my $customerSet = $self->customerSet;
    $self->{detailedVolumes} = [
        map {
            Dataset(
                rows          => $customerSet,
                defaultFormat => '0hard',
                name          => $_,
                data          => defaultData( $customerSet, $_ ),
            );
        } @{ $self->{setup}->volumeComponents }
    ];
    return $self->{detailedVolumes};
}

sub tariffSet {
    my ($self) = @_;
    $self->{tariffSet} ||= Labelset(
        name => 'Set of customer categories',
        list => $self->customerSet->{groups},
    );
}

sub finish {
    my ($self) = @_;
    Columnset(
        name     => 'Forecast volumes',
        number   => 1512,
        appendTo => $self->{model}{inputTables},
        dataset  => $self->{model}{dataset},
        columns  => $self->{detailedVolumes},
    ) if $self->{detailedVolumes};
    Columnset(
        name     => 'Definition of scenarios',
        number   => 1514,
        appendTo => $self->{model}{inputTables},
        dataset  => $self->{model}{dataset},
        columns  => $self->{coefficients},
    ) if $self->{coefficients};
}

1;
