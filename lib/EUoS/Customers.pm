﻿package EUoS::Customers;

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
use SpreadsheetModel::Shortcuts ':all';

sub new {
    my ( $class, $model, $setup ) = @_;
    bless { model => $model, setup => $setup }, $class;
}

sub totalDemand {
    my ( $self, $usetName ) = @_;
    return $self->{totalDemand}{$usetName} if $self->{totalDemand}{$usetName};
    my $tariffSet       = $self->tariffSet;
    my $userLabelset    = $self->userLabelset;
    my $detailedVolumes = $self->detailedVolumes;
    push @{ $self->{scenarioProportions} }, my $prop = Dataset(
        name          => "Proportion in $usetName",
        rows          => $userLabelset,
        defaultFormat => '%hardnz',
        data          => [ map { 1; } @{ $userLabelset->{list} } ]
    );
    my $columns = [
        map {
            SumProduct(
                name          => $_->{name},
                matrix        => $prop,
                vector        => $_,
                rows          => $tariffSet,
                usetName      => $usetName,
                defaultFormat => '0softnz',
            );
        } @$detailedVolumes
    ];
    push @{ $self->{model}{volumeTables} },
      Columnset(
        name    => "Forecast volume for $usetName",
        columns => $columns,
      );
    $self->{totalDemand}{$usetName} = $columns;
}

sub individualDemand {
    my ( $self, $usetName ) = @_;
    return $self->{individualDemand}{$usetName}
      if $self->{individualDemand}{$usetName};
    my $spcol   = $self->totalDemand($usetName);
    my $columns = [
        map {
            Arithmetic(
                name          => $_->{name},
                arguments     => { IV1 => $_->{matrix}, IV2 => $_->{vector}, },
                arithmetic    => '=IV1*IV2',
                defaultFormat => '0soft',
            );
        } @$spcol
    ];
    push @{ $self->{model}{volumeTables} },
      Columnset(
        name    => "Individual customer volumes in $usetName",
        columns => $columns,
      );
    $self->{individualDemand}{$usetName} = $columns;
}

sub userLabelset {
    my ($self) = @_;
    $self->{userLabelset} ||= Labelset(
        name   => 'Detailed list of customers',
        groups => [
            map { Labelset( name => keys %$_, list => values %$_ ); }
              @{ $self->{model}{ulist} }
        ]
    );
}

sub detailedVolumes {
    my ($self) = @_;
    return $self->{detailedVolumes} if $self->{detailedVolumes};
    my $userLabelset = $self->userLabelset;
    $self->{detailedVolumes} = [
        map {
            Dataset(
                rows          => $userLabelset,
                defaultFormat => '0hard',
                name          => $_,
                data          => [ map { 0 } @{ $userLabelset->{list} } ],
            );
        } @{ $self->{setup}->volumeComponents }
    ];
    return $self->{detailedVolumes};
}

sub tariffSet {
    my ($self) = @_;
    $self->{tariffSet} ||= Labelset(
        name => 'Set of customer categories',
        list => $self->userLabelset->{groups},
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
        name     => 'Definition of user sets',
        number   => 1514,
        appendTo => $self->{model}{inputTables},
        dataset  => $self->{model}{dataset},
        columns  => $self->{scenarioProportions},
    ) if $self->{scenarioProportions};
}

1;
