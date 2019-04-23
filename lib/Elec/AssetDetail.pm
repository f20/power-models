package Elec::AssetDetail;

# Copyright 2019 Franck Latrémolière, Reckon LLP and others.
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

sub new {
    my ( $class, $model, $setup, ) = @_;
    $model->register(
        bless {
            model => $model,
            setup => $setup,
        },
        $class
    );
}

sub assetLabelset {
    my ($self) = @_;
    $self->{assetLabelset} ||=
      Labelset( name => 'Asset types', list => [ split /\n/, <<EOL ], );
33kV underground non-pressurised km
33kV transformer ground-mounted
33kV breaker indoors
11kV breaker primary
11kV breaker secondary
11kV underground km
11kV transformer ground-mounted
11kV ring main unit
LV main underground km
LV plant unit
EOL
}

sub assetValuesVolumesColumnset {
    my ($self) = @_;
    $self->{assetValuesVolumesColumnset} ||= Columnset(
        name     => 'Network assets',
        number   => 1551,
        appendTo => $self->{model}{inputTables},
        dataset  => $self->{model}{dataset},
        columns  => [
            Dataset(
                name => Label( 'Value (£)', 'Asset valuation (£/unit)' ),
                defaultFormat => '0hard',
                rows          => $self->assetLabelset,
                data => [ map { 1 } 1 .. @{ $self->assetLabelset->{list} } ],
            ),
            Dataset(
                name          => Label( 'Volume', 'Asset volume (units)' ),
                rows          => $self->assetLabelset,
                data => [ map { 1 } 1 .. @{ $self->assetLabelset->{list} } ],
            ),
        ],
    );
}

sub assetValues {
    my ($self) = @_;
    $self->assetValuesVolumesColumnset->{columns}[0];
}

sub assetVolumes {
    my ($self) = @_;
    $self->assetValuesVolumesColumnset->{columns}[1];
}

sub assetRate {
    my ($self) = @_;
    return $self->{assetRate} if $self->{assetRate};
    my $notionalCapacity = Dataset(
        name          => 'Notional scheme capacities',
        defaultFormat => '0hard',
        number        => 1554,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        cols          => $self->{setup}->usageSet,
        data => [ map { 1 } 1 .. @{ $self->{setup}->usageSet->{list} } ],
    );
    my $notionalValue = SumProduct(
        name          => 'Notional scheme asset valuation (£)',
        defaultFormat => '0soft',
        matrix        => Dataset(
            name     => 'Notional scheme asset volumes',
            number   => 1553,
            appendTo => $self->{model}{inputTables},
            dataset  => $self->{model}{dataset},
            rows     => $self->assetLabelset,
            cols     => $self->{setup}->usageSet,
            data => [ map { 1 } 1 .. @{ $self->{setup}->usageSet->{list} } ],
        ),
        vector => $self->assetValues,
    );
    $self->{assetRate} = Arithmetic(
        name          => 'Notional asset rate (£/unit of usage)',
        defaultFormat => '0soft',
        arithmetic    => '=IF(A3,A1/A2,0)',
        arguments     => {
            A1 => $notionalValue,
            A2 => $notionalCapacity,
            A3 => $notionalCapacity,
        },
    );
}

sub maxAssets {
    my ($self) = @_;
    $self->{maxAssets} ||= SumProduct(
        name          => 'Maximum network asset valuation (£)',
        defaultFormat => '0soft',
        matrix        => $self->assetVolumes,
        vector        => $self->assetValues,
    );
}

sub finish { }

1;
