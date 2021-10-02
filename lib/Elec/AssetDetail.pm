package Elec::AssetDetail;

# Copyright 2019-2021 Franck Latrémolière and others.
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
    $self->{assetLabelset} ||= Labelset(
        name => 'Asset types',
        list => ref $self->{model}{assetDetail}
        ? $self->{model}{assetDetail}
        : [ split /\n/, <<EOL ], );
MV breaker
MV circuit km
MV transformer
MV ring main unit
LV frame
LV circuit km
LV cut-out
EOL
}

sub valuesVolumesLives {
    my ($self) = @_;
    return $self->{valuesVolumesLives} if $self->{valuesVolumesLives};
    if ( $self->{model}{interpolator} ) {
        my ( $val, $life ) = $self->{model}{interpolator}
          ->assetValuesLives( $self->assetLabelset );
        return $self->{valuesVolumesLives} = [
            $val,
            $self->{model}{interpolator}->assetVolumes( $self->assetLabelset ),
            $life
        ];
    }
    $self->{valuesVolumesLives} = [
        Dataset(
            name          => Label( 'Value (£)', 'Asset valuation (£/unit)' ),
            defaultFormat => '0hard',
            rows          => $self->assetLabelset,
            data => [ map { 1 } 1 .. @{ $self->assetLabelset->{list} } ],
        ),
        Dataset(
            name => Label( 'Volume', 'Asset volume (units)' ),
            rows => $self->assetLabelset,
            data => [ map { 1 } 1 .. @{ $self->assetLabelset->{list} } ],
        ),
        $self->{model}{assetDetailAnnualisationPeriod}
        ? Dataset(
            name => Label(
                'Annualisation period (years)',
                'Asset annualisation period (years)'
            ),
            rows => $self->assetLabelset,
            data => [ map { 1 } 1 .. @{ $self->assetLabelset->{list} } ],
          )
        : (),
    ];
    Columnset(
        name     => 'Network assets',
        number   => 1551,
        appendTo => $self->{model}{inputTables},
        dataset  => $self->{model}{dataset},
        columns  => $self->{valuesVolumesLives},
    );
    $self->{valuesVolumesLives};
}

sub assetValues {
    my ($self) = @_;
    $self->valuesVolumesLives->[0];
}

sub assetVolumes {
    my ($self) = @_;
    $self->valuesVolumesLives->[1];
}

sub assetLives {
    my ($self) = @_;
    $self->valuesVolumesLives->[2];
}

sub notionalCapacity {
    my ($self) = @_;
    $self->{notionalCapacity} ||= Dataset(
        name          => 'Notional scheme capacities',
        defaultFormat => '0hard',
        number        => 1554,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        cols          => $self->{setup}->usageSet,
        data => [ map { 1 } 1 .. @{ $self->{setup}->usageSet->{list} } ],
    );
}

sub notionalVolumes {
    my ( $self, $setNotionalVolumes ) = @_;
    return $self->{notionalVolumes} = $setNotionalVolumes
      if $setNotionalVolumes;
    $self->{notionalVolumes} ||= $self->{notionalVolumesInput} ||= Dataset(
        name    => 'Notional scheme asset volumes',
        number  => 1553,
        dataset => $self->{model}{dataset},
        rows    => $self->assetLabelset,
        cols    => $self->{setup}->usageSet,
        data    => [ map { 1 } 1 .. @{ $self->{setup}->usageSet->{list} } ],
    );
}

sub assetRate {
    my ($self) = @_;
    $self->{assetRate} ||= Arithmetic(
        name       => 'Notional asset rate (£/unit of usage)',
        arithmetic => '=IF(A3,A1/A2,0)',
        arguments  => {
            A1 => SumProduct(
                name          => 'Notional scheme asset valuation (£)',
                defaultFormat => '0soft',
                matrix        => $self->notionalVolumes,
                vector        => $self->assetValues,
            ),
            A2 => $self->notionalCapacity,
            A3 => $self->notionalCapacity,
        },
    );
}

sub annuity {
    my ( $self, $rateOfReturn, $defaultAssetLife ) = @_;
    $self->{annuity} ||= Arithmetic(
        name       => 'Notional asset annuity rate (£/unit of usage/year)',
        arithmetic => '=IF(A3,A1/A2,0)',
        arguments  => {
            A1 => SumProduct(
                name          => 'Notional scheme asset annuity (£/year)',
                defaultFormat => '0soft',
                matrix        => $self->notionalVolumes,
                vector        => Arithmetic(
                    name          => 'Asset annuity (£/year)',
                    defaultFormat => '0soft',
                    arithmetic    => '=-1*PMT(A2,'
                      . ( $defaultAssetLife ? 'IF(A3,A4,A5)' : 'A3' ) . ',A1)',
                    arguments => {
                        A1 => $self->assetValues,
                        A2 => $rateOfReturn,
                        A3 => $self->assetLives,
                        $defaultAssetLife
                        ? (
                            A4 => $self->assetLives,
                            A5 => $defaultAssetLife,
                          )
                        : (),
                    },
                ),
            ),
            A2 => $self->notionalCapacity,
            A3 => $self->notionalCapacity,
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

sub addNotionalVolumesRulesInput {
    my ( $self, $rule, ) = @_;
    push @{ $self->{additionalColumns1553} }, $rule;
}

sub addNotionalVolumesFeedback {
    my ( $self, $feedback, @more ) = @_;
    push @{ $self->{additionalColumns1553} },
      Stack(
        sources        => [$feedback],
        deferWritingTo => $feedback,
        @more,
      );
}

sub finish {
    my ($self) = @_;
    if ( $self->{additionalColumns1553} ) {
        Columnset(
            name     => $self->{notionalVolumesInput}->objectShortName,
            number   => 1553,
            dataset  => $self->{model}{dataset},
            appendTo => $self->{model}{inputTables},
            columns  => [
                $self->{notionalVolumesInput},
                @{ $self->{additionalColumns1553} },
            ],
        );
    }
    elsif ( $self->{notionalVolumesInput} ) {
        push @{ $self->{model}{inputTables} }, $self->{notionalVolumesInput};
    }
}

1;
