package EUoS::Charging;

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
    my ( $class, $model, $setup, $usage ) = @_;
    bless { model => $model, setup => $setup, usage => $usage }, $class;
}

sub assetRate {
    my ($self) = @_;
    return $self->{assetRate} if $self->{assetRate};
    my $usageSet = $self->{usage}->usageSet;
    $self->{assetRate} = Dataset(
        name  => 'Notional asset rates (£/kVA or £/point)',
        lines => [
'This table is input data but needs to have a backing sheet explaining the assumptions underpinning the numbers.'
        ],
        defaultFormat => '0hardnz',
        number        => 1550,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        cols          => $usageSet,
        data          => [ map { 1 } @{ $usageSet->{list} } ],
    );
}

sub boundaryCharge {
    my ($self) = @_;
    return $self->{boundaryCharge} if $self->{boundaryCharge};
    $self->{boundaryCharge} = Dataset(
        name     => 'Boundary charging rate (£/kVA/year)',
        number   => 1552,
        appendTo => $self->{model}{inputTables},
        dataset  => $self->{model}{dataset},
        cols     => $self->{usage}->boundaryUsageSet,
        data     => [10],
    );
}

sub annuityRate {
    my ($self) = @_;
    return $self->{annuityRate} if $self->{annuityRate};
    $self->{annuityRate} = $self->{setup}->annuityRate;
}

sub runningRate {
    my ($self) = @_;
    return $self->{runningRate} if $self->{runningRate};
    $self->{runningRate} = Dataset(
        name          => 'Annual running costs (relative to notional assets)',
        number        => 1554,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        data          => [0.02],
        defaultFormat => '%hardnz',
    );
}

sub assetCharge {
    my ($self) = @_;
    return $self->{assetCharge} if $self->{assetCharge};
    $self->{assetCharge} = Arithmetic(
        name       => 'Asset-related charging rate (£/kVA or point/year)',
        arithmetic => '=IV1*(IV2+IV3)',
        arguments  => {
            IV1 => $self->assetRate,
            IV2 => $self->annuityRate,
            IV3 => $self->runningRate,
        }
    );
}

sub matchBoundary {
    my ( $self, $totalUsage ) = @_;
    my $boundaryUsageSet = $self->{usage}->boundaryUsageSet;
    my $charges          = Dataset(
        name          => 'Relevant boundary charges (£/year)',
        number        => 1556,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        data          => [5e5],
        defaultFormat => '0hardnz',
    );
    $self->{boundaryCharge} = Arithmetic(
        name       => 'Boundary charging rate (£/kVA/year)',
        cols       => $boundaryUsageSet,
        arithmetic => '=IV1/IV2',
        arguments  => { IV1 => $charges, IV2 => $totalUsage, }
    );
}

sub matchAssets {
    my ( $self, $totalUsage, $doNotApply ) = @_;
    my $beforeMatching = $self->assetRate;
    my $totalBefore    = SumProduct(
        name          => 'Total notional assets before matching (£)',
        matrix        => $totalUsage,
        vector        => $beforeMatching,
        defaultFormat => '0softnz',
    );
    my $target = Dataset(
        name          => 'Target total notional assets (£)',
        number        => 1558,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        data          => [1e7],
        defaultFormat => '0hardnz',
    );
    if ($doNotApply) {
        push @{ $self->{model}{usageTables} },
          Columnset(
            name => 'For information: comparison of calculated '
              . 'and target gross modern equivalent asset value (£)',
            columns => [
                $totalBefore,
                Stack( sources => [$target] ),
                Arithmetic(
                    name       => 'Ratio',
                    arithmetic => '=IV1/IV2',
                    arguments  => { IV1 => $totalBefore, IV2 => $target },
                )
            ]
          );
    }
    else {
        $self->{assetRate} = Arithmetic(
            name => 'Adjusted notional assets for each type of usage'
              . ' (£/KVA or £/point)',
            arithmetic => '=IV1*IV2/IV3',
            arguments =>
              { IV1 => $beforeMatching, IV2 => $target, IV3 => $totalBefore, },
        );
    }
}

sub matchRunning {
    my ( $self, $totalUsage ) = @_;
    my $assetRate   = $self->assetRate;
    my $totalAssets = SumProduct(
        name          => 'Total relevant notional assets (£)',
        matrix        => $totalUsage,
        vector        => $assetRate,
        defaultFormat => '0softnz',
    );
    my $target = Dataset(
        name          => 'Target total running costs (£)',
        number        => 1559,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        data          => [6e5],
        defaultFormat => '0hardnz',
    );
    $self->{runningRate} = Arithmetic(
        name          => 'Annual running costs (relative to notional assets)',
        arithmetic    => '=IV1/IV3',
        defaultFormat => '%softnz',
        arguments     => { IV1 => $target, IV3 => $totalAssets, },
    );
}

sub matchRevenue { }

sub charges {
    my ($self) = @_;
    ( $self->boundaryCharge, $self->assetCharge );
}

sub finish { }

1;
