package Elec::Charging;

=head Copyright licence and disclaimer

Copyright 2012-2016 Franck Latrémolière, Reckon LLP and others.

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
    my ( $class, $model, $setup, $usage ) = @_;
    $model->register(
        bless {
            model => $model,
            setup => $setup,
            usage => $usage,
        },
        $class
    );
}

sub energyCharge {
    my ($self) = @_;
    $self->{energyCharge} ||= Arithmetic(
        name      => 'Energy charging rate £/kW/year',
        arguments => {
            A2 => $self->{setup}->daysInYear,
            A1 => Dataset(
                name     => 'Energy charging rate p/kWh',
                cols     => $self->{setup}->energyUsageSet,
                number   => 1585,
                appendTo => $self->{model}{inputTables},
                dataset  => $self->{model}{dataset},
                data     => [10],
            ),
        },
        arithmetic => '=A1*0.01*A2*24',
    );
}

sub assetRate {
    my ($self) = @_;
    $self->{assetRate} ||= Dataset(
        name          => 'Notional asset rates (£/kVA or £/point)',
        defaultFormat => '0hard',
        number        => 1550,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        cols          => $self->{setup}->usageSet,
        data =>
          [ 0, ( map { 1 } 3 .. @{ $self->{setup}->usageSet->{list} } ), 0 ],
    );
}

sub contributionDiscount {
    my ($self) = @_;
    $self->{contributionDiscount} ||= Dataset(
        name          => 'Contribution-related discount factors',
        defaultFormat => '%hard',
        number        => 1555,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        cols          => $self->{setup}->usageSet,
        data          => [ map { 0 } @{ $self->{setup}->usageSet->{list} } ],
    );
}

sub boundaryCharge {
    my ($self) = @_;
    $self->{boundaryCharge} ||= Dataset(
        name     => 'Boundary charging rate (£/year/unit of usage)',
        number   => 1552,
        appendTo => $self->{model}{inputTables},
        dataset  => $self->{model}{dataset},
        cols     => $self->{setup}->boundaryUsageSet,
        data => [ map { 10; } @{ $self->{setup}->boundaryUsageSet->{list} } ],
    );
}

sub annuityRate {
    my ($self) = @_;
    $self->{annuityRate} ||= $self->{setup}->annuityRate;
}

sub runningRate {
    my ($self) = @_;
    $self->{runningRate} ||= Dataset(
        name          => 'Annual running costs (relative to notional assets)',
        number        => 1554,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        data          => [0.02],
        defaultFormat => '%hard',
    );
}

sub assetCharge {
    my ($self) = @_;
    $self->{assetCharge} ||= Arithmetic(
        name       => 'Asset-related charging rate (£/unit of usage)',
        arithmetic => '=A1*('
          . ( $self->{model}{contributions} ? '(1-A4)*' : '' )
          . 'A2+A3)',
        arguments => {
            A1 => $self->assetRate,
            A2 => $self->annuityRate,
            A3 => $self->runningRate,
            $self->{model}{contributions}
            ? ( A4 => $self->contributionDiscount )
            : (),
        }
    );
}

sub usetBoundaryCosts {
    my ( $self, $totalUsage ) = @_;
    $self->{boundaryCharge} ||= Arithmetic(
        name       => 'Boundary charging rate (£/unit of usage/year)',
        cols       => $self->{setup}->boundaryUsageSet,
        arithmetic => '=A1/A2',
        arguments  => {
            A1 => Dataset(
                name          => 'Relevant boundary charges (£/year)',
                number        => 1556,
                appendTo      => $self->{model}{inputTables},
                dataset       => $self->{model}{dataset},
                data          => [5e5],
                defaultFormat => '0hard',
            ),
            A2 => $totalUsage,
        }
    );
}

sub detailedAssets {
    my ( $self, $usage, %flags ) = @_;
    my $notionalAssetMatrix = Arithmetic(
        name          => 'Notional asset matrix (£)',
        defaultFormat => '0soft',
        arithmetic    => '=A1*A2',
        arguments     => {
            A1 => $usage,
            A2 => $self->assetRate,
        },
    );
    my $notionalAssetsByUser = GroupBy(
        name          => 'Notional assets (£)',
        rows          => $usage->{rows},
        source        => $notionalAssetMatrix,
        defaultFormat => '0soft',
    );
    push @{ $self->{model}{detailedTablesBottom} },
      Columnset(
        name => 'Notional assets by ' . ( $flags{userName} || 'user' ),
        columns => [
            $usage->{names} ? Stack( sources => [ $usage->{names} ] ) : (),
            $notionalAssetMatrix, $notionalAssetsByUser,
        ]
      );
    push @{ $self->{model}{detailedTablesBottom} },
      Columnset(
        name    => 'Total notional assets',
        columns => [
            $usage->{names} ? Stack( sources => [ $usage->{names} ] ) : (),
            GroupBy(
                name          => 'Total notional assets (£)',
                defaultFormat => '0soft',
                cols          => $usage->{cols},
                source        => $notionalAssetMatrix,
            ),
            GroupBy(
                name          => 'Grand total notional assets (£)',
                defaultFormat => '0soft',
                source        => $notionalAssetsByUser,
            ),
        ]
      ) if $flags{showTotals};
}

sub usetMatchAssets {
    my ( $self, $totalUsage, $doNotApply ) = @_;
    my $beforeMatching = $self->assetRate;
    my $totalBefore    = SumProduct(
        name          => 'Total notional assets before matching (£)',
        matrix        => $totalUsage,
        vector        => $beforeMatching,
        defaultFormat => '0soft',
    );
    my $maxAssets = Dataset(
        name          => 'Maximum total notional asset value (£)',
        number        => 1558,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        data          => [1e7],
        defaultFormat => '0hard',
    );
    if ($doNotApply) {
        push @{ $self->{model}{checkTables} },
          Columnset(
            name => 'For information: comparison of calculated '
              . 'and target gross modern equivalent asset value (£)',
            columns => [
                $totalBefore,
                Stack( sources => [$maxAssets] ),
                Arithmetic(
                    name       => 'Ratio',
                    arithmetic => '=A1/A2',
                    arguments  => { A1 => $totalBefore, A2 => $maxAssets },
                )
            ]
          );
    }
    else {
        $self->{assetRate} = Arithmetic(
            name => 'Adjusted notional assets for each type of usage'
              . ' (£/kVA or £/point)',
            arithmetic => '=A1*MIN(1,A2/A3)',
            arguments =>
              { A1 => $beforeMatching, A2 => $maxAssets, A3 => $totalBefore, },
        );
    }
}

sub usetRunningCosts {
    my ( $self, $totalUsage ) = @_;
    my $assetRate   = $self->assetRate;
    my $totalAssets = SumProduct(
        name          => 'Total relevant notional assets (£)',
        matrix        => $totalUsage,
        vector        => $assetRate,
        defaultFormat => '0soft',
    );
    my $target = Dataset(
        name          => 'Total running costs (£/year)',
        number        => 1559,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        data          => [1e6],
        defaultFormat => '0hard',
    );
    $self->{runningRate} = Arithmetic(
        name          => 'Annual running costs (relative to notional assets)',
        arithmetic    => '=A1/A3',
        defaultFormat => '%soft',
        arguments     => { A1 => $target, A3 => $totalAssets, },
    );
}

sub usetMatchRevenue { }

sub charges {
    my ($self) = @_;
    (
        $self->boundaryCharge, $self->assetCharge,
        $self->{model}{noEnergy} ? () : $self->energyCharge,
    );
}

sub finish { }

1;
