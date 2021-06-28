package Elec::Charging;

# Copyright 2012-2021 Franck Latrémolière and others.
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
    my ( $class, $model, $setup, $usage, $assets ) = @_;
    $model->register(
        bless {
            model  => $model,
            setup  => $setup,
            usage  => $usage,
            assets => $assets,
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
                name => 'Cost of procuring energy for electrical losses p/kWh',
                cols => $self->{setup}->energyUsageSet,
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
    $self->{assetRate} ||=
      $self->{assets} ? $self->{assets}->assetRate : Dataset(
        name          => 'Notional asset rates (£/unit of usage)',
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

sub nonAssetCharge {
    my ($self) = @_;
    $self->{nonAssetCharge} ||= Dataset(
        name     => 'Non asset charges (£/year/unit of usage)',
        number   => 1552,
        appendTo => $self->{model}{inputTables},
        dataset  => $self->{model}{dataset},
        cols     => $self->{setup}->nonAssetUsageSet,
        data => [ map { 10; } @{ $self->{setup}->nonAssetUsageSet->{list} } ],
    );
}

sub annuityRate {
    my ($self) = @_;
    $self->{annuityRate} ||= Arithmetic(
        name          => 'Annuity rate',
        defaultFormat => '%soft',
        arithmetic    => '=PMT(A1,A2,-1)',
        arguments     => {
            A1 => $self->{setup}->rateOfReturn,
            A2 => $self->{setup}->annuitisationPeriod,
        }
    );
}

sub runningRate {
    my ($self) = @_;
    $self->{runningRate} ||= Dataset(
        name          => 'Annual running costs relative to notional assets',
        number        => 1554,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        data          => [0.02],
        defaultFormat => '%hard',
    );
}

sub assetCharge {
    my ($self) = @_;
    $self->{assetCharge} ||=
      $self->{assets} && $self->{assets}->assetLives
      ? Arithmetic(
        name       => 'Asset-related charges (£/unit of usage)',
        arithmetic => '='
          . ( $self->{model}{contributions} ? '(1-A4)*' : '' )
          . ( $self->{assetMatchingFactor}  ? 'A5*'     : '' )
          . 'A2+A1*A3',
        arguments => {
            A1 => $self->assetRate,
            A2 => $self->{assets}->annuity( $self->{setup}->rateOfReturn ),
            A3 => $self->runningRate,
            $self->{assetMatchingFactor}
            ? ( A5 => $self->{assetMatchingFactor} )
            : (),
            $self->{model}{contributions}
            ? ( A4 => $self->contributionDiscount )
            : (),
        }
      )
      : Arithmetic(
        name       => 'Asset-related charges (£/unit of usage)',
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

sub usetMatchAssetDetail {
    my ( $self, $totalUsage, $applicationOptions ) = @_;
    my $totalBefore = SumProduct(
        name   => 'Detailed notional assets before matching',
        vector => Arithmetic(
            name => 'Number of notional schemes implied by volume forecast',
            arithmetic => '=A1/A2',
            arguments  => {
                A1 => $totalUsage,
                A2 => $self->{assets}->notionalCapacity,
            },
        ),
        matrix        => $self->{assets}->notionalVolumes,
        defaultFormat => '0soft',
    );
    my $assetVolumes =
      $self->{model}{usetMatchAssets}
      ? Stack( sources => [ $self->{assets}->assetVolumes ] )
      : $self->{assets}->assetVolumes;
    my $assetMatchingFactors =
      $applicationOptions =~ /top ([0-9]+)/ ? Arithmetic(
        name => 'Detailed notional asset adjustment factors'
          . $applicationOptions,
        arithmetic => '=MIN(A3,IF(A11,A2/A1,10))',
        arguments  => {
            A1  => $totalBefore,
            A11 => $totalBefore,
            A2  => $assetVolumes,
            A3  => Constant(
                name          => 'Asset adjustment factor cap',
                defaultFormat => '0.000con',
                rows          => $assetVolumes->{rows},
                data          => [
                    ( map { ''; } 0 .. ( $1 - 1 ) ),
                    ( map { 1; } $1 .. $#{ $assetVolumes->{rows}{list} } ),
                ],
            ),
        },
      ) : Arithmetic(
        name => 'Detailed notional asset adjustment factors'
          . $applicationOptions,
        arithmetic => $applicationOptions =~ /cap([0-9.]+)/i
        ? "=MIN($1,IF(A11,A2/A1,10))"
        : $applicationOptions =~ /cap/i ? '=MIN(1,IF(A11,A2/A1,10))'
        : '=IF(A11,A2/A1,10)',
        arguments => {
            A1  => $totalBefore,
            A11 => $totalBefore,
            A2  => $assetVolumes,
        },
      );
    if ( $applicationOptions =~ /info/i ) {
        $self->{assets}->addNotionalVolumesFeedback(
            $assetMatchingFactors,
            conditionalFormatting => {
                type      => '3_color_scale',
                min_type  => 'num',
                mid_type  => 'num',
                max_type  => 'num',
                min_value => sqrt(.5),
                mid_value => 1,
                max_value => sqrt(2),
                min_color => '#ccccff',
                mid_color => '#ccffcc',
                max_color => '#ffcccc',
            },
        );
        push @{ $self->{model}{checkTables} },
          Columnset(
            name => 'Matching detailed notional assets to actual assets'
              . $applicationOptions,
            columns => [ $totalBefore, $assetVolumes, $assetMatchingFactors, ],
          );
    }
    else {
        Columnset(
            name    => 'Calculation of asset adjustment factors',
            columns => [
                grep { defined $_; } $assetMatchingFactors->{arguments}{A3},
                $assetMatchingFactors->{arguments}{A1},
                $assetMatchingFactors->{arguments}{A2},
                $assetMatchingFactors,
            ],
        );
        $self->{assets}->notionalVolumes(
            Arithmetic(
                name => 'Adjusted detailed notional assets'
                  . ' for each type of usage',
                arithmetic => '=A1*A2',
                arguments  => {
                    A1 => $self->{assets}->notionalVolumes,
                    A2 => $assetMatchingFactors,
                },
            )
        );
    }
}

sub usetMatchAssets {
    my ( $self, $totalUsage, $applicationOptions ) = @_;
    my $beforeMatching = $self->assetRate;
    my $totalBefore    = SumProduct(
        name          => 'Total notional assets before matching (£)',
        matrix        => $totalUsage,
        vector        => $beforeMatching,
        defaultFormat => '0soft',
    );
    my $maxAssets = $self->{assets} ? $self->{assets}->maxAssets : Dataset(
        name          => 'Maximum total notional asset value (£)',
        number        => 1558,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        data          => [1e7],
        defaultFormat => '0hard',
    );
    my $assetMatchingFactor = Arithmetic(
        name       => 'Asset adjustment factor' . $applicationOptions,
        arithmetic => '=MIN(1,A2/A3)',
        arguments  => { A2 => $maxAssets, A3 => $totalBefore, },
    );
    my $columnset = Columnset(
        name => 'Application of maximum total notional asset value'
          . $applicationOptions,
        columns => [
            $totalBefore, $self->{assets} ? $maxAssets : (),
            $assetMatchingFactor,
        ],
    );
    if ( $applicationOptions =~ /info/i ) {
        push @{ $self->{model}{checkTables} }, $columnset;
    }
    else {
        $self->{assetRate} = Arithmetic(
            name => 'Adjusted notional assets for each type of usage'
              . ' (£/unit of usage)',
            arithmetic => '=A1*A2',
            arguments  => {
                A1 => $beforeMatching,
                A2 => $assetMatchingFactor,
            },
        );
    }
}

sub usetRunningCosts {
    my ( $self, $totalUsage ) = @_;
    my $totalCosts =
        $self->{model}{interpolator}
      ? $self->{model}{interpolator}
      ->runningCosts( Labelset( list => ['Asset running costs'] ) )
      : Dataset(
        name          => 'Total asset running costs (£/year)',
        number        => 1559,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        data          => [1e6],
        defaultFormat => '0hard',
      );
    my $totalAssets = SumProduct(
        name          => 'Total relevant notional assets (£)',
        matrix        => $totalUsage,
        vector        => $self->assetRate,
        defaultFormat => '0soft',
    );
    $self->{runningRate} = Arithmetic(
        name       => 'Annual running costs relative to notional asset value',
        arithmetic => '=A1/A3',
        defaultFormat => '%soft',
        arguments     => { A1 => $totalCosts, A3 => $totalAssets, },
    );
    Columnset(
        name    => 'Calculation of running costs charging rate',
        columns => [
            $self->{model}{interpolator} ? $totalCosts : (), $totalAssets,
            $self->{runningRate},
        ],
    );
}

sub usetNonAssetCosts {
    my ( $self, $totalUsage ) = @_;
    my $nonAssetCosts =
        $self->{model}{interpolator}
      ? $self->{model}{interpolator}
      ->runningCosts( $self->{setup}->nonAssetUsageSet )
      : Dataset(
        name     => 'Relevant non-asset charges (£/year)',
        number   => 1556,
        cols     => $self->{setup}->nonAssetUsageSet,
        appendTo => $self->{model}{inputTables},
        dataset  => $self->{model}{dataset},
        data => [ map { 5e5; } @{ $self->{setup}->nonAssetUsageSet->{list} } ],
        defaultFormat => '0hard',
      );
    $self->{nonAssetCharge} = Arithmetic(
        name       => 'Non-asset-based charges (£/unit of usage/year)',
        arithmetic => '=A1/A2',
        arguments  => {
            A1 => $nonAssetCosts,
            A2 => $totalUsage,
        }
    );
}

sub usetMatchRevenue {
    die 'Not implemented';
}

sub charges {
    my ($self) = @_;
    $self->nonAssetCharge, $self->assetCharge,
      $self->{model}{noEnergy} ? () : $self->energyCharge;
}

sub finish { }

1;
