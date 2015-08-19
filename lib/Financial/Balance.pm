package Financial::Balance;

=head Copyright licence and disclaimer

Copyright 2015 Franck Latrémolière, Reckon LLP and others.

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
    my ( $class, %hash ) = @_;
    $hash{$_} || die __PACKAGE__ . " needs a $_ attribute"
      foreach qw(model assets sales expenses debt cashCalc);
    bless \%hash, $class;
}

sub statement {
    my ( $balance, $periods ) = @_;
    $balance->{statement}{ 0 + $periods } ||= CalcBlock(
        name =>
          $periods->decorate( 'Balance sheet' . ( $balance->{suffix} || '' ) ),
        items => [
            [
                [
                    $balance->{assets}->netValue($periods),
                    [
                        $balance->{sales}->balance($periods),
                        A5 => $balance->{cashCalc}
                          ->total( $periods, $balance->{reserve} ),
                        $periods->decorate('Current assets (£)'),
                    ],
                    A1 => $periods->decorate('Total assets (£)'),
                ],
                $balance->{expenses}->balance($periods),
                $periods->decorate('Total assets less current liabilities (£)'),
            ],
            $balance->{debt}->due($periods),
            A2 => $periods->decorate(
                'Equity' . ( $balance->{suffix} || '' ) . ' (£)'
            ),
        ],
    );
}

sub equity {
    my ( $balance, $periods ) = @_;
    $balance->statement($periods)->{A2};
}

sub cash {
    my ( $balance, $periods ) = @_;
    $balance->statement($periods)->{A5};
}

sub workingCapital {
    my ( $balance, $periods ) = @_;
    $balance->{workingCapital}{ 0 + $periods } ||= CalcBlock(
        name  => $periods->decorate('Working capital analysis'),
        items => [
            $balance->{sales}->balance($periods),
            $balance->{expenses}->balance($periods),
            $balance->{cashCalc}->total( $periods, $balance->{reserve} ),
            A1 => $periods->decorate('Working capital (£)'),
        ],
    )->{A1};
}

sub fixedAssetAnalysis {
    my ( $balance, $periods ) = @_;
    $balance->{fixedAssetAnalysis}{ 0 + $periods } ||= CalcBlock(
        name  => $periods->decorate('Fixed asset analysis'),
        items => [
            A1 => $balance->{assets}->grossValue($periods),
            A2 => $balance->{assets}->netValue($periods),
            {
                name => $periods->decorate('Accumulated depreciation (£)'),
                defaultFormat => '0boldsoft',
                arithmetic    => '=A1-A2',
            },
        ],
    );
}

sub initialEquity {
    my ( $balance, $periods ) = @_;
    return $balance->{initialEquity}{ 0 + $periods }
      if $balance->{initialEquity}{ 0 + $periods };
    $balance->{initialEquity}{ 0 + $periods } = Arithmetic(
        name          => 'Initial equity (£)',
        defaultFormat => '0soft',
        arithmetic    => '=INDEX(A31_A32,'
          . ( $periods->{reverseTime} ? @{ $periods->{list} } : 1 ) . ')',
        arguments => {
            A31_A32 => $balance->equity($periods),
        },
    );
}

sub shareCapital {
    my ( $balance, $periods ) = @_;
    return $balance->equity($periods) unless $balance->{reserve};
    $balance->{shareCapital}{ 0 + $periods } ||= Arithmetic(
        name =>
          $periods->decorate('Share capital and pre-existing reserves (£)'),
        defaultFormat => '0soft',
        arithmetic    => '=A6+SUMIF(A2_A3,"<="&A1,A4_A5)',
        arguments     => {
            A1    => $periods->lastDay,
            A2_A3 => $balance->{reserve}->raisingDates,
            A4_A5 => $balance->{reserve}->amountsRaised,
            A6    => $balance->initial,
        },
    );
}

1;
