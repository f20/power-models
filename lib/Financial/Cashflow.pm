package Financial::Cashflow;

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
      foreach qw(model income balance);
    bless \%hash, $class;
}

sub statement {
    my ( $cashflow, $periods ) = @_;
    $cashflow->{statement}{ 0 + $periods } ||= CalcBlock(
        name => $periods->decorate(
            'Cashflow statement' . ( $cashflow->{suffix} || '' )
        ),
        items => [
            [
                [
                    $cashflow->{income}->ebitda($periods),
                    A11 => Arithmetic(
                        name => $periods->decorate(
                            'Cash released/absorbed in debtors (£)'),
                        defaultFormat => '0soft',
                        arithmetic    => '=INDEX(A5_A6,A2)-A1',
                        arguments     => {
                            A1 =>
                              $cashflow->{balance}{sales}->balance($periods),
                            A5_A6 =>
                              $cashflow->{balance}{sales}->balance($periods),
                            A2 => $periods->indexPrevious,
                        },
                    ),
                    A12 => Arithmetic(
                        name => $periods->decorate(
                            'Cash released/absorbed in creditors (£)'),
                        defaultFormat => '0soft',
                        arithmetic => '=INDEX(A5_A6,A2)-A1+INDEX(A7_A8,A3)-A4',
                        arguments  => {
                            A1 => $cashflow->{balance}{costSales}
                              ->balance($periods),
                            A5_A6 => $cashflow->{balance}{costSales}
                              ->balance($periods),
                            A4 =>
                              $cashflow->{balance}{adminExp}->balance($periods),
                            A7_A8 =>
                              $cashflow->{balance}{adminExp}->balance($periods),
                            A2 => $periods->indexPrevious,
                            A3 => $periods->indexPrevious,
                        },
                    ),
                    $cashflow->{income}->tax($periods),
                    A5 => $periods->decorate(
                        'Cashflow from operating activities (£)'),
                ],
                [
                    $cashflow->{balance}{assets}->capitalExpenditure($periods),
                    $cashflow->{balance}{assets}->capitalReceipts($periods),
                    A6 => $periods->decorate(
                        'Cashflow from investing activities (£)'),
                ],
                [
                    $cashflow->{income}{debt}->interest($periods),
                    $cashflow->{balance}{debt}->raised($periods),
                    $cashflow->{balance}{debt}->repaid($periods),
                    $cashflow->{balance}{reserve}
                    ? $cashflow->{balance}{reserve}->raised($periods)
                    : (),
                    $periods->decorate(
                        'Cashflow from financing activities (£)'),
                ],
                $periods->decorate(
                        'Cashflow from operating,'
                      . ' investing and financing activities (£)'
                ),
            ],
            A13 => Arithmetic(
                name => $periods->decorate('Cash retained/released (£)'),
                defaultFormat => '0soft',
                arithmetic    => '=INDEX(A5_A6,A2)-A1',
                arguments     => {
                    A1    => $cashflow->{balance}->cash($periods),
                    A5_A6 => $cashflow->{balance}->cash($periods),
                    A2    => $periods->indexPrevious,
                },
            ),
            A1 => {
                name => $periods->decorate(
                    $cashflow->{balance}{reserve} ? 'Distributions (£)'
                    : 'Cashflow to/from investors (£)'
                ),
                rounding => 1,
            },
        ]
    );
}

sub investors {
    my ( $cashflow, $periods ) = @_;
    $cashflow->statement($periods)->{A1};
}

sub fromOperations {
    my ( $cashflow, $periods ) = @_;
    $cashflow->statement($periods)->{A5};
}

sub workingCapitalMovements {
    my ( $cashflow, $periods ) = @_;
    $cashflow->{workingCapitalMovements}{ 0 + $periods } ||= CalcBlock(
        name  => 'Movements in working capital',
        items => [
            [
                A5_A6 => $cashflow->{balance}->workingCapital($periods),
                A1    => $periods->indexPrevious,
                {
                    name => $periods->decorate('Opening working capital (£)'),
                    defaultFormat => '0boldsoft',
                    arithmetic    => '=INDEX(A5_A6,A1)',
                }
            ],
            [
                (
                    map {
                        Arithmetic(
                            name => (
                                map {
                                    local $_ = $_;
s#^Cash released/absorbed in#Increase/decrease through#;
                                    $_;
                                } $_->objectShortName
                            ),
                            defaultFormat => '0soft',
                            arithmetic    => '=-1*A1',
                            arguments     => { A1 => $_ },
                        );
                    } @{ $cashflow->statement($periods) }{qw(A11 A12 A13)}
                ),
                $periods->decorate('Increase/decrease in working capital (£)'),
            ],
            $periods->decorate('Closing working capital (£)'),
        ],
    );
}

sub profitAndLossReserveMovements {
    my ( $cashflow, $periods ) = @_;
    $cashflow->{profitAndLossReserveMovements}{ 0 + $periods } ||= CalcBlock(
        name  => 'Movements in the profit and loss reserve',
        items => [
            [
                A5_A6 => $cashflow->{balance}->profitAndLossReserve($periods),
                A1    => $periods->indexPrevious,
                {
                    name => $periods->decorate('Opening P&L reserve (£)'),
                    defaultFormat => '0boldsoft',
                    arithmetic    => '=INDEX(A5_A6,A1)',
                }
            ],
            $cashflow->{income}->earnings($periods),
            Arithmetic(
                name          => $periods->decorate('Dividends paid (£)'),
                defaultFormat => '0soft',
                arithmetic    => '=-1*A1',
                arguments => { A1 => $cashflow->statement($periods)->{A1}, },
            ),
            $periods->decorate('Closing P&L reserve (£)'),
        ],
    );
}

sub equityInitialAndRaised {
    my ( $cashflow, $periods ) = @_;
    $cashflow->{equityInitialAndRaised}{ 0 + $periods } ||= Arithmetic(
        name          => 'Equity raised, including initial equity (£)',
        defaultFormat => '0soft',
        arithmetic    => '=A1+IF(A2>A3,A4,0)',
        arguments     => {
            A1 => $cashflow->{balance}{reserve}->raised($periods),
            A2 => $periods->firstDay,
            A3 => $periods->lastDay,
            A4 => $cashflow->{balance}->initialEquity($periods),
        },
    );
}

sub chart_equity_dividends {
    my ( $cashflow, $periods ) = @_;
    require SpreadsheetModel::Chart;
    SpreadsheetModel::Chart->new(
        name         => 'Equity raised and dividends',
        type         => 'column',
        height       => 280,
        width        => 640,
        instructions => [
            add_series => $cashflow->investors($periods),
            add_series => $cashflow->equityInitialAndRaised($periods),
            set_legend => [ position => 'top' ],
        ],
    );
}

1;
