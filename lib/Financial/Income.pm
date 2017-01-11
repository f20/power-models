package Financial::Income;

=head Copyright licence and disclaimer

Copyright 2015-2017 Franck Latrémolière, Reckon LLP and others.

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
use SpreadsheetModel::CalcBlock;

sub new {
    my ( $class, %hash ) = @_;
    $hash{$_} || die __PACKAGE__ . " needs a $_ attribute"
      foreach qw(model assets sales costSales expenses debt);
    bless \%hash, $class;
}

sub taxRate {
    my ( $income, $periods ) = @_;
    $income->{taxRate}{ 0 + $periods } ||= Arithmetic(
        name          => $periods->decorate('Applicable tax rate'),
        defaultFormat => '%copy',
        cols          => $periods->labelset,
        arithmetic    => '=A1',
        arguments     => {
            A1 => $income->{constantRate} ||= Constant(
                name          => 'Tax rate',
                defaultFormat => '%con',
                data          => [0.2],
            ),
        }
    );
}

sub tax {
    my ( $income, $periods ) = @_;
    return $income->{tax}{ 0 + $periods } if $income->{tax}{ 0 + $periods };
    my $tax;
    if ( !$income->{model}{tax} ) {
        $tax = Constant(
            name => $periods->decorate('No tax (£)'),
            cols => $periods->labelset,
            data => [ map { 0 } @{ $periods->labelset->{list} } ],
        );
    }
    elsif ( $income->{model}{tax} =~ /uk/i ) {
        $tax = Arithmetic(
            name          => $periods->decorate('Tax (£)'),
            defaultFormat => '0soft',
            arithmetic    => join( '+',
                '=-1*A9*(A1+A21',
                ( map { 'A' . ( $_ + 22 ); } 0 .. $#{ $income->{expenses} } ),
                'A3+A8)' ),
            arguments => {
                A1  => $income->{sales}->stream($periods),
                A21 => $income->{costSales}->stream($periods),
                (
                    map {
                        ( 'A'
                              . ( $_ + 22 ) =>
                              $income->{expenses}[$_]->stream($periods) );
                    } 0 .. $#{ $income->{expenses} }
                ),
                A3 => $income->{assets}->capitalAllowance($periods),
                A8 => $income->{debt}->interest($periods),
                A9 => $income->taxRate($periods),
            },
        );
    }
    else {
        $tax = Arithmetic(
            name          => $periods->decorate('Tax (£)'),
            defaultFormat => '0soft',
            arithmetic    => join( '+',
                '=-1*A9*(A1+A21',
                ( map { 'A' . ( $_ + 22 ); } 0 .. $#{ $income->{expenses} } ),
                'A3+A4+A8)' ),
            arguments => {
                A1  => $income->{sales}->stream($periods),
                A21 => $income->{costSales}->stream($periods),
                (
                    map {
                        ( 'A'
                              . ( $_ + 22 ) =>
                              $income->{expenses}[$_]->stream($periods) );
                    } 0 .. $#{ $income->{expenses} }
                ),
                A3 => $income->{assets}->depreciationCharge($periods),
                A4 => $income->{assets}->disposalGainLoss($periods),
                A8 => $income->{debt}->interest($periods),
                A9 => $income->taxRate($periods),
            },
        );
    }
    $income->{tax}{ 0 + $periods } = $tax;
}

sub statement {
    my ( $income, $periods ) = @_;
    my $incomeBlock = $income->{statement}{ 0 + $periods } ||= CalcBlock(
        name => $income->{model}{oldTerminology}
        ? 'Profit and loss account'
        : 'Income statement',
        consolidate => 1,
        items       => [
            [
                [
                    [
                        [
                            $income->{sales}->stream($periods),
                            $income->{costSales}->stream($periods),
                            A4 => $periods->decorate('Gross profits (£)'),
                        ],
                        (
                            map { $_->stream($periods); }
                              @{ $income->{expenses} }
                        ),
                        A1 => $periods->decorate(
                            Label(
                                'EBITDA (£)',
                                'EBITDA: earnings before'
                                  . ' interest, tax and depreciation (£)'
                            )
                        ),
                    ],
                    $income->{assets}->depreciationCharge($periods),
                    $income->{assets}->disposalGainLoss($periods),
                    A2 => $periods->decorate(
                        Label(
                            'EBIT (£)',
                            'EBIT: earnings before interest and tax (£)'
                        )
                    ),
                ],
                $income->{debt}->interest($periods),
                $periods->decorate('Earnings before tax (£)'),
            ],
            $income->tax($periods),
            A3 => $periods->decorate('Earnings (£)'),
        ],
    );
}

sub gross {
    my ( $income, $periods ) = @_;
    $income->statement($periods)->{A4};
}

sub ebitda {
    my ( $income, $periods ) = @_;
    $income->statement($periods)->{A1};
}

sub ebit {
    my ( $income, $periods ) = @_;
    $income->statement($periods)->{A2};
}

sub earnings {
    my ( $income, $periods ) = @_;
    $income->statement($periods)->{A3};
}

sub chart {
    my ( $income, $periods ) = @_;
    require SpreadsheetModel::Chart;
    SpreadsheetModel::Chart->new(
        name   => 'Performance',
        type   => 'column',
        height => 280,
        width  => 640,
        $periods->{priorPeriod} ? ( ignore_left => 1 ) : (),
        instructions => [
            add_series => $income->{sales}->stream($periods),
            add_series => $income->gross($periods),
            add_series => $income->ebitda($periods),
            add_series => $income->ebit($periods),
            add_series => $income->earnings($periods),
            set_legend => [
                position => 'right',
                font     => { name => 'Calibri', size => 13, },
            ],
            set_x_axis => [
                num_font => { name => 'Calibri', size => 13, },
            ],
            set_y_axis => [
                num_font => { name => 'Calibri', size => 13, },
            ],
        ],
    );
}

1;
