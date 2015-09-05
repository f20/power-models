package Financial::Ratios;

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
      foreach qw(model income balance cashflow);
    bless \%hash, $class;
}

sub statement {

    my ( $ratios, $periods ) = @_;

    return $ratios->{statement}{ 0 + $periods }
      if $ratios->{statement}{ 0 + $periods };

    my @ratios;

    push @ratios,
      Arithmetic(
        name          => 'Gross profit margin',
        defaultFormat => '%soft',
        arithmetic    => '=IF(A3>0,A2/A1,"")',
        arguments     => {
            A1 => $ratios->{income}{sales}->stream($periods),
            A2 => $ratios->{income}->gross($periods),
            A3 => $ratios->{income}{sales}->stream($periods),
        },
      );

    push @ratios,
      Arithmetic(
        name          => 'EBITDA profit margin',
        defaultFormat => '%soft',
        arithmetic    => '=IF(A3>0,A2/A1,"")',
        arguments     => {
            A1 => $ratios->{income}{sales}->stream($periods),
            A2 => $ratios->{income}->ebitda($periods),
            A3 => $ratios->{income}{sales}->stream($periods),
        },
      );

    push @ratios,
      Arithmetic(
        name          => 'Earnings profit margin',
        defaultFormat => '%soft',
        arithmetic    => '=IF(A3>0,A2/A1,"")',
        arguments     => {
            A1 => $ratios->{income}{sales}->stream($periods),
            A2 => $ratios->{income}->earnings($periods),
            A3 => $ratios->{income}{sales}->stream($periods),
        },
      );

    push @ratios,
      Arithmetic(
        name          => 'Interest cover by cashflow from operations',
        defaultFormat => '0.00soft',
        arithmetic    => '=IF(A3<0,-1*A2/A1,"")',
        arguments     => {
            A1 => $ratios->{balance}{debt}->interest($periods),
            A2 => $ratios->{cashflow}->fromOperations($periods),
            A3 => $ratios->{balance}{debt}->interest($periods),
        },
      );

    push @ratios,
      A12 => Arithmetic(
        name          => 'Interest cover by EBITDA',
        defaultFormat => '0.00soft',
        arithmetic    => '=IF(A3<0,-1*A2/A1,"")',
        arguments     => {
            A1 => $ratios->{balance}{debt}->interest($periods),
            A2 => $ratios->{income}->ebitda($periods),
            A3 => $ratios->{balance}{debt}->interest($periods),
        },
      );

    push @ratios,
      Arithmetic(
        name          => 'Interest cover by EBIT',
        defaultFormat => '0.00soft',
        arithmetic    => '=IF(A3<0,-1*A2/A1,"")',
        arguments     => {
            A1 => $ratios->{balance}{debt}->interest($periods),
            A2 => $ratios->{income}->ebit($periods),
            A3 => $ratios->{balance}{debt}->interest($periods),
        },
      );

    push @ratios,
      Arithmetic(
        name          => 'Interest cover by earnings before interest',
        defaultFormat => '0.00soft',
        arithmetic    => '=IF(A3<0,1-1*A2/A1,"")',
        arguments     => {
            A1 => $ratios->{balance}{debt}->interest($periods),
            A2 => $ratios->{income}->earnings($periods),
            A3 => $ratios->{balance}{debt}->interest($periods),
        },
      );

    push @ratios,
      A10 => Arithmetic(
        name          => 'Gearing',
        defaultFormat => '%soft',
        arithmetic    => '=A1/(A11-A2)',
        arguments     => {
            A1  => $ratios->{balance}{debt}->due($periods),
            A11 => $ratios->{balance}{debt}->due($periods),
            A2  => $ratios->{balance}->equity($periods),
        },
      );

    $ratios->{statement}{ 0 + $periods } = CalcBlock(
        name        => $periods->decorate('Ratios'),
        consolidate => 1,
        items       => \@ratios,
    );

}

sub charts {
    my ( $ratios, $periods ) = @_;
    require SpreadsheetModel::Chart;
    SpreadsheetModel::Chart->new(
        name         => 'EBITDA interest cover',
        type         => 'column',
        height       => 360,
        width        => 110 * @{ $periods->labelset->{list} },
        instructions => [
            add_series => $ratios->statement($periods)->{A12},
            set_x_axis => [ name => 'Year' ],
            set_y_axis => [ name => 'Cover ratio' ],
            set_legend => [ position => 'none' ],
            combine    => [
                type         => 'line',
                instructions => [
                    add_series => Constant(
                        name => 'Reference EBITDA interest cover',
                        cols => $periods->labelset,
                        data => [ map { 3 } @{ $periods->labelset->{list} } ],
                    ),
                ],
            ],
        ],
      ),
      SpreadsheetModel::Chart->new(
        name         => 'Gearing',
        type         => 'column',
        height       => 360,
        width        => 110 * @{ $periods->labelset->{list} },
        instructions => [
            add_series => $ratios->statement($periods)->{A10},
            set_x_axis => [ name => 'Year' ],
            set_y_axis => [ name => 'Gearing' ],
            set_legend => [ position => 'none' ],
        ],
      );
}

1;
