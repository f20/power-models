package Financial::CashCalc;

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
      foreach qw(model sales expenses);
    bless \%hash, $class;
}

sub coincidence {
    my ( $cashCalc, $periods ) = @_;
    $cashCalc->{coincidence} ||= Dataset(
        name          => 'Cash buffer coincidence factor',
        singleRowName => 'Allowance',
        defaultFormat => '%hard',
        number        => 1448,
        dataset       => $cashCalc->{model}->{dataset},
        appendTo      => $cashCalc->{model}->{inputTables},
        data          => [0.8],
    );
}

sub required {
    my ( $cashCalc, $periods ) = @_;
    $cashCalc->{required}{ 0 + $periods } ||= Arithmetic(
        name => $periods->decorate('Cash reserve required for operations (£)'),
        defaultFormat => '0soft',
        arithmetic    => join( '',
            '=(A1', ( map { "+A3$_" } 0 .. $#{ $cashCalc->{expenses} } ),
            , ')*A2' ),
        arguments => {
            A1 => $cashCalc->{sales}->buffer($periods),
            (
                map {
                    ( "A3$_" => $cashCalc->{expenses}[$_]->buffer($periods) );
                } 0 .. $#{ $cashCalc->{expenses} }
            ),
            A2 => $cashCalc->coincidence,
        },
    );
}

sub total {
    my ( $cashCalc, $periods, $reserve ) = @_;
    return $cashCalc->required($periods) unless $reserve;
    $cashCalc->{total}{ 0 + $reserve }{ 0 + $periods } ||= Arithmetic(
        name          => $periods->decorate('Total cash (£)'),
        defaultFormat => '0soft',
        arithmetic    => '=A1+A2',
        arguments     => {
            A1 => $cashCalc->required($periods),
            A2 => $reserve->spareCash($periods),
        },
    );
}

1;
