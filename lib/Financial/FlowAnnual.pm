package Financial::FlowAnnual;

=head Copyright licence and disclaimer

Copyright 2015, 2016 Franck Latrémolière, Reckon LLP and others.

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
use base 'Financial::FlowBase';

sub finish {
    my ($flow) = @_;
    return unless $flow->{database};
    my @columns =
      grep { $_ }
      @{ $flow->{database} }
      {qw(names startDate endDate annual growth averageDays maxDays minDays)};
    Columnset(
        name     => $flow->{name},
        columns  => \@columns,
        appendTo => $flow->{model}{inputTables},
        dataset  => $flow->{model}{dataset},
        number   => $flow->{number},
    );
}

sub annual {
    my ($flow) = @_;
    $flow->{database}{annual} ||= Dataset(
        name          => 'Annual ' . $flow->{show_flow},
        defaultFormat => $flow->{show_formatBase} . 'hard',
        rows          => $flow->labelsetNoNames,
        data          => [ map { 0 } @{ $flow->labelsetNoNames->{list} } ],
    );
}

sub growth {
    my ($flow) = @_;
    $flow->{database}{growth} ||= Dataset(
        name          => 'Annual growth rate',
        defaultFormat => '%hard',
        rows          => $flow->labelsetNoNames,
        data          => [ map { 0 } @{ $flow->labelsetNoNames->{list} } ],
    );
}

sub stream {

    my ( $flow, $periods, $recache ) = @_;

    if ( my $cached = $flow->{flow}{ 0 + $periods } ) {
        return $cached unless $recache;
        return $flow->{flow}{ 0 + $periods } = Stack(
            name    => $cached->objectShortName,
            sources => [$cached]
        );
    }

    my $first = Arithmetic(
        name => $periods->decorate( $flow->{name} . ': offset of first day' ),
        defaultFormat => '0soft',
        rows          => $flow->labelset,
        cols          => $periods->labelset,
        arithmetic    => '=IF(A201,MAX(0,A802-A202),0)',
        arguments     => {
            A201 => $flow->startDate,
            A202 => $flow->startDate,
            A802 => $periods->firstDay,
        },
    );

    my $last = Arithmetic(
        name => $periods->decorate( $flow->{name} . ': offset of last day' ),
        defaultFormat => '0soft',
        rows          => $flow->labelset,
        cols          => $periods->labelset,
        arithmetic    => '=IF(A201,IF(A302,MIN(A901,A301),A902)-A202,0)',
        arguments     => {
            A201 => $flow->startDate,
            A202 => $flow->startDate,
            A301 => $flow->endDate,
            A302 => $flow->endDate,
            A901 => $periods->lastDay,
            A902 => $periods->lastDay,
        },
    );

    $flow->{flow}{ 0 + $periods } ||= GroupBy(
        name          => $periods->decorate( ucfirst( $flow->{show_flow} ) ),
        defaultFormat => $flow->{show_formatBase} . 'soft',
        cols          => $periods->labelset,
        source        => Arithmetic(
            name => $periods->decorate("Details of $flow->{show_flow}"),
            defaultFormat => $flow->{show_formatBase} . 'soft',
            arithmetic    => $flow->prefix_calc
              . 'IF(OR(A103<0,A104<A203),0,A501*IF(A701,'
              . '((1+A702)^((A102+1)/365.25)-(1+A703)^(A202/365.25))/LN(1+A704)'
              . ',(A1+1-A201)/365.25))',
            arguments => {
                A1   => $last,
                A102 => $last,
                A103 => $last,
                A104 => $last,
                A201 => $first,
                A202 => $first,
                A203 => $first,
                A501 => $flow->annual,
                A701 => $flow->growth,
                A702 => $flow->growth,
                A703 => $flow->growth,
                A704 => $flow->growth,
            }
        ),
    );

}

sub balance {
    my ( $flow, $periods ) = @_;
    $flow->{balance}{ 0 + $periods } ||= GroupBy(
        name          => $periods->decorate( ucfirst( $flow->{show_balance} ) ),
        defaultFormat => $flow->{show_formatBase} . 'soft',
        cols          => $periods->labelset,
        source        => Arithmetic(
            name => $periods->decorate( 'Details of ' . $flow->{show_balance} ),
            defaultFormat => $flow->{show_formatBase} . 'soft',
            rows          => $flow->labelset,
            cols          => $periods->labelset,
            arithmetic    => $flow->prefix_calc
              . '(1+A701)^((A903-A202)/365.25)*A501*'
              . 'MIN(0+A601,'
              . 'IF(A302>0,MAX(0,A301+A602-A901+1),0+A603),'
              . 'MAX(0,A902-A201+1))'
              . '/365.25',
            arguments => {
                A201 => $flow->startDate,
                A202 => $flow->startDate,
                A301 => $flow->endDate,
                A302 => $flow->endDate,
                A501 => $flow->annual,
                A601 => $flow->averageDays,
                A602 => $flow->averageDays,
                A603 => $flow->averageDays,
                A701 => $flow->growth,
                A901 => $periods->lastDay,
                A902 => $periods->lastDay,
                A903 => $periods->lastDay,
            }
        ),
    );
}

sub buffer {
    my ( $flow, $periods ) = @_;
    $flow->{buffer}{ 0 + $periods } ||= GroupBy(
        name          => $periods->decorate( ucfirst( $flow->{show_buffer} ) ),
        defaultFormat => $flow->{show_formatBase} . 'soft',
        cols          => $periods->labelset,
        source        => Arithmetic(
            name => $periods->decorate( 'Details of ' . $flow->{show_buffer} ),
            defaultFormat => $flow->{show_formatBase} . 'soft',
            arithmetic    => $flow->prefix_calc
              . '((1+A701)^((A903-A202)/365.25)*A501*'
              . 'MIN(0+A601,'
              . 'IF(A302>0,MAX(0,A301+A602-A901+1),0+A603),'
              . 'MAX(0,A902-A201+1))'
              . '/365.25-A1)',
            arguments => {
                A1   => $flow->balance($periods)->{source},
                A201 => $flow->startDate,
                A202 => $flow->startDate,
                A301 => $flow->endDate,
                A302 => $flow->endDate,
                A501 => $flow->annual,
                A601 => $flow->worstDays,
                A602 => $flow->worstDays,
                A603 => $flow->worstDays,
                A701 => $flow->growth,
                A901 => $periods->lastDay,
                A902 => $periods->lastDay,
                A903 => $periods->lastDay,
            }
        ),
    );
}

1;

