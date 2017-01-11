package Financial::FlowOnce;

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
use base 'Financial::FlowBase';

sub finish {
    my ($flow) = @_;
    return unless $flow->{inputDataColumns};
    my @columns =
      grep { ref $_ =~ /Dataset/; }
      @{ $flow->{inputDataColumns} }
      {qw(names startDate endDate amount averageDays maxDays minDays)};
    Columnset(
        name     => $flow->{name},
        columns  => \@columns,
        appendTo => $flow->{model}{inputTables},
        dataset  => $flow->{model}{dataset},
        number   => $flow->{number},
    );
}

sub amount {
    my ($flow) = @_;
    $flow->{inputDataColumns}{amount} ||= Dataset(
        name          => 'Total ' . $flow->{show_flow},
        defaultFormat => $flow->{show_formatBase} . 'hard',
        rows          => $flow->labelsetNoNames,
        data          => [ map { 0 } @{ $flow->labelsetNoNames->{list} } ],
    );
}

sub stream {

    my ( $flow, $periods, $recache ) = @_;

    if ( my $cached = $flow->{stream}{ 0 + $periods } ) {
        return $cached unless $recache;
        return $flow->{stream}{ 0 + $periods } = Stack(
            name    => $cached->objectShortName,
            sources => [$cached],
        );
    }

    $flow->{stream}{ 0 + $periods } ||= GroupBy(
        name          => $periods->decorate( ucfirst( $flow->{show_flow} ) ),
        defaultFormat => $flow->{show_formatBase} . 'soft',
        cols          => $periods->labelset,
        source        => Arithmetic(
            name => $periods->decorate("Details of $flow->{show_flow}"),
            rows => $flow->labelset,
            cols => $periods->labelset,
            defaultFormat => $flow->{show_formatBase} . 'soft',
            arithmetic    => $flow->prefix_calc
              . 'IF(OR(A201<A301,A401<A101),0,A501*'
              . '(MIN(A202,A402)-MAX(A102,A302)+1)/(A203-A103+1))',
            arguments => {
                A101 => $flow->startDate,
                A102 => $flow->startDate,
                A103 => $flow->startDate,
                A201 => $flow->endDate,
                A202 => $flow->endDate,
                A203 => $flow->endDate,
                A301 => $periods->firstDay,
                A302 => $periods->firstDay,
                A401 => $periods->lastDay,
                A402 => $periods->lastDay,
                A501 => $flow->amount,
            }
        ),
    );

}

sub aggregate {

    my ( $flow, $periods, $recache ) = @_;

    if ( my $cached = $flow->{aggregate}{ 0 + $periods } ) {
        return $cached unless $recache;
        return $flow->{aggregate}{ 0 + $periods } = Stack(
            name    => $cached->objectShortName,
            sources => [$cached],
        );
    }

    $flow->{aggregate}{ 0 + $periods } ||= GroupBy(
        name => $periods->decorate( 'Total to date: ' . $flow->{show_flow} ),
        defaultFormat => $flow->{show_formatBase} . 'soft',
        cols          => $periods->labelset,
        source        => Arithmetic(
            name => $periods->decorate(
                'Details of total to date: ' . $flow->{show_flow}
            ),
            rows          => $flow->labelset,
            cols          => $periods->labelset,
            defaultFormat => $flow->{show_formatBase} . 'soft',
            arithmetic    => $flow->prefix_calc
              . 'IF(A401<A101,0,A501*'
              . '(MIN(A202,A402)-A102+1)/(A203-A103+1))',
            arguments => {
                A101 => $flow->startDate,
                A102 => $flow->startDate,
                A103 => $flow->startDate,
                A201 => $flow->endDate,
                A202 => $flow->endDate,
                A203 => $flow->endDate,
                A401 => $periods->lastDay,
                A402 => $periods->lastDay,
                A501 => $flow->amount,
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
              . 'A501/(A301-A201+1)*'
              . 'MIN(0+A601,'
              . 'MAX(0,A302+A602-A901+1),'
              . 'MAX(0,A902-A202+1))',
            arguments => {
                A201 => $flow->startDate,
                A202 => $flow->startDate,
                A301 => $flow->endDate,
                A302 => $flow->endDate,
                A501 => $flow->amount,
                A601 => $flow->averageDays,
                A602 => $flow->averageDays,
                A901 => $periods->lastDay,
                A902 => $periods->lastDay,
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
              . 'A501/(A303-A202+1)*'
              . 'MIN(0+A601,'
              . 'IF(A302>0,MAX(0,A301+A602-A901+1),0+A603),'
              . 'MAX(0,A902-A201+1))-A1',
            arguments => {
                A1   => $flow->balance($periods)->{source},
                A201 => $flow->startDate,
                A202 => $flow->startDate,
                A301 => $flow->endDate,
                A302 => $flow->endDate,
                A303 => $flow->endDate,
                A501 => $flow->amount,
                A601 => $flow->worstDays,
                A602 => $flow->worstDays,
                A603 => $flow->worstDays,
                A901 => $periods->lastDay,
                A902 => $periods->lastDay,
            }
        ),
    );
}

1;

