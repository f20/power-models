package Financial::Stream;

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
    my ( $class, $model, %args ) = @_;
    $args{signAdjustment}    ||= '';
    $args{databaseName}      ||= 'sales items';
    $args{itemName}          ||= 'Item ';
    $args{flowName}          ||= 'sales (£)';
    $args{balanceName}       ||= 'debtors (£)';
    $args{bufferName}        ||= 'debtor cash buffer (£)';
    $args{defaultFormatBase} ||= '0';
    $args{numLines}          ||= 4;
    $args{model} = $model;
    bless \%args, $class;
}

sub finish {
    my ($stream) = @_;
    return unless $stream->{database};
    my @columns =
      grep { $_ } @{ $stream->{database} }{
        qw(
          names
          startDate
          endDate
          annual
          growth
          averageDays
          worstDays
          )
      };
    Columnset(
        name     => ucfirst( $stream->{databaseName} ),
        columns  => \@columns,
        appendTo => $stream->{model}{inputTables},
        dataset  => $stream->{model}{dataset},
        number   => $stream->{inputTableNumber},
    );
}

sub labelset {
    my ($stream) = @_;
    $stream->{labelset} ||= Labelset(
        editable => (
            $stream->{database}{names} ||= Dataset(
                name          => "Names for $stream->{databaseName}",
                defaultFormat => 'texthard',
                rows          => $stream->labelsetNoNames,
                data => [ map { '' } $stream->labelsetNoNames->indices ],
            )
        ),
    );
}

sub labelsetNoNames {
    my ($stream) = @_;
    $stream->{labelsetNoNames} ||= Labelset(
        name          => 'Items without names',
        defaultFormat => 'thitem',
        list          => [ 1 .. $stream->{numLines} ]
    );
}

sub annual {
    my ($stream) = @_;
    $stream->{database}{annual} ||= Dataset(
        name          => 'Annual ' . $stream->{flowName},
        defaultFormat => $stream->{defaultFormatBase} . 'hard',
        rows          => $stream->labelsetNoNames,
        data          => [ map { 0 } @{ $stream->labelsetNoNames->{list} } ],
    );
}

sub growth {
    my ($stream) = @_;
    $stream->{database}{growth} ||= Dataset(
        name          => 'Annual growth rate',
        defaultFormat => '%hard',
        rows          => $stream->labelsetNoNames,
        data          => [ map { 0 } @{ $stream->labelsetNoNames->{list} } ],
    );
}

sub startDate {
    my ($stream) = @_;
    $stream->{database}{startDate} ||= Dataset(
        name          => 'Start date',
        defaultFormat => 'datehard',
        rows          => $stream->labelsetNoNames,
        data          => [ map { '' } @{ $stream->labelsetNoNames->{list} } ],
    );
}

sub endDate {
    my ($stream) = @_;
    $stream->{database}{endDate} ||= Dataset(
        name          => 'End date',
        defaultFormat => 'datehard',
        rows          => $stream->labelsetNoNames,
        data          => [ map { '' } @{ $stream->labelsetNoNames->{list} } ],
    );
}

sub averageDays {
    my ($stream) = @_;
    $stream->{database}{averageDays} ||= Dataset(
        name          => 'Average credit and inventory days',
        defaultFormat => '0.0hard',
        rows          => $stream->labelsetNoNames,
        data          => [ map { 0; } @{ $stream->labelsetNoNames->{list} } ],
    );
}

sub worstDays {
    my ($stream) = @_;
    $stream->{database}{worstDays} ||= Dataset(
        name          => 'Worst-case credit and inventory days',
        defaultFormat => '0.0hard',
        rows          => $stream->labelsetNoNames,
        data          => [ map { 0; } @{ $stream->labelsetNoNames->{list} } ],
    );
}

sub stream {

    my ( $stream, $periods, $recache ) = @_;

    if ( my $cached = $stream->{flow}{ 0 + $periods } ) {
        return $cached unless $recache;
        return $stream->{flow}{ 0 + $periods } = Stack(
            name    => $cached->objectShortName,
            sources => [$cached]
        );
    }

    my $first = Arithmetic(
        name => $periods->decorate(
            ucfirst( $stream->{databaseName} ) . ': offset of first day'
        ),
        defaultFormat => '0soft',
        rows          => $stream->labelset,
        cols          => $periods->labelset,
        arithmetic    => '=IF(A201,MAX(0,A802-A202),0)',
        arguments     => {
            A201 => $stream->startDate,
            A202 => $stream->startDate,
            A802 => $periods->firstDay,
        },
    );

    my $last = Arithmetic(
        name => $periods->decorate(
            ucfirst( $stream->{databaseName} ) . ': offset of last day'
        ),
        defaultFormat => '0soft',
        rows          => $stream->labelset,
        cols          => $periods->labelset,
        arithmetic    => '=IF(A201,IF(A302,MIN(A901,A301),A902)-A202,0)',
        arguments     => {
            A201 => $stream->startDate,
            A202 => $stream->startDate,
            A301 => $stream->endDate,
            A302 => $stream->endDate,
            A901 => $periods->lastDay,
            A902 => $periods->lastDay,
        },
    );

    $stream->{flow}{ 0 + $periods } ||= GroupBy(
        name          => $periods->decorate( ucfirst( $stream->{flowName} ) ),
        defaultFormat => $stream->{defaultFormatBase} . 'soft',
        cols          => $periods->labelset,
        source        => Arithmetic(
            name => $periods->decorate("Details of $stream->{flowName}"),
            defaultFormat => $stream->{defaultFormatBase} . 'soft',
            arithmetic    => '='
              . $stream->{signAdjustment}
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
                A501 => $stream->annual,
                A701 => $stream->growth,
                A702 => $stream->growth,
                A703 => $stream->growth,
                A704 => $stream->growth,
            }
        ),
    );

}

sub balance {
    my ( $stream, $periods ) = @_;
    $stream->{balance}{ 0 + $periods } ||= GroupBy(
        name => $periods->decorate( ucfirst( $stream->{balanceName} ) ),
        defaultFormat => $stream->{defaultFormatBase} . 'soft',
        cols          => $periods->labelset,
        source        => Arithmetic(
            name =>
              $periods->decorate( 'Details of ' . $stream->{balanceName} ),
            defaultFormat => $stream->{defaultFormatBase} . 'soft',
            rows          => $stream->labelset,
            cols          => $periods->labelset,
            arithmetic    => '='
              . $stream->{signAdjustment}
              . '(1+A701)^((A903-A202)/365.25)*A501*'
              . 'MIN(A601,'
              . 'IF(A302>0,MAX(0,A301+A602-A901+1),A603),'
              . 'MAX(0,A902-A201+1))'
              . '/365.25',
            arguments => {
                A201 => $stream->startDate,
                A202 => $stream->startDate,
                A301 => $stream->endDate,
                A302 => $stream->endDate,
                A501 => $stream->annual,
                A601 => $stream->averageDays,
                A602 => $stream->averageDays,
                A603 => $stream->averageDays,
                A701 => $stream->growth,
                A901 => $periods->lastDay,
                A902 => $periods->lastDay,
                A903 => $periods->lastDay,
            }
        ),
    );
}

sub buffer {
    my ( $stream, $periods ) = @_;
    $stream->{buffer}{ 0 + $periods } ||= GroupBy(
        name          => $periods->decorate( ucfirst( $stream->{bufferName} ) ),
        defaultFormat => $stream->{defaultFormatBase} . 'soft',
        cols          => $periods->labelset,
        source        => Arithmetic(
            name => $periods->decorate( 'Details of ' . $stream->{bufferName} ),
            defaultFormat => $stream->{defaultFormatBase} . 'soft',
            arithmetic    => '=MAX(0,'
              . $stream->{signAdjustment}
              . '(1+A701)^((A903-A202)/365.25)*A501*'
              . 'MIN(A601,'
              . 'IF(A302>0,MAX(0,A301+A602-A901+1),A603),'
              . 'MAX(0,A902-A201+1))'
              . '/365.25)-A1',
            arguments => {
                A1   => $stream->balance($periods)->{source},
                A201 => $stream->startDate,
                A202 => $stream->startDate,
                A301 => $stream->endDate,
                A302 => $stream->endDate,
                A501 => $stream->annual,
                A601 => $stream->worstDays,
                A602 => $stream->worstDays,
                A603 => $stream->worstDays,
                A701 => $stream->growth,
                A901 => $periods->lastDay,
                A902 => $periods->lastDay,
                A903 => $periods->lastDay,
            }
        ),
    );
}

1;

