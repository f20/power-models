package Financial::Debt;

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
    my ( $class, $model, ) = @_;
    bless { model => $model, }, $class;
}

sub finish {
    my ($debt) = @_;
    return unless $debt->{database};
    my @columns =
      grep { $_ } @{ $debt->{database} }{
        qw(
          names
          startDate
          endDate
          amount
          rate
          )
      };
    Columnset(
        name     => 'Borrowings',
        columns  => \@columns,
        appendTo => $debt->{model}{inputTables},
        dataset  => $debt->{model}{dataset},
        number   => 1460,
    );
}

sub labelset {
    my ($debt) = @_;
    $debt->{labelset} ||= Labelset(
        editable => (
            $debt->{database}{names} ||= Dataset(
                name          => 'Name of debt tranche',
                defaultFormat => 'texthard',
                rows          => $debt->labelsetNoNames,
                data          => [ map { '' } $debt->labelsetNoNames->indices ],
            )
        ),
    );
}

sub labelsetNoNames {
    my ($debt) = @_;
    $debt->{labelsetNoNames} ||= Labelset(
        name => 'Debt tranches without names',
        list =>
          [ map { 'Debt issue #' . $_ } 1 .. $debt->{model}{numDebt} || 4 ]
    );
}

sub amount {
    my ($debt) = @_;
    $debt->{database}{amount} ||= Dataset(
        name          => 'Amount (£)',
        defaultFormat => '0hard',
        rows          => $debt->labelsetNoNames,
        data          => [ map { 0 } @{ $debt->labelsetNoNames->{list} } ],
    );
}

sub rate {
    my ($debt) = @_;
    $debt->{database}{rate} ||= Dataset(
        name          => 'Interest rate',
        defaultFormat => '%hard',
        rows          => $debt->labelsetNoNames,
        data          => [ map { 0 } @{ $debt->labelsetNoNames->{list} } ],
    );
}

sub startDate {
    my ($debt) = @_;
    $debt->{database}{startDate} ||= Dataset(
        name          => 'Start date',
        defaultFormat => 'datehard',
        rows          => $debt->labelsetNoNames,
        data          => [ map { '' } @{ $debt->labelsetNoNames->{list} } ],
    );
}

sub endDate {
    my ($debt) = @_;
    $debt->{database}{endDate} ||= Dataset(
        name          => 'End date',
        defaultFormat => 'datehard',
        rows          => $debt->labelsetNoNames,
        data          => [ map { '' } @{ $debt->labelsetNoNames->{list} } ],
    );
}

sub interest {
    my ( $debt, $periods, $recache ) = @_;
    return $debt->{interest}{ 0 + $periods } = Stack(
        name    => $recache->objectShortName,
        sources => [$recache]
    ) if $recache and $recache = $debt->{interest}{ 0 + $periods };
    $debt->{interest}{ 0 + $periods } ||= GroupBy(
        name          => $periods->decorate('Interest payable (£)'),
        defaultFormat => '0soft',
        cols          => $periods->labelset,
        source        => Arithmetic(
            name => $periods->decorate('Details of interest payable (£)'),
            defaultFormat => '0soft',
            rows          => $debt->labelset,
            cols          => $periods->labelset,
            arithmetic    => '=IF(A201,MAX(0,'
              . 'IF(A301,MIN(A302,A901),A902)+1-MAX(A202,A801)'
              . ')/-365.25*A501*A601,0)',
            arguments => {
                A201 => $debt->startDate,
                A202 => $debt->startDate,
                A301 => $debt->endDate,
                A302 => $debt->endDate,
                A501 => $debt->rate,
                A601 => $debt->amount,
                A801 => $periods->firstDay,
                A901 => $periods->lastDay,
                A902 => $periods->lastDay,
            }
        ),
    );
}

sub due {
    my ( $debt, $periods ) = @_;
    $debt->{due}{ 0 + $periods } ||= GroupBy(
        name          => $periods->decorate('Debt (£)'),
        cols          => $periods->labelset,
        defaultFormat => '0soft',
        source        => Arithmetic(
            name          => $periods->decorate('Details of debt (£)'),
            defaultFormat => '0soft',
            rows          => $debt->labelset,
            cols          => $periods->labelset,
            arithmetic =>
              '=IF(A201>A901,0,IF(A301,IF(A302>A902,0-A501,0),0-A502))',
            arguments => {
                A201 => $debt->startDate,
                A301 => $debt->endDate,
                A302 => $debt->endDate,
                A501 => $debt->amount,
                A502 => $debt->amount,
                A901 => $periods->lastDay,
                A902 => $periods->lastDay,
            }
        ),
    );
}

sub raised {
    my ( $debt, $periods ) = @_;
    $debt->{capitalExpenditure}{ 0 + $periods } ||= GroupBy(
        name          => $periods->decorate('Debt raised (£)'),
        cols          => $periods->labelset,
        defaultFormat => '0soft',
        source        => Arithmetic(
            name          => $periods->decorate('Details of debt raised (£)'),
            defaultFormat => '0soft',
            rows          => $debt->labelset,
            cols          => $periods->labelset,
            arithmetic    => '=IF(OR(A201<A802,A202>A902),0,A501)',
            arguments     => {
                A201 => $debt->startDate,
                A202 => $debt->startDate,
                A501 => $debt->amount,
                A802 => $periods->firstDay,
                A902 => $periods->lastDay,
            }
        ),
    );
}

sub repaid {
    my ( $debt, $periods ) = @_;
    $debt->{capitalReceipts}{ 0 + $periods } ||= GroupBy(
        name          => $periods->decorate('Debt repaid (£)'),
        cols          => $periods->labelset,
        defaultFormat => '0soft',
        source        => Arithmetic(
            name => $periods->decorate(
                'Details of scrap value of decommissioned assets (£)'),
            defaultFormat => '0soft',
            rows          => $debt->labelset,
            cols          => $periods->labelset,
            arithmetic    => '=IF(OR(A301<A801,A302>A901),0,0-A601)',
            arguments     => {
                A301 => $debt->endDate,
                A302 => $debt->endDate,
                A601 => $debt->amount,
                A801 => $periods->firstDay,
                A901 => $periods->lastDay,
            }
        ),
    );
}

1;

