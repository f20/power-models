package Financial::Reserve;

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
      foreach qw(model cashflow periods);
    bless \%hash, $class;
}

sub finish {
}

sub raisingDates {
    my ($reserve) = @_;
    $reserve->{raisingDates} ||= Dataset(
        name          => 'Equity raising dates in chronological order',
        defaultFormat => 'datehard',
        rows          => $reserve->labelset,
        data          => [ map { '' } @{ $reserve->labelset->{list} } ],
        appendTo      => $reserve->{model}{inputTables},
        dataset       => $reserve->{model}{dataset},
        number        => 1470,
    );
}

sub labelset {
    my ($reserve) = @_;
    $reserve->{labelset} ||= Labelset(
        name => 'Equity tranches',
        list => [
            map { 'Equity tranche #' . $_ } 1 .. $reserve->{model}{numEquity}
              || 3
        ]
    );
}

sub trancheId {
    my ($reserve) = @_;
    return $reserve->{trancheId} if $reserve->{trancheId};
    die if $reserve->{periods}{reverseTime};
    my $match = Arithmetic(
        name          => 'Index of relevant equity tranche',
        defaultFormat => '0soft',
        arithmetic    => '=MATCH(A1-1,A2_A3)',
        arguments     => {
            A1    => $reserve->{periods}->firstDay,
            A2_A3 => $reserve->raisingDates,
        },
    );
    $reserve->{trancheId} = Arithmetic(
        name          => 'Relevant equity tranche',
        defaultFormat => 'datesoft',
        arithmetic    => '=IF(A1-A201>-1,IF(ISNUMBER(A702),'
          . 'INDEX(A4_A5,A701),"Initial equity"),"Not applicable")',
        arguments => {
            A1    => $reserve->{periods}->lastDay,
            A201  => $reserve->{periods}->firstDay,
            A4_A5 => $reserve->raisingDates,
            A701  => $match,
            A702  => $match,
        },
    );
}

sub openingCashNeeded {
    my ($reserve) = @_;
    $reserve->{openingCashNeeded} ||= SpreadsheetModel::Custom->new(
        name => $reserve->{periods}->decorate(
            'Opening spare cash needed to reach next equity raising tranche (£)'
        ),
        defaultFormat => '0soft',
        cols          => $reserve->{periods}->labelset,
        arithmetic    => '=MAX(0,IF(ISNUMBER(A51),IF(A81=INDEX(A82,A52),'
          . 'INDEX(self,A53),0),0)-A6)',
        custom => [
                '=MAX(0,IF(ISNUMBER(A51),IF(A81=INDEX(A82:A83,A52),'
              . 'INDEX(A21:A22,A53),0),0)-A6)'
        ],
        arguments => {
            A51 => $reserve->{periods}->indexNext,
            A52 => $reserve->{periods}->indexNext,
            A53 => $reserve->{periods}->indexNext,
            A81 => $reserve->trancheId,
            A82 => $reserve->trancheId,
            A83 => $reserve->trancheId,
            A6  => $reserve->{cashflow}->investors( $reserve->{periods} ),
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            my $lastCol = $self->lastCol;
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0],
                  qr/\bA21\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $self->{$wb}{row},
                    $self->{$wb}{col},
                    0, 1
                  ),
                  qr/\bA22\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $self->{$wb}{row},
                    $self->{$wb}{col} + $lastCol,
                    0, 1
                  ),
                  map {
                    qr/\b$_\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                          /\bA82\b/ ? ( $rowh->{$_}, $colh->{$_}, 0, 1 )
                        : /\bA83\b/
                        ? ( $rowh->{$_}, $colh->{$_} + $lastCol, 0, 1 )
                        : ( $rowh->{$_}, $colh->{$_} + $x )
                      )
                  } @$pha;
            };
        },
    );
}

sub amountsRaised {
    my ($reserve) = @_;
    return $reserve->{amountsRaised} if $reserve->{amountsRaised};
    die if $reserve->{periods}{reverseTime};
    $reserve->{amountsRaised} = Arithmetic(
        name          => 'Tranches of equity to be raised (£)',
        defaultFormat => '0soft',
        arithmetic    => '=IF(A1,INDEX(A31_A32,MATCH(A102,A41_A42,0)),0)',
        arguments     => {
            A1      => $reserve->raisingDates,
            A102    => $reserve->raisingDates,
            A31_A32 => Arithmetic(
                name => 'Opening cash needed,'
                  . ' adjusted for previous period cashflow (£)',
                defaultFormat => '0soft',
                arithmetic    => '=MAX(0,A1-MAX(0,INDEX(A2_A3,A4)))',
                arguments     => {
                    A1 => $reserve->openingCashNeeded,
                    A2_A3 =>
                      $reserve->{cashflow}->investors( $reserve->{periods} ),
                    A4 => $reserve->{periods}->indexPrevious,
                },
            ),
            A41_A42 => $reserve->trancheId,
        },
    );
}

sub raisingSchedule {
    my ($reserve) = @_;
    $reserve->{raisingSchedule} ||= Columnset(
        name    => 'Equity raising schedule',
        columns => [
            Stack(
                name    => 'Date of equity raising',
                sources => [ $reserve->raisingDates ]
            ),
            Stack( sources => [ $reserve->amountsRaised ] ),
        ],
    );
}

sub spareCash {
    my ( $reserve, $periods ) = @_;
    $reserve->{spareCash}{ 0 + $periods } ||= Arithmetic(
        name          => $periods->decorate('Additional spare cash (£)'),
        defaultFormat => '0copy',
        arithmetic    => '=IFERROR(INDEX(A2_A3,MATCH(A1+1,A4_A5)),0)',
        arguments     => {
            A1    => $periods->lastDay,
            A4_A5 => $reserve->{periods}->firstDay,
            A2_A3 => $reserve->openingCashNeeded,
        },
    );
}

1;

