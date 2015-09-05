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
    return $reserve->{raisingDates} if $reserve->{raisingDates};
    Columnset(
        name     => 'Equity raising rounds in chronological order',
        appendTo => $reserve->{model}{inputTables},
        dataset  => $reserve->{model}{dataset},
        number   => 1470,
        columns  => [
            Dataset(
                name          => 'Name of round (in chronological order)',
                defaultFormat => 'texthard',
                rows          => $reserve->labelset,
                data          => [ map { '' } @{ $reserve->labelset->{list} } ],
            ),
            $reserve->{raisingDates} = Dataset(
                name          => 'Date',
                defaultFormat => 'datehard',
                rows          => $reserve->labelset,
                data          => [ map { '' } @{ $reserve->labelset->{list} } ],
            ),
        ],
    );
    $reserve->{raisingDates};
}

sub labelset {
    my ($reserve) = @_;
    $reserve->{labelset} ||= Labelset(
        name          => 'Equity tranches',
        defaultFormat => 'thitem',
        list          => [ 1 .. $reserve->{model}{numEquity} || 3 ]
    );
}

sub trancheId {
    my ($reserve) = @_;
    return $reserve->{trancheId} if $reserve->{trancheId};
    die if $reserve->{periods}{reverseTime};
    my $match = Arithmetic(
        name          => 'Index of relevant equity tranche',
        defaultFormat => '0soft',
        arithmetic    => '=MATCH(A1,A2_A3)',
        arguments     => {
            A1    => $reserve->{periods}->lastDay,
            A2_A3 => $reserve->raisingDates,
        },
    );
    $reserve->{trancheId} = Arithmetic(
        name          => 'Relevant equity tranche',
        defaultFormat => 'datesoft',
        arithmetic    => '=IF(A1-A201>-1,IF(ISNUMBER(A702),'
          . 'INDEX(A4_A5,A701),"Initial equity"),"Initial equity")',
        arguments => {
            A1    => $reserve->{periods}->lastDay,
            A201  => $reserve->{periods}->firstDay,
            A4_A5 => $reserve->raisingDates,
            A701  => $match,
            A702  => $match,
        },
    );
}

sub cashNeededToNextTranche {
    my ($reserve) = @_;
    $reserve->{cashNeededToNextTranche} ||= SpreadsheetModel::Custom->new(
        name => $reserve->{periods}->decorate(
                'Spare cash (opening or raised) needed to reach'
              . ' next equity raising tranche (£)'
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
            A31_A32 => $reserve->cashNeededToNextTranche,
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
        name          => $periods->decorate('Closing spare cash (£)'),
        defaultFormat => '0soft',
        arithmetic    => '=IF(ISNUMBER(MATCH(A1+1,A5_A6,0)),'
          . 'INDEX(A2_A3,MATCH(A11+1,A51_A61,0)),0)',
        arguments => {
            A1      => $periods->lastDay,
            A11     => $periods->lastDay,
            A5_A6   => $reserve->{periods}->openingDay,
            A51_A61 => $reserve->{periods}->openingDay,
            A2_A3   => Arithmetic(
                name          => 'Opening spare cash (£)',
                defaultFormat => '0soft',
                arithmetic =>
                  '=IF(ISNUMBER(A1),IF(A81=INDEX(A82_A83,A11),A2,0),A3)',
                arguments => {
                    A1      => $reserve->{periods}->indexPrevious,
                    A11     => $reserve->{periods}->indexPrevious,
                    A2      => $reserve->cashNeededToNextTranche,
                    A3      => $reserve->cashNeededToNextTranche,
                    A81     => $reserve->trancheId,
                    A82_A83 => $reserve->trancheId,
                }
            ),
        },
    );
}

sub raised {
    my ( $reserve, $periods ) = @_;
    $reserve->{raised}{ 0 + $periods } ||= Arithmetic(
        name          => $periods->decorate('Equity raised (£)'),
        defaultFormat => '0soft',
        arithmetic    => '=SUMPRODUCT((A2_A3<=A1)*(A21_A31>=A11)*A4_A5)',
        arguments     => {
            A1      => $periods->lastDay,
            A11     => $periods->firstDay,
            A2_A3   => $reserve->raisingDates,
            A21_A31 => $reserve->raisingDates,
            A4_A5   => $reserve->amountsRaised,
        },
    );
}

sub shareCapital {
    my ( $reserve, $periods ) = @_;
    $reserve->{shareCapital}{ 0 + $periods } ||= Arithmetic(
        name          => $periods->decorate('Share capital (£)'),
        defaultFormat => '0soft',
        arithmetic    => '=SUMIF(A2_A3,"<="&A1,A4_A5)',
        arguments     => {
            A1    => $periods->lastDay,
            A2_A3 => $reserve->raisingDates,
            A4_A5 => $reserve->amountsRaised,
        },
    );
}

1;

