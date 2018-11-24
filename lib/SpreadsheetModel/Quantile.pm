package SpreadsheetModel::Quantile;

# Copyright 2008-2013 Franck Latrémolière, Reckon LLP and others.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice,
# this list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
# this list of conditions and the following disclaimer in the documentation
# and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
# EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
# THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Shortcuts ':all';

our @ISA = qw(SpreadsheetModel::Arithmetic);

use Spreadsheet::WriteExcel::Utility;

sub check {

    my ($self) = @_;

    return "$self->{name} $self->{debug}: "
      . "quantile $self->{quantile}{name} is not a single cell"
      unless ref $self->{quantile}
      && $self->{quantile}->isa('SpreadsheetModel::Dataset')
      && !$self->{quantile}->lastRow
      && !$self->{quantile}->lastCol;

    return "No data to process in $self->{name} $self->{debug}"
      unless ref $self->{toUse} eq 'ARRAY';

    my @data = @{ $self->{toUse} };
    my @cond;
    @cond = @{ $self->{conditions} } if 'ARRAY' eq ref $self->{conditions};

    my @values;
    for ( my $n = 0 ; $n < @data ; ++$n ) {
        foreach my $r ( 0 .. $data[$n]->lastRow ) {
            foreach my $c ( 0 .. $data[$n]->lastCol ) {
                push @values, [ $n, $r, $c ];
            }
        }
    }

    my $valueSet =
      new SpreadsheetModel::Labelset(
        list => [ map { "Value $_" } 1 .. @values ] );

    my $kx = new SpreadsheetModel::Custom(
        name      => 'Value',
        rows      => $valueSet,
        custom    => [ map { "=A2$_"; } 0 .. $#data, ],
        arguments => { map { ( "A2$_" => $data[$_] ); } 0 .. $#data, },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                my ( $n, $r, $c ) = @{ $values[$y] };
                return '', $format, $formula->[$n],
                  qr/\bA2$n\b/ => xl_rowcol_to_cell( $rowh->{"A2$n"} + $r,
                    $colh->{"A2$n"} + $c );
            };
        }
    );

    my $cd;
    $cd = new SpreadsheetModel::Custom(
        name      => 'Value',
        rows      => $valueSet,
        custom    => [ map { "=A2$_"; } 0 .. $#cond, ],
        arguments => { map { ( "A2$_" => $cond[$_] ); } 0 .. $#cond, },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                my ( $n, $r, $c ) = @{ $values[$y] };
                return '', $format, $formula->[$n],
                  qr/\bA2$n\b/ => xl_rowcol_to_cell( $rowh->{"A2$n"} + $r,
                    $colh->{"A2$n"} + $c );
            };
        }
    ) if @cond;
    $cd = $self->{conditionMaker}->($kx)
      if 'CODE' eq ref $self->{conditionMaker};

    if ( $self->{PERCENTILE} ) {
        if ($cd) {
            $kx = Arithmetic(
                name       => 'Use this',
                arithmetic => '=IF(A2,A1,"no")',
                arguments  => { A1 => $kx, A2 => $cd }
            );
            Columnset(
                name    => "Steps to calculate $self->{name}",
                columns => [ $kx->{arguments}{A1}, $cd, $kx, ]
            );
        }
        $self->{arithmetic} = '=PERCENTILE(A1_A2,A3)';
        $self->{arguments}  = {
            A1_A2 => $kx,
            A3    => $self->{quantile},
        };
        return $self->SUPER::check;
    }

    my $counter = new SpreadsheetModel::Constant(
        name          => 'Counter',
        rows          => $valueSet,
        data          => [ 1 .. @values ],
        defaultFormat => '0connz'
    );

    my $kr1 = new SpreadsheetModel::Custom(
        name          => 'Ranking before tie break',
        defaultFormat => '0softnz',
        rows          => $valueSet,
        custom        => [
            $cd
            ? ( '=IF(A7,RANK(A1,A2:A3,1),' . ( 1 + @values ) . ')' )
            : '=RANK(A1,A2:A3,1)'
        ],
        arguments => {
            A1    => $kx,
            A2_A3 => $kx,
            $cd ? ( A7 => $cd ) : (),
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }

# http://groups.google.com/group/spreadsheet-writeexcel/browse_thread/thread/6ba2e7e8e32fb21e

            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0],
                  qr/\bA1\b/ =>
                  xl_rowcol_to_cell( $rowh->{A1} + $y, $colh->{A1} ),
                  $cd
                  ? ( qr/\bA7\b/ =>
                      xl_rowcol_to_cell( $rowh->{A7} + $y, $colh->{A7} ) )
                  : (),
                  qr/\bA2\b/ =>
                  xl_rowcol_to_cell( $rowh->{A1}, $colh->{A1}, 1, 0 ),
                  qr/\bA3\b/ => xl_rowcol_to_cell( $rowh->{A1} + $#values,
                    $colh->{A1}, 1, 0 );
            };
        }
    );

    my $kr2 = new SpreadsheetModel::Arithmetic(
        name          => 'Tie breaker',
        arguments     => { A1 => $kr1, A4 => $counter },
        arithmetic    => '=A1*' . @values . '+A4',
        defaultFormat => '0softnz'
    );

    my $kr = new SpreadsheetModel::Custom(
        name          => 'Ranking',
        defaultFormat => '0softnz',
        rows          => $valueSet,
        custom        => ['=RANK(A1,A2:A3,1)'],
        arguments     => {
            A1    => $kr2,
            A2_A3 => $kr2,
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0],
                  qr/\bA1\b/ =>
                  xl_rowcol_to_cell( $rowh->{A1} + $y, $colh->{A1} ),
                  qr/\bA2\b/ =>
                  xl_rowcol_to_cell( $rowh->{A1}, $colh->{A1}, 1, 0 ),
                  qr/\bA3\b/ => xl_rowcol_to_cell( $rowh->{A1} + $#values,
                    $colh->{A1}, 1, 0 );
            };
        }
    );

    my $ror = new SpreadsheetModel::Custom(
        name          => 'Reordering',
        defaultFormat => '0softnz',
        rows          => $valueSet,
        custom        => ['=MATCH(A1,A2:A3,0)'],
        arguments     => {
            A1    => $counter,
            A2_A3 => $kr,
            A2    => $kr,
            A3    => $kr,
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0],
                  qr/\bA1\b/ =>
                  xl_rowcol_to_cell( $rowh->{A1} + $y, $colh->{A1} ),
                  qr/\bA2\b/ =>
                  xl_rowcol_to_cell( $rowh->{A2}, $colh->{A2}, 1, 0 ),
                  qr/\bA3\b/ => xl_rowcol_to_cell( $rowh->{A3} + $#values,
                    $colh->{A3}, 1, 0 );
            };
        }
    );

    my $kxs = new SpreadsheetModel::Custom(
        name      => 'Ordered values',
        rows      => $valueSet,
        custom    => [ '=INDEX(A2:A3,A1,1)', '=A2' ],
        arguments => {
            A2_A3 => $kx,
            A1    => $ror,
            A2    => $kx,
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0],
                  qr/\bA1\b/ =>
                  xl_rowcol_to_cell( $rowh->{A1} + $y, $colh->{A1} ),
                  qr/\bA2\b/ =>
                  xl_rowcol_to_cell( $rowh->{A2}, $colh->{A2}, 1, 0 ),
                  qr/\bA3\b/ => xl_rowcol_to_cell( $rowh->{A2} + $#values,
                    $colh->{A2}, 1, 0 );
            };
        }
    );

    Columnset(
        name    => "Steps to calculate $self->{name} (part 1)",
        columns => [ $kx, $cd || (), $kr1, $counter, $kr2, $kr, $ror, $kxs, ]
    );

    my $hit =
      $cd
      ? Arithmetic(
        name       => 'Target rank',
        arithmetic => '=1+A1*(COUNTIF(A2_A3,"TRUE")-1)',
        arguments  => { A1 => $self->{quantile}, A2_A3 => $cd }
      )
      : Arithmetic(
        name       => 'Target rank',
        arithmetic => '=1+A1*(COUNT(A2_A3)-1)',
        arguments  => { A1 => $self->{quantile}, A2_A3 => $kx }
      );

    my $hiti = Arithmetic(
        name       => 'Rank immediately below',
        arithmetic => '=FLOOR(A1,1)',
        arguments  => { A1 => $hit }
    );

    my $v1 = Arithmetic(
        name       => 'Value below',
        arithmetic => '=INDEX(A1_A2,A3)',
        arguments  => { A1_A2 => $kxs, A3 => $hiti }
    );

    my $v2 = Arithmetic(
        name       => 'Value above',
        arithmetic => '=INDEX(A1_A2,A3+1)',
        arguments  => { A1_A2 => $kxs, A3 => $hiti }
    );

    my $p1 = Arithmetic(
        defaultFormat => '%soft',
        name          => 'Weight below',
        arithmetic    => '=1-A1+A2',
        arguments     => { A1 => $hit, A2 => $hiti }
    );

    Columnset(
        name    => "Steps to calculate $self->{name} (part 2)",
        columns => [ $hit, $hiti, $v1, $v2, $p1 ]
    );

    $self->{arithmetic} = '=A1*A2+IF(A5=1,0,(1-A3)*A4)';
    $self->{arguments}  = {
        A1 => $p1,
        A3 => $p1,
        A5 => $p1,
        A2 => $v1,
        A4 => $v2,
    };

    $self->SUPER::check;

}

1;
