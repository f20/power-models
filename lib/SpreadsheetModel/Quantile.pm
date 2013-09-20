package SpreadsheetModel::Quantile;

=head Copyright licence and disclaimer

Copyright 2008-2013 Franck Latrémolière, Reckon LLP and others.

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
        custom    => [ map { "=IV2$_"; } 0 .. $#data, ],
        arguments => { map { ( "IV2$_" => $data[$_] ); } 0 .. $#data, },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                my ( $n, $r, $c ) = @{ $values[$y] };
                return '', $format, $formula->[$n],
                  "IV2$n" => xl_rowcol_to_cell( $rowh->{"IV2$n"} + $r,
                    $colh->{"IV2$n"} + $c );
            };
        }
    );

    my $cd;
    $cd = new SpreadsheetModel::Custom(
        name      => 'Value',
        rows      => $valueSet,
        custom    => [ map { "=IV2$_"; } 0 .. $#cond, ],
        arguments => { map { ( "IV2$_" => $cond[$_] ); } 0 .. $#cond, },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                my ( $n, $r, $c ) = @{ $values[$y] };
                return '', $format, $formula->[$n],
                  "IV2$n" => xl_rowcol_to_cell( $rowh->{"IV2$n"} + $r,
                    $colh->{"IV2$n"} + $c );
            };
        }
    ) if @cond;
    $cd = $self->{conditionMaker}->($kx)
      if 'CODE' eq ref $self->{conditionMaker};

    if ( $self->{PERCENTILE} ) {
        if ($cd) {
            $kx = Arithmetic(
                name       => 'Use this',
                arithmetic => '=IF(IV2,IV1,"no")',
                arguments  => { IV1 => $kx, IV2 => $cd }
            );
            Columnset(
                name    => "Steps to calculate $self->{name}",
                columns => [ $kx->{arguments}{IV1}, $cd, $kx, ]
            );
        }
        $self->{arithmetic} = '=PERCENTILE(IV1_IV2,IV3)';
        $self->{arguments}  = {
            IV1_IV2 => $kx,
            IV3     => $self->{quantile},
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
            ? ( '=IF(IV7,RANK(IV1,IV2:IV3,1),' . ( 1 + @values ) . ')' )
            : '=RANK(IV1,IV2:IV3,1)'
        ],
        arguments => { IV1 => $kx, $cd ? ( IV7 => $cd ) : () },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }

# http://groups.google.com/group/spreadsheet-writeexcel/browse_thread/thread/6ba2e7e8e32fb21e

            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0],
                  IV1 => xl_rowcol_to_cell( $rowh->{IV1} + $y, $colh->{IV1} ),
                  $cd
                  ? (
                    IV7 => xl_rowcol_to_cell( $rowh->{IV7} + $y, $colh->{IV7} )
                  )
                  : (),
                  IV2 => xl_rowcol_to_cell( $rowh->{IV1}, $colh->{IV1}, 1, 0 ),
                  IV3 => xl_rowcol_to_cell( $rowh->{IV1} + $#values,
                    $colh->{IV1}, 1, 0 );
            };
        }
    );

    my $kr2 = new SpreadsheetModel::Arithmetic(
        name          => 'Tie breaker',
        arguments     => { IV1 => $kr1, IV4 => $counter },
        arithmetic    => '=IV1*' . @values . '+IV4',
        defaultFormat => '0softnz'
    );

    my $kr = new SpreadsheetModel::Custom(
        name          => 'Ranking',
        defaultFormat => '0softnz',
        rows          => $valueSet,
        custom        => ['=RANK(IV1,IV2:IV3,1)'],
        arguments     => { IV1 => $kr2 },
        wsPrepare     => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0],
                  IV1 => xl_rowcol_to_cell( $rowh->{IV1} + $y, $colh->{IV1} ),
                  IV2 => xl_rowcol_to_cell( $rowh->{IV1}, $colh->{IV1}, 1, 0 ),
                  IV3 => xl_rowcol_to_cell( $rowh->{IV1} + $#values,
                    $colh->{IV1}, 1, 0 );
            };
        }
    );

    my $ror = new SpreadsheetModel::Custom(
        name          => 'Reordering',
        defaultFormat => '0softnz',
        rows          => $valueSet,
        custom        => ['=MATCH(IV1,IV2:IV3,0)'],
        arguments     => { IV1 => $counter, IV2 => $kr },
        wsPrepare     => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0],
                  IV1 => xl_rowcol_to_cell( $rowh->{IV1} + $y, $colh->{IV1} ),
                  IV2 => xl_rowcol_to_cell( $rowh->{IV2}, $colh->{IV2}, 1, 0 ),
                  IV3 => xl_rowcol_to_cell( $rowh->{IV2} + $#values,
                    $colh->{IV2}, 1, 0 );
            };
        }
    );

    my $kxs = new SpreadsheetModel::Custom(
        name      => 'Ordered values',
        rows      => $valueSet,
        custom    => [ '=INDEX(IV2:IV3,IV1,1)', '=IV2' ],
        arguments => { IV2 => $kx, IV1 => $ror },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0],
                  IV1 => xl_rowcol_to_cell( $rowh->{IV1} + $y, $colh->{IV1} ),
                  IV2 => xl_rowcol_to_cell( $rowh->{IV2}, $colh->{IV2}, 1, 0 ),
                  IV3 => xl_rowcol_to_cell( $rowh->{IV2} + $#values,
                    $colh->{IV2}, 1, 0 );
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
        arithmetic => '=1+IV1*(COUNTIF(IV2_IV3,"TRUE")-1)',
        arguments  => { IV1 => $self->{quantile}, IV2_IV3 => $cd }
      )
      : Arithmetic(
        name       => 'Target rank',
        arithmetic => '=1+IV1*(COUNT(IV2_IV3)-1)',
        arguments  => { IV1 => $self->{quantile}, IV2_IV3 => $kx }
      );

    my $hiti = Arithmetic(
        name       => 'Rank immediately below',
        arithmetic => '=FLOOR(IV1,1)',
        arguments  => { IV1 => $hit }
    );

    my $v1 = Arithmetic(
        name       => 'Value below',
        arithmetic => '=INDEX(IV1_IV2,IV3)',
        arguments  => { IV1_IV2 => $kxs, IV3 => $hiti }
    );

    my $v2 = Arithmetic(
        name       => 'Value above',
        arithmetic => '=INDEX(IV1_IV2,IV3+1)',
        arguments  => { IV1_IV2 => $kxs, IV3 => $hiti }
    );

    my $p1 = Arithmetic(
        defaultFormat => '%soft',
        name          => 'Weight below',
        arithmetic    => '=1-IV1+IV2',
        arguments     => { IV1 => $hit, IV2 => $hiti }
    );

    Columnset(
        name    => "Steps to calculate $self->{name} (part 2)",
        columns => [ $hit, $hiti, $v1, $v2, $p1 ]
    );

    $self->{arithmetic} = '=IV1*IV2+IF(IV5=1,0,(1-IV3)*IV4)';
    $self->{arguments}  = {
        IV1 => $p1,
        IV3 => $p1,
        IV5 => $p1,
        IV2 => $v1,
        IV4 => $v2,
    };

    $self->SUPER::check;

}

1;
