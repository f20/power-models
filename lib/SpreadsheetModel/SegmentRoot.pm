package SpreadsheetModel::SegmentRoot;

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

sub _getColumns {
    my ($obj) = @_;
    return unless $obj;
    return $obj if $obj->isa('SpreadsheetModel::Dataset');
    die "$obj not usable" unless $obj->isa('SpreadsheetModel::Columnset');
    @{ $obj->{columns} };
}

sub check {

    my ($self) = @_;

    return "$self->{name} $self->{debug}: "
      . "target $self->{target}{name} is not a single cell"
      unless ref $self->{target}
      && $self->{target}->isa('SpreadsheetModel::Dataset')
      && !$self->{target}->lastRow
      && !$self->{target}->lastCol;

    return "No slopes in $self->{name} $self->{debug}" unless $self->{slopes};

    my @slopes = _getColumns( $self->{slopes} );

    my $unconstrained = Arithmetic(
        name       => 'Constraint-free solution',
        arithmetic => '=A3/SUM('
          . join( ',', map { "A1${_}_A2$_" } 0 .. $#slopes ) . ')',
        arguments => {
            A3 => $self->{target},
            map { ( "A1${_}_A2$_" => $slopes[$_] ); } 0 .. $#slopes,
        },
    );

    my @min           = _getColumns( $self->{min} );
    my @max           = _getColumns( $self->{max} );
    my @minmax        = ( @min, @max );
    my $startingPoint = Arithmetic(
        name       => 'Starting point',
        arithmetic => '=MIN('
          . join( ',', A3 => map { "A1${_}_A2$_" } 0 .. $#minmax ) . ')',
        arguments => {
            A3 => $unconstrained,
            map { ( "A1${_}_A2$_" => $minmax[$_] ); } 0 .. $#minmax,
        },
    );

    my @kinks;
    for ( my $n = 0 ; $n < @minmax ; ++$n ) {
        foreach my $r ( 0 .. $minmax[$n]->lastRow ) {
            foreach my $c ( 0 .. $minmax[$n]->lastCol ) {
                push @kinks, [ $n, $r, $c ];
            }
        }
    }

    my $kinkSet =
      new SpreadsheetModel::Labelset(
        list => [ 'Starting point', map { "Kink $_" } 1 .. @kinks ] );

    my $kx = new SpreadsheetModel::Custom(
        name      => 'Location',
        rows      => $kinkSet,
        custom    => ['=A1'],      # assumes all on the same sheet
        arguments => {
            A1 => $startingPoint,
            map { ( "A2$_" => $minmax[$_] ); } 0 .. $#minmax,
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                return '', $format, $formula->[0],
                  qr/\bA1\b/ => xl_rowcol_to_cell( $rowh->{A1}, $colh->{A1} )
                  if !$y;
                my ( $n, $r, $c ) = @{ $kinks[ $y - 1 ] };
                return '', $format, $formula->[0],
                  qr/\bA1\b/ => xl_rowcol_to_cell( $rowh->{"A2$n"} + $r,
                    $colh->{"A2$n"} + $c );
            };
        }
    );

    my $kk = new SpreadsheetModel::Custom(
        name       => 'Kink',
        rows       => $kinkSet,
        custom     => [ '=A2', '=0-A2' ],      # assumes all on the same sheet
        arithmetic => 'Special calculation',
        arguments  => {
            A1 => $startingPoint,
            map { ( "A4$_" => $slopes[$_] ); } 0 .. $#slopes,
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                return '', $wb->getFormat('unavailable')
                  if !$y;
                my ( $n, $r, $c ) = @{ $kinks[ $y - 1 ] };
                my $m = $n % @slopes;
                '', $format, $formula->[ $n < @min ? 0 : 1 ],
                  qr/\bA2\b/ => xl_rowcol_to_cell( $rowh->{"A4$m"} + $r,
                    $colh->{"A4$m"} + $c );
            };
        }
    );

    my $startingSlope = new SpreadsheetModel::Custom(
        name   => 'Starting slope contributions',
        rows   => $kinkSet,
        custom => ['=IF(ISERROR(A1),A2,0)']
        ,    # ISERROR is true for #N/A and errors
        arguments => {
            A2 => $kk,
            A1 => $kx
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                return '', $wb->getFormat('unavailable')
                  if !$y;
                my ( $n, $r, $c ) = @{ $kinks[ $y - 1 ] };
                return '', $wb->getFormat('unavailable')
                  unless $n < @min;
                '', $format, $formula->[0], map {
                    qr/\b$_\b/ =>
                      xl_rowcol_to_cell( $rowh->{$_} + $y, $colh->{$_} + $x )
                } @$pha;
            };
        }
    );

    my $startingValue = new SpreadsheetModel::Custom(
        name      => 'Starting values',
        rows      => $kinkSet,
        custom    => ['=MAX(A3,A1)*A2'],
        arguments => {
            A2 => $kk,
            A1 => $kx,
            A3 => $startingPoint
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                return '', $wb->getFormat('unavailable')
                  if !$y;
                my ( $n, $r, $c ) = @{ $kinks[ $y - 1 ] };
                return '', $wb->getFormat('unavailable')
                  unless $n < @min;
                '', $format, $formula->[0],
                  qr/\bA3\b/ =>
                  xl_rowcol_to_cell( $rowh->{A3}, $colh->{A3}, 1, 1 ),
                  map {
                    qr/\b$_\b/ =>
                      xl_rowcol_to_cell( $rowh->{$_} + $y, $colh->{$_} + $x )
                  } qw(A1 A2);
            };
        }
    );

    my $counter = new SpreadsheetModel::Constant(
        name          => 'Counter',
        rows          => $kinkSet,
        data          => [ 0 .. @kinks ],
        defaultFormat => '0connz'
    );

    my $kr1 = new SpreadsheetModel::Custom(
        name          => 'Ranking before tie break',
        defaultFormat => '0softnz',
        rows          => $kinkSet,
        custom        => ['=RANK(A1,A2:A3,1)'],
        arguments     => {
            A1     => $kx,
            A2_A3 => $kx,
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }

# http://groups.google.com/group/spreadsheet-writeexcel/browse_thread/thread/6ba2e7e8e32fb21e

            sub {
                my ( $x, $y ) = @_;
                return '', $wb->getFormat('unavailable') unless $y;
                '', $format, $formula->[0],
                  qr/\bA1\b/ =>
                  xl_rowcol_to_cell( $rowh->{A1} + $y, $colh->{A1} ),
                  qr/\bA2\b/ =>
                  xl_rowcol_to_cell( $rowh->{A1} + 1, $colh->{A1}, 1, 0 ),
                  qr/\bA3\b/ =>
                  xl_rowcol_to_cell( $rowh->{A1} + @kinks, $colh->{A1}, 1, 0 );
            };
        }
    );

    my $kr2 = new SpreadsheetModel::Arithmetic(
        name          => 'Tie breaker',
        arguments     => { A1 => $kr1, A4 => $counter },
        arithmetic    => '=A1*' . @kinks . '+A4',
        defaultFormat => '0softnz'
    );

    my $kr = new SpreadsheetModel::Custom(
        name          => 'Ranking',
        defaultFormat => '0softnz',
        rows          => $kinkSet,
        custom        => ['=RANK(A1,A2:A3,1)'],
        arguments     => {
            A1     => $kr2,
            A2_A3 => $kr2,
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }
            sub {
                my ( $x, $y ) = @_;
                return '', $wb->getFormat('unavailable') unless $y;
                '', $format, $formula->[0],
                  qr/\bA1\b/ =>
                  xl_rowcol_to_cell( $rowh->{A1} + $y, $colh->{A1} ),
                  qr/\bA2\b/ =>
                  xl_rowcol_to_cell( $rowh->{A1} + 1, $colh->{A1}, 1, 0 ),
                  qr/\bA3\b/ =>
                  xl_rowcol_to_cell( $rowh->{A1} + @kinks, $colh->{A1}, 1, 0 );
            };
        }
    );

    my $ror = new SpreadsheetModel::Custom(
        name          => 'Kink reordering',
        defaultFormat => '0softnz',
        rows          => $kinkSet,
        custom        => ['=MATCH(A1,A2:A3,0)'],
        arguments     => {
            A1     => $counter,
            A2     => $kr,
            A2_A3 => $kr,
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }
            sub {
                my ( $x, $y ) = @_;
                return '', $wb->getFormat('unavailable') unless $y;
                '', $format, $formula->[0],
                  qr/\bA1\b/ =>
                  xl_rowcol_to_cell( $rowh->{A1} + $y, $colh->{A1} ),
                  qr/\bA2\b/ =>
                  xl_rowcol_to_cell( $rowh->{A2} + 1, $colh->{A2}, 1, 0 ),
                  qr/\bA3\b/ =>
                  xl_rowcol_to_cell( $rowh->{A2} + @kinks, $colh->{A2}, 1, 0 );
            };
        }
    );

    my $kxs = new SpreadsheetModel::Custom(
        name      => 'Location (ordered)',
        rows      => $kinkSet,
        custom    => [ '=INDEX(A2:A3,A1,1)', '=A2' ],
        arguments => {
            A2     => $kx,
            A1     => $ror,
            A2_A3 => $kx,
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }
            sub {
                my ( $x, $y ) = @_;
                return '', $format, $formula->[1],
                  qr/\bA2\b/ => xl_rowcol_to_cell( $rowh->{A2}, $colh->{A2} )
                  unless $y;
                '', $format, $formula->[0],
                  qr/\bA1\b/ =>
                  xl_rowcol_to_cell( $rowh->{A1} + $y, $colh->{A1} ),
                  qr/\bA2\b/ =>
                  xl_rowcol_to_cell( $rowh->{A2} + 1, $colh->{A2}, 1, 0 ),
                  qr/\bA3\b/ =>
                  xl_rowcol_to_cell( $rowh->{A2} + @kinks, $colh->{A2}, 1, 0 );
            };
        }
    );

    my $kks = new SpreadsheetModel::Custom(
        name       => 'New slope',
        rows       => $kinkSet,
        custom     => [ '=A7+INDEX(A2:A3,A1,1)', '=SUM(A5:A6)' ],
        arithmetic => 'Special calculation',
        arguments  => {
            A2 => $kk,
            A1 => $ror,
            A5 => $startingSlope,
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }
            sub {
                my ( $x, $y ) = @_;
                return '', $format, $formula->[1],
                  qr/\bA5\b/ =>
                  xl_rowcol_to_cell( $rowh->{A5}, $colh->{A5}, 1, 0 ),
                  qr/\bA6\b/ => xl_rowcol_to_cell( $rowh->{A5} + @kinks - 1,
                    $colh->{A5}, 1, 0 )
                  unless $y;
                '', $format, $formula->[0],
                  qr/\bA1\b/ =>
                  xl_rowcol_to_cell( $rowh->{A1} + $y, $colh->{A1} ),
                  qr/\bA2\b/ =>
                  xl_rowcol_to_cell( $rowh->{A2} + 1, $colh->{A2}, 1, 0 ),
                  qr/\bA3\b/ =>
                  xl_rowcol_to_cell( $rowh->{A2} + @kinks, $colh->{A2}, 1, 0 ),
                  qr/\bA7\b/ => xl_rowcol_to_cell( $me->{$wb}{row} + $y - 1,
                    $me->{$wb}{col} );
            };
        }
    );

    my $kvs = new SpreadsheetModel::Custom(
        name       => 'Value',
        rows       => $kinkSet,
        custom     => [ '=A7+(A4-A3)*A2', '=SUM(A5:A6)-A9' ],
        arithmetic => 'Special calculation',
        arguments  => {
            A4 => $kxs,
            A2 => $kks,
            A5 => $startingValue,
            A9 => $self->{target},
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                return '', $format, $formula->[1],
                  qr/\bA9\b/ =>
                  xl_rowcol_to_cell( $rowh->{A9}, $colh->{A9}, 1, 1 ),
                  qr/\bA5\b/ =>
                  xl_rowcol_to_cell( $rowh->{A5}, $colh->{A5}, 1, 1 ),
                  qr/\bA6\b/ => xl_rowcol_to_cell( $rowh->{A5} + @kinks - 1,
                    $colh->{A5}, 1, 1 )
                  unless $y;
                '', $format, $formula->[0],
                  qr/\bA2\b/ =>
                  xl_rowcol_to_cell( $rowh->{A2} + $y - 1, $colh->{A2} ),
                  qr/\bA3\b/ =>
                  xl_rowcol_to_cell( $rowh->{A4} + $y - 1, $colh->{A4} ),
                  qr/\bA4\b/ =>
                  xl_rowcol_to_cell( $rowh->{A4} + $y, $colh->{A4} ),
                  qr/\bA7\b/ => xl_rowcol_to_cell( $me->{$wb}{row} + $y - 1,
                    $me->{$wb}{col} );
            };
        }
    );

    my $root = new SpreadsheetModel::Custom(
        name => 'Root',
        rows => $kinkSet,
        custom =>
          [ '=IF((A2>0)=(A3>0),"",A1-A9/A4)', '=IF(A2>0,A1,IF(A3>0,"",A5))' ],
        arithmetic => 'Special calculation',
        arguments  => {
            A9 => $kvs,
            A4 => $kks,
            A1 => $kxs,
            A5 => $unconstrained,
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;

                return '', $format, $formula->[1],
                  qr/\bA1\b/ => xl_rowcol_to_cell( $rowh->{A1}, $colh->{A1} ),
                  qr/\bA5\b/ =>
                  xl_rowcol_to_cell( $rowh->{A5}, $colh->{A5}, 1, 1 ),
                  qr/\bA2\b/ =>
                  xl_rowcol_to_cell( $rowh->{A9}, $colh->{A9}, 1, 0 ),
                  qr/\bA3\b/ => xl_rowcol_to_cell( $rowh->{A9} + $kxs->lastRow,
                    $colh->{A9}, 1, 0 )
                  unless $y;

                '', $format, $formula->[0],
                  qr/\bA1\b/ =>
                  xl_rowcol_to_cell( $rowh->{A1} + $y, $colh->{A1} ),
                  qr/\bA2\b/ =>
                  xl_rowcol_to_cell( $rowh->{A9} + $y - 1, $colh->{A9} ),
                  qr/\bA3\b/ =>
                  xl_rowcol_to_cell( $rowh->{A9} + $y, $colh->{A9} ),
                  qr/\bA9\b/ =>
                  xl_rowcol_to_cell( $rowh->{A9} + $y, $colh->{A9} ),
                  qr/\bA4\b/ =>
                  xl_rowcol_to_cell( $rowh->{A4} + $y - 1, $colh->{A4} );

            };
        }
    );

    Columnset(
        name    => "Solve for $self->{name}",
        columns => [
            $kx, $kk, $startingSlope, $startingValue, $kr1, $counter, $kr2, $kr,
            $ror, $kxs, $kks, $kvs, $root
        ]
    );

    $self->{arithmetic} = '=MIN(A1_A2)';
    $self->{arguments} = { A1_A2 => $root };

    $self->SUPER::check;

}

1;
