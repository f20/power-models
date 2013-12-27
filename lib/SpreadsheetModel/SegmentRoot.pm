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
        arithmetic => '=IV3/SUM('
          . join( ',', map { "IV1${_}_IV2$_" } 0 .. $#slopes ) . ')',
        arguments => {
            IV3 => $self->{target},
            map { ( "IV1${_}_IV2$_" => $slopes[$_] ); } 0 .. $#slopes,
        },
    );

    my @min           = _getColumns( $self->{min} );
    my @max           = _getColumns( $self->{max} );
    my @minmax        = ( @min, @max );
    my $startingPoint = Arithmetic(
        name       => 'Starting point',
        arithmetic => '=MIN('
          . join( ',', IV3 => map { "IV1${_}_IV2$_" } 0 .. $#minmax ) . ')',
        arguments => {
            IV3 => $unconstrained,
            map { ( "IV1${_}_IV2$_" => $minmax[$_] ); } 0 .. $#minmax,
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
        custom    => ['=IV1'],     # assumes all on the same sheet
        arguments => {
            IV1 => $startingPoint,
            map { ( "IV2$_" => $minmax[$_] ); } 0 .. $#minmax,
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                return '', $format, $formula->[0],
                  IV1 => xl_rowcol_to_cell( $rowh->{IV1}, $colh->{IV1} )
                  if !$y;
                my ( $n, $r, $c ) = @{ $kinks[ $y - 1 ] };
                return '', $format, $formula->[0],
                  IV1 => xl_rowcol_to_cell( $rowh->{"IV2$n"} + $r,
                    $colh->{"IV2$n"} + $c );
            };
        }
    );

    my $kk = new SpreadsheetModel::Custom(
        name      => 'Kink',
        rows      => $kinkSet,
        custom    => [ '=IV2', '=0-IV2' ],    # assumes all on the same sheet
        arguments => {
            IV1 => $startingPoint,
            map { ( "IV4$_" => $slopes[$_] ); } 0 .. $#slopes,
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
                  IV2 => xl_rowcol_to_cell( $rowh->{"IV4$m"} + $r,
                    $colh->{"IV4$m"} + $c );
            };
        }
    );

    my $startingSlope = new SpreadsheetModel::Custom(
        name   => 'Starting slope contributions',
        rows   => $kinkSet,
        custom => ['=IF(ISERROR(IV1),IV2,0)']
        ,    # ISERROR is true for #N/A and errors
        arguments => {
            IV2 => $kk,
            IV1 => $kx
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
                    $_ =>
                      xl_rowcol_to_cell( $rowh->{$_} + $y, $colh->{$_} + $x )
                } @$pha;
            };
        }
    );

    my $startingValue = new SpreadsheetModel::Custom(
        name      => 'Starting values',
        rows      => $kinkSet,
        custom    => ['=MAX(IV3,IV1)*IV2'],
        arguments => {
            IV2 => $kk,
            IV1 => $kx,
            IV3 => $startingPoint
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
                  IV3 => xl_rowcol_to_cell( $rowh->{IV3}, $colh->{IV3}, 1, 1 ),
                  map {
                    $_ =>
                      xl_rowcol_to_cell( $rowh->{$_} + $y, $colh->{$_} + $x )
                  } qw(IV1 IV2);
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
        custom        => ['=RANK(IV1,IV2:IV3,1)'],
        arguments     => { IV1 => $kx },
        wsPrepare     => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }

# http://groups.google.com/group/spreadsheet-writeexcel/browse_thread/thread/6ba2e7e8e32fb21e

            sub {
                my ( $x, $y ) = @_;
                return '', $wb->getFormat('unavailable') unless $y;
                '', $format, $formula->[0],
                  IV1 => xl_rowcol_to_cell( $rowh->{IV1} + $y, $colh->{IV1} ),
                  IV2 =>
                  xl_rowcol_to_cell( $rowh->{IV1} + 1, $colh->{IV1}, 1, 0 ),
                  IV3 => xl_rowcol_to_cell( $rowh->{IV1} + @kinks,
                    $colh->{IV1}, 1, 0 );
            };
        }
    );

    my $kr2 = new SpreadsheetModel::Arithmetic(
        name          => 'Tie breaker',
        arguments     => { IV1 => $kr1, IV4 => $counter },
        arithmetic    => '=IV1*' . @kinks . '+IV4',
        defaultFormat => '0softnz'
    );

    my $kr = new SpreadsheetModel::Custom(
        name          => 'Ranking',
        defaultFormat => '0softnz',
        rows          => $kinkSet,
        custom        => ['=RANK(IV1,IV2:IV3,1)'],
        arguments     => { IV1 => $kr2 },
        wsPrepare     => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }
            sub {
                my ( $x, $y ) = @_;
                return '', $wb->getFormat('unavailable') unless $y;
                '', $format, $formula->[0],
                  IV1 => xl_rowcol_to_cell( $rowh->{IV1} + $y, $colh->{IV1} ),
                  IV2 =>
                  xl_rowcol_to_cell( $rowh->{IV1} + 1, $colh->{IV1}, 1, 0 ),
                  IV3 => xl_rowcol_to_cell( $rowh->{IV1} + @kinks,
                    $colh->{IV1}, 1, 0 );
            };
        }
    );

    my $ror = new SpreadsheetModel::Custom(
        name          => 'Kink reordering',
        defaultFormat => '0softnz',
        rows          => $kinkSet,
        custom        => ['=MATCH(IV1,IV2:IV3,0)'],
        arguments     => { IV1 => $counter, IV2 => $kr },
        wsPrepare     => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }
            sub {
                my ( $x, $y ) = @_;
                return '', $wb->getFormat('unavailable') unless $y;
                '', $format, $formula->[0],
                  IV1 => xl_rowcol_to_cell( $rowh->{IV1} + $y, $colh->{IV1} ),
                  IV2 =>
                  xl_rowcol_to_cell( $rowh->{IV2} + 1, $colh->{IV2}, 1, 0 ),
                  IV3 => xl_rowcol_to_cell( $rowh->{IV2} + @kinks,
                    $colh->{IV2}, 1, 0 );
            };
        }
    );

    my $kxs = new SpreadsheetModel::Custom(
        name      => 'Location (ordered)',
        rows      => $kinkSet,
        custom    => [ '=INDEX(IV2:IV3,IV1,1)', '=IV2' ],
        arguments => { IV2 => $kx, IV1 => $ror },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }
            sub {
                my ( $x, $y ) = @_;
                return '', $format, $formula->[1],
                  IV2 => xl_rowcol_to_cell( $rowh->{IV2}, $colh->{IV2} )
                  unless $y;
                '', $format, $formula->[0],
                  IV1 => xl_rowcol_to_cell( $rowh->{IV1} + $y, $colh->{IV1} ),
                  IV2 =>
                  xl_rowcol_to_cell( $rowh->{IV2} + 1, $colh->{IV2}, 1, 0 ),
                  IV3 => xl_rowcol_to_cell( $rowh->{IV2} + @kinks,
                    $colh->{IV2}, 1, 0 );
            };
        }
    );

    my $kks = new SpreadsheetModel::Custom(
        name      => 'New slope',
        rows      => $kinkSet,
        custom    => [ '=IV7+INDEX(IV2:IV3,IV1,1)', '=SUM(IV5:IV6)' ],
        arguments => { IV2 => $kk, IV1 => $ror, IV5 => $startingSlope },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }
            sub {
                my ( $x, $y ) = @_;
                return '', $format, $formula->[1],
                  IV5 => xl_rowcol_to_cell( $rowh->{IV5}, $colh->{IV5}, 1, 0 ),
                  IV6 => xl_rowcol_to_cell( $rowh->{IV5} + @kinks - 1,
                    $colh->{IV5}, 1, 0 )
                  unless $y;
                '', $format, $formula->[0],
                  IV1 => xl_rowcol_to_cell( $rowh->{IV1} + $y, $colh->{IV1} ),
                  IV2 =>
                  xl_rowcol_to_cell( $rowh->{IV2} + 1, $colh->{IV2}, 1, 0 ),
                  IV3 => xl_rowcol_to_cell( $rowh->{IV2} + @kinks,
                    $colh->{IV2}, 1, 0 ),
                  IV7 => xl_rowcol_to_cell( $me->{$wb}{row} + $y - 1,
                    $me->{$wb}{col} );
            };
        }
    );

    my $kvs = new SpreadsheetModel::Custom(
        name      => 'Value',
        rows      => $kinkSet,
        custom    => [ '=IV7+(IV4-IV3)*IV2', '=SUM(IV5:IV6)-IV9' ],
        arguments => {
            IV4 => $kxs,
            IV2 => $kks,
            IV5 => $startingValue,
            IV9 => $self->{target}
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                return '', $format, $formula->[1],
                  IV9 => xl_rowcol_to_cell( $rowh->{IV9}, $colh->{IV9}, 1, 1 ),
                  IV5 => xl_rowcol_to_cell( $rowh->{IV5}, $colh->{IV5}, 1, 1 ),
                  IV6 => xl_rowcol_to_cell( $rowh->{IV5} + @kinks - 1,
                    $colh->{IV5}, 1, 1 )
                  unless $y;
                '', $format, $formula->[0],
                  IV2 =>
                  xl_rowcol_to_cell( $rowh->{IV2} + $y - 1, $colh->{IV2} ),
                  IV3 =>
                  xl_rowcol_to_cell( $rowh->{IV4} + $y - 1, $colh->{IV4} ),
                  IV4 => xl_rowcol_to_cell( $rowh->{IV4} + $y, $colh->{IV4} ),
                  IV7 => xl_rowcol_to_cell( $me->{$wb}{row} + $y - 1,
                    $me->{$wb}{col} );
            };
        }
    );

    my $root = new SpreadsheetModel::Custom(
        name   => 'Root',
        rows   => $kinkSet,
        custom => [
            '=IF((IV2>0)=(IV3>0),"",IV1-IV9/IV4)',
            '=IF(IV2>0,IV1,IF(IV3>0,"",IV5))'
        ],
        arguments =>
          { IV9 => $kvs, IV4 => $kks, IV1 => $kxs, IV5 => $unconstrained },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;

                return '', $format, $formula->[1],
                  IV1 => xl_rowcol_to_cell( $rowh->{IV1}, $colh->{IV1} ),
                  IV5 => xl_rowcol_to_cell( $rowh->{IV5}, $colh->{IV5}, 1, 1 ),
                  IV2 => xl_rowcol_to_cell( $rowh->{IV9}, $colh->{IV9}, 1, 0 ),
                  IV3 => xl_rowcol_to_cell( $rowh->{IV9} + $kxs->lastRow,
                    $colh->{IV9}, 1, 0 )
                  unless $y;

                '', $format, $formula->[0],
                  IV1 => xl_rowcol_to_cell( $rowh->{IV1} + $y, $colh->{IV1} ),
                  IV2 =>
                  xl_rowcol_to_cell( $rowh->{IV9} + $y - 1, $colh->{IV9} ),
                  IV3 => xl_rowcol_to_cell( $rowh->{IV9} + $y, $colh->{IV9} ),
                  IV9 => xl_rowcol_to_cell( $rowh->{IV9} + $y, $colh->{IV9} ),
                  IV4 =>
                  xl_rowcol_to_cell( $rowh->{IV4} + $y - 1, $colh->{IV4} );

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

    $self->{arithmetic} = '=MIN(IV1_IV2)';
    $self->{arguments} = { IV1_IV2 => $root };

    $self->SUPER::check;

}

1;
