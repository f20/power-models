package SpreadsheetModel::SegmentRoot;

=head Copyright licence and disclaimer

Copyright 2008-2013 Reckon LLP and others. All rights reserved.

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

# Forward links bug:
# This class does not work correctly with forwardLinks features because
# tables are being created within the wsPrepare call.

use warnings;
use strict;

require SpreadsheetModel::Dataset;
our @ISA = qw(SpreadsheetModel::Dataset);

use SpreadsheetModel::Miscellaneous;
use Spreadsheet::WriteExcel::Utility;

sub objectType {
    'Optimisation result';
}

sub check {
    my ($self) = @_;
    return "Target $self->{target}{name} is no good in $self->{debug}"
      unless $self->{target}
      && $self->{target}->isa('SpreadsheetModel::Dataset')
      && !$self->{target}->lastRow
      && !$self->{target}->lastCol;
    return "No slopes in $self->{debug}" unless $self->{slopes};
    $self->{arithmetic} = 'Special calculation';
    $self->{arguments}  = {};
    $self->SUPER::check;
}

sub _getRectangle {
    my ( $source, $wb, $ws ) = @_;
    return unless $source;

    my ( $slopsheet, $slopr1, $slopc1, $slopr2, $slopc2 );

    if ( $source->isa('SpreadsheetModel::Dataset') ) {
        ( $slopsheet, $slopr1, $slopc1 ) = $source->wsWrite( $wb, $ws );
        $slopr2 = $slopr1 + $source->lastRow;
        $slopc2 = $slopc1 + $source->lastCol;
    }
    else {
        my $lastColumn = $source->{columns}[ $#{ $source->{columns} } ];
        ( $slopsheet, $slopr1, $slopc1 ) =
          $source->{columns}[0]->wsWrite( $wb, $ws );
        ( $slopsheet, $slopr2, $slopc2 ) = $lastColumn->wsWrite( $wb, $ws );
        $slopr2 += $lastColumn->lastRow;
        $slopc2 += $lastColumn->lastCol;
    }
    $slopsheet = $slopsheet == $ws ? '' : "'" . $slopsheet->get_name . "'!";
    $slopsheet, $slopr1, $slopc1, $slopr2, $slopc2;

}

sub wsPrepare {

    my ( $self, $wb, $ws ) = @_;

    my ( $targsheet, $targr, $targc ) = $self->{target}->wsWrite( $wb, $ws );
    $targsheet = $targsheet == $ws ? '' : "'" . $targsheet->get_name . "'!";

    my ( $slopsheet, $slopr1, $slopc1, $slopr2, $slopc2 ) =
      _getRectangle( $self->{slopes}, $wb, $ws );
    my ( $minsheet, $minr1, $minc1, $minr2, $minc2 ) =
      _getRectangle( $self->{min}, $wb, $ws );
    my ( $maxsheet, $maxr1, $maxc1, $maxr2, $maxc2 ) =
      _getRectangle( $self->{max}, $wb, $ws );

    my $unconstrained = new SpreadsheetModel::Custom(
        name        => 'Constraint-free solution',
        sourceLines => [
            $self->{target},
            $self->{slopes}->isa('SpreadsheetModel::Dataset')
            ? $self->{slopes}
            : @{ $self->{slopes}{columns} }
        ],
        custom    => ["=${targsheet}IV3/SUM(${slopsheet}IV1:IV2)"],
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula ) = @_;
            sub {
                '', $format, $formula->[0],
                  IV3 => xl_rowcol_to_cell( $targr,  $targc ),
                  IV1 => xl_rowcol_to_cell( $slopr1, $slopc1 ),
                  IV2 => xl_rowcol_to_cell( $slopr2, $slopc2 );
            };
        }
    );
    my ( $unconsheet, $unconr, $unconc ) = $unconstrained->wsWrite( $wb, $ws );
    $unconsheet = $unconsheet == $ws ? '' : "'" . $unconsheet->get_name . "'!";

    my $startingPoint = new SpreadsheetModel::Custom(
        name        => 'Starting point',
        sourceLines => [
            $unconstrained,
            $self->{min}
            ? (
                  $self->{min}->isa('SpreadsheetModel::Dataset')
                ? $self->{min}
                : @{ $self->{min}{columns} }
              )
            : (),
            $self->{max}
            ? (
                  $self->{max}->isa('SpreadsheetModel::Dataset')
                ? $self->{max}
                : @{ $self->{max}{columns} }
              )
            : ()
        ],
        custom => [
                "=MIN(${unconsheet}IV3"
              . ( $self->{min} ? ",${minsheet}IV4:IV5" : '' )
              . ( $self->{max} ? ",${maxsheet}IV6:IV7" : '' ) . ')'
        ],
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula ) = @_;
            sub {
                '', $format, $formula->[0],
                  IV3 => xl_rowcol_to_cell( $unconr, $unconc ),
                  IV1 => xl_rowcol_to_cell( $slopr1, $slopc1 ),
                  IV2 => xl_rowcol_to_cell( $slopr2, $slopc2 ),
                  $self->{min}
                  ? (
                    IV4 => xl_rowcol_to_cell( $minr1, $minc1 ),
                    IV5 => xl_rowcol_to_cell( $minr2, $minc2 )
                  )
                  : (), $self->{max} ? (
                    IV6 => xl_rowcol_to_cell( $maxr1, $maxc1 ),
                    IV7 => xl_rowcol_to_cell( $maxr2, $maxc2 )
                  )
                  : ();
            };
        }
    );

    0 and new SpreadsheetModel::Columnset(
        name    => "Prepare to solve for $self->{name}",
        columns => [ $unconstrained, $startingPoint ]
    );

    my $numRows = $slopr2 - $slopr1 + 1;
    my $numCols = $slopc2 - $slopc1 + 1;
    my $numKinks =
      ( ( $self->{min} ? 1 : 0 ) + ( $self->{max} ? 1 : 0 ) ) *
      $numRows *
      $numCols;
    my $kinkSet =
      new SpreadsheetModel::Labelset(
        list => [ 'Starting point', map { "Kink $_" } 1 .. $numKinks ] );

    my $kx = new SpreadsheetModel::Custom(
        name        => 'Location',
        sourceLines => [
            $startingPoint,
            $self->{min}
            ? (
                  $self->{min}->isa('SpreadsheetModel::Dataset')
                ? $self->{min}
                : @{ $self->{min}{columns} }
              )
            : (),
            $self->{max}
            ? (
                  $self->{max}->isa('SpreadsheetModel::Dataset')
                ? $self->{max}
                : @{ $self->{max}{columns} }
              )
            : ()
        ],
        rows      => $kinkSet,
        custom    => [ '=IV1', '=IV2', '=IV3' ],
        arguments => {
            IV1 => $startingPoint,
            $self->{min}
            ? (
                IV2 => $self->{min}{columns}
                ? $self->{min}{columns}[0]
                : $self->{min}
              )
            : (),
            $self->{max}
            ? (
                IV3 => $self->{max}{columns}
                ? $self->{max}{columns}[0]
                : $self->{max}
              )
            : ()
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                return '', $format, $formula->[0],
                  IV1 => xl_rowcol_to_cell( $rowh->{IV1}, $colh->{IV1} )
                  if !$y;
                --$y;
                return '', $format, $formula->[1], IV2 => xl_rowcol_to_cell(
                    $rowh->{IV2} + $y % $numRows,
                    $colh->{IV2} + int( $y / $numRows )
                ) if $self->{min} && ( !$self->{max} || $y * 2 < $numKinks );
                return '', $format, $formula->[2], IV3 => xl_rowcol_to_cell(
                    $rowh->{IV3} + $y % $numRows,
                    $colh->{IV3} + int( $y / $numRows ) - (
                          $self->{min}
                        ? $numCols
                        : 0
                    )
                );
            };
        }
    );

    my $kk = new SpreadsheetModel::Custom(
        name        => 'Kink',
        sourceLines => [
              $self->{slopes}->isa('SpreadsheetModel::Dataset')
            ? $self->{slopes}
            : @{ $self->{slopes}{columns} }
        ],
        rows      => $kinkSet,
        custom    => [ '=IV2', '=0-IV3' ],
        arguments => {
            $self->{min}
            ? (
                IV2 => $self->{slopes}{columns}
                ? $self->{slopes}{columns}[0]
                : $self->{slopes}
              )
            : (),
            $self->{max}
            ? (
                IV3 => $self->{slopes}{columns}
                ? $self->{slopes}{columns}[0]
                : $self->{slopes}
              )
            : ()
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                return '', $wb->getFormat('unavailable')
                  if !$y;
                --$y;
                return '', $format, $formula->[0], IV2 => xl_rowcol_to_cell(
                    $rowh->{IV2} + $y % $numRows,
                    $colh->{IV2} + int( $y / $numRows )
                ) if $self->{min} && ( !$self->{max} || $y * 2 < $numKinks );
                return '', $format, $formula->[1], IV3 => xl_rowcol_to_cell(
                    $rowh->{IV3} + $y % $numRows,
                    $colh->{IV3} + int( $y / $numRows ) - (
                          $self->{min}
                        ? $numCols
                        : 0
                    )
                );
            };
        }
    );

    my $startingSlope = new SpreadsheetModel::Custom(
        name        => 'Starting slopes',
        sourceLines => [ $kk, $kx ],
        rows        => $kinkSet,
        custom      => ['=IF(ISNUMBER(IV1),0,IV2)'],
        arguments   => {
            IV2 => $kk,
            IV1 => $kx
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                return '', $wb->getFormat('unavailable')
                  if !$y
                  || !$self->{min}
                  || $self->{max} && $y * 2 >= $numKinks;
                '', $format, $formula->[0], map {
                    $_ =>
                      xl_rowcol_to_cell( $rowh->{$_} + $y, $colh->{$_} + $x )
                } @$pha;
            };
        }
    );

    my $startingValue = new SpreadsheetModel::Custom(
        name        => 'Starting values',
        sourceLines => [ $kk, $kx, $startingPoint ],
        rows        => $kinkSet,
        custom      => ['=MAX(IV3,IV1)*IV2'],
        arguments   => {
            IV2 => $kk,
            IV1 => $kx,
            IV3 => $startingPoint
        },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                return '', $wb->getFormat('unavailable')
                  if !$y
                  || !$self->{min}
                  || $self->{max} && $y * 2 >= $numKinks;
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
        data          => [ 0 .. $numKinks ],
        defaultFormat => '0connz'
    );

    my $kr1 = new SpreadsheetModel::Custom(
        name          => 'Ranking before tie break',
        sourceLines   => [$kx],
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
                0
                  and return 1, $format,
                  '=RANK('
                  . xl_rowcol_to_cell( $rowh->{IV1} + $y, $colh->{IV1} ) . ','
                  . xl_rowcol_to_cell( $rowh->{IV1} + 1, $colh->{IV1}, 1, 0 )
                  . ':'
                  . xl_rowcol_to_cell( $rowh->{IV1} + $numKinks,
                    $colh->{IV1}, 1, 0 )
                  . ',1)';
                '', $format, $formula->[0],
                  IV1 => xl_rowcol_to_cell( $rowh->{IV1} + $y, $colh->{IV1} ),
                  IV2 =>
                  xl_rowcol_to_cell( $rowh->{IV1} + 1, $colh->{IV1}, 1, 0 ),
                  IV3 => xl_rowcol_to_cell( $rowh->{IV1} + $numKinks,
                    $colh->{IV1}, 1, 0 );
            };
        }
    );

    my $kr2 = new SpreadsheetModel::Arithmetic(
        name          => 'Tie breaker',
        sourceLines   => [ $kr1, $counter ],
        arguments     => { IV1 => $kr1, IV4 => $counter },
        arithmetic    => "=IV1*$numKinks+IV4",
        defaultFormat => '0softnz'
    );

    my $kr = new SpreadsheetModel::Custom(
        name          => 'Ranking',
        sourceLines   => [$kr2],
        defaultFormat => '0softnz',
        rows          => $kinkSet,
        custom        => ['=RANK(IV1,IV2:IV3,1)'],
        arguments     => { IV1 => $kr2 },
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
                  IV3 => xl_rowcol_to_cell( $rowh->{IV1} + $numKinks,
                    $colh->{IV1}, 1, 0 );
            };
        }
    );

    my $ror = new SpreadsheetModel::Custom(
        name          => 'Kink reordering',
        sourceLines   => [ $kr, $counter ],
        defaultFormat => '0softnz',
        rows          => $kinkSet,
        custom        => ['=MATCH(IV1,IV2:IV3,0)'],
        arguments     => { IV1 => $counter, IV2 => $kr },
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
                  xl_rowcol_to_cell( $rowh->{IV2} + 1, $colh->{IV2}, 1, 0 ),
                  IV3 => xl_rowcol_to_cell( $rowh->{IV2} + $numKinks,
                    $colh->{IV2}, 1, 0 );
            };
        }
    );

    my $kxs = new SpreadsheetModel::Custom(
        name        => 'Location (ordered)',
        sourceLines => [ $kx, $ror ],
        rows        => $kinkSet,
        custom      => [ '=INDEX(IV2:IV3,IV1,1)', '=IV2' ],
        arguments   => { IV2 => $kx, IV1 => $ror },
        wsPrepare   => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }

# http://groups.google.com/group/spreadsheet-writeexcel/browse_thread/thread/6ba2e7e8e32fb21e
            sub {
                my ( $x, $y ) = @_;
                return '', $format, $formula->[1],
                  IV2 => xl_rowcol_to_cell( $rowh->{IV2}, $colh->{IV2} )
                  unless $y;
                '', $format, $formula->[0],
                  IV1 => xl_rowcol_to_cell( $rowh->{IV1} + $y, $colh->{IV1} ),
                  IV2 =>
                  xl_rowcol_to_cell( $rowh->{IV2} + 1, $colh->{IV2}, 1, 0 ),
                  IV3 => xl_rowcol_to_cell( $rowh->{IV2} + $numKinks,
                    $colh->{IV2}, 1, 0 );
            };
        }
    );

    my $kks = new SpreadsheetModel::Custom(
        name        => 'New slope',
        sourceLines => [ $kk, $ror, $startingSlope ],
        rows        => $kinkSet,
        custom => [ '=IV7+INDEX(IV2:IV3,IV1,1)', '=SUM(IV5:IV6)' ],
        arguments => { IV2 => $kk, IV1 => $ror, IV5 => $startingSlope },
        wsPrepare => sub {
            my ( $me, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;

            foreach (@$formula) { s/_ref2d/_ref2dV/ foreach @$_; }

# http://groups.google.com/group/spreadsheet-writeexcel/browse_thread/thread/6ba2e7e8e32fb21e

            sub {
                my ( $x, $y ) = @_;
                return '', $format, $formula->[1],
                  IV5 => xl_rowcol_to_cell( $rowh->{IV5}, $colh->{IV5}, 1, 0 ),
                  IV6 => xl_rowcol_to_cell( $rowh->{IV5} + $numKinks - 1,
                    $colh->{IV5}, 1, 0 )
                  unless $y;
                '', $format, $formula->[0],
                  IV1 => xl_rowcol_to_cell( $rowh->{IV1} + $y, $colh->{IV1} ),
                  IV2 =>
                  xl_rowcol_to_cell( $rowh->{IV2} + 1, $colh->{IV2}, 1, 0 ),
                  IV3 => xl_rowcol_to_cell( $rowh->{IV2} + $numKinks,
                    $colh->{IV2}, 1, 0 ),
                  IV7 => xl_rowcol_to_cell( $me->{$wb}{row} + $y - 1,
                    $me->{$wb}{col} );
            };
        }
    );

    my $kvs = new SpreadsheetModel::Custom(
        name        => 'Value',
        sourceLines => [ $kxs, $kks, $startingValue, $self->{target} ],
        rows        => $kinkSet,
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
                  IV6 => xl_rowcol_to_cell( $rowh->{IV5} + $numKinks - 1,
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
        name        => 'Root',
        sourceLines => [ $kvs, $kks, $kxs, $unconstrained ],
        rows        => $kinkSet,
        custom      => [
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

    new SpreadsheetModel::Columnset(
        name    => "Solve for $self->{name}",
        columns => [
            $kx, $kk, $startingSlope, $startingValue, $kr1, $counter, $kr2, $kr,
            $ror, $kxs, $kks, $kvs, $root
        ]
    );

    my ( $rootsheet, $rootr, $rootc ) = $root->wsWrite( $wb, $ws );
    $rootsheet = $rootsheet == $ws ? '' : "'" . $rootsheet->get_name . "'!";

    my $formula = $ws->store_formula("=MIN(${rootsheet}IV1:IV2)");
    my $format = $wb->getFormat( $self->{defaultFormat} || '0.000soft' );

    $self->{sourceLines} = [$root];
    $self->{arithmetic}  = '=MIN(IV1)';
    $self->{arguments}   = { IV1 => $root };

    sub {
        '', $format, $formula,
          IV1 => xl_rowcol_to_cell( $rootr,                  $rootc ),
          IV2 => xl_rowcol_to_cell( $rootr + $root->lastRow, $rootc ),
          ;
    };

}

1;
