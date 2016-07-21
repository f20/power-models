package SpreadsheetModel::MatrixSheet;

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

# Avoid using this container for complex calculations
# as it breaks the golden rule of sensible model ordering.

use warnings;
use strict;
use utf8;

use Spreadsheet::WriteExcel::Utility;
use SpreadsheetModel::Object ':_util';

sub new {
    my ( $class, %pairs ) = @_;
    bless {
        titlesRow => defined $pairs{titlesRow} ? $pairs{titlesRow} : 2,
        dataRow => $pairs{dataRow},
        captionDecorations => $pairs{captionDecorations} || [undef],
    }, $class;
}

sub maxLines {
    $_[0]{maxLines};
}

sub positionNextGroup {
    my ( $self, $wb, $ws, $ncol ) = @_;
    my $col       = $self->{$wb}{nextColumn};
    my $titlesRow = $self->{$wb}{titlesRow} ||= $self->{titlesRow};
    my $dataRow   = $self->{$wb}{dataRow};
    unless ( defined $dataRow ) {
        $dataRow = $self->{dataRow};
        my $minDataRow = $titlesRow + 5 + $self->maxLines;
        $dataRow = $minDataRow unless $dataRow && $dataRow > $minDataRow;
        $self->{$wb}{dataRow} = $dataRow;
    }
    my $docRow = $dataRow - 5 - $self->maxLines;
    if ( $self->{$wb}{worksheet} ) {
        $ws = $self->{$wb}{worksheet};
    }
    else {
        $self->{$wb}{worksheet} = $ws;
        $col = 0;
        my $row = $dataRow;
        if ( $wb->{noLinks} && $dataRow - $titlesRow > 3 )
        {    # Hide documentation
            if ( $docRow - $titlesRow > 1 ) {
                $ws->set_row( $_, undef, undef, 1, 2 )
                  foreach $titlesRow + 1 .. $docRow - 1;
            }
            $ws->set_row( $_, undef, undef, 1, 1,
                $_ == $docRow && $docRow - $titlesRow > 1 ? 2 : () )
              foreach $docRow .. $dataRow - 3;
            $ws->set_row( $dataRow - 2, undef, undef, undef, 0, 0, 1 );
        }
        if ( $self->{rows} ) {
            my $thFormat =
              $wb->getFormat( $self->{rows}{defaultFormat} || 'th' );
            my $thgFormat = $wb->getFormat('thg');
            $ws->write( $row++, $col, _shortNameRow( $self->{rows}{list}[$_] ),
                 !$self->{rows}{groups} || defined $self->{rows}{groupid}[$_]
                ? $thFormat
                : $thgFormat )
              foreach 0 .. $#{ $self->{rows}{list} };
        }
        else {
            ++$row;
        }
        $ws->{nextFree} = $row unless $ws->{nextFree} > $row;
        ++$col;
    }
    $self->{$wb}{nextColumn} = $col + $ncol;
    my $deco =
      $self->{captionDecorations}[ $self->{$wb}{nextDecoration} ||= 0 ];
    ++$self->{$wb}{nextDecoration} if $ncol;
    $self->{$wb}{nextDecoration} %= @{ $self->{captionDecorations} };
    $ws, $col, $deco, $titlesRow, $docRow, $dataRow;
}

sub nextColumn {
    my ( $self, $wb ) = @_;
    $self->{$wb}{nextColumn} || 1;
}

sub addDatasetGroup {
    my $self = shift;
    my $group =
      SpreadsheetModel::MatrixSheet::DatasetGroup->new( @_, location => $self );
    return unless @{ $group->{columns} };
    if ( defined $self->{rows} ) {
        die <<ERR
Mismatch in DatasetGroup $self->{name} $self->{debug}
Rows in DatasetGroup: $self->{rows}
Rows in $_->{name} $_->{debug}: $_->{rows}
ERR
          unless $self->{rows} == $group->{rows};
    }
    else {
        $self->{rows} = $group->{rows};
    }
    $self->{maxLines} = $group->{maxLines}
      unless $self->{maxLines} && $self->{maxLines} > $group->{maxLines};
    $self;
}

package SpreadsheetModel::MatrixSheet::DatasetGroup;
use SpreadsheetModel::Object ':_util';
our @ISA = qw(SpreadsheetModel::Object);

sub check {
    my ($self) = @_;
    return "Broken DatasetGroup in $self->{name}"
      unless 'ARRAY' eq ref $self->{columns};
    foreach ( @{ $self->{columns} } ) {
        if ( defined $self->{rows} ) {
            unless ( !$_->{rows} && !$self->{rows}
                || $_->{rows} == $self->{rows} )
            {
                return <<ERR ;
Mismatch in DatasetGroup $self->{name} $self->{debug}
Rows in DatasetGroup: $self->{rows}
Rows in $_->{name} $_->{debug}: $_->{rows}
ERR
            }
        }
        else {
            $self->{rows} = $_->{rows} || 0;
        }
        $_->{location} = $self;
        if ( $_->{arithmetic} && $_->{arguments} ) {
            my @formula = $_->{arithmetic};
            $_->{sourceLines} =
              [ _rewriteFormulas( \@formula, [ $_->{arguments} ] ) ];
            $_->{formulaLines} = [ $_->objectType, @formula ];
        }
        my $lines =
          ( $_->{lines}        ? @{ $_->{lines} }        : 0 ) +
          ( $_->{sourceLines}  ? @{ $_->{sourceLines} }  : 0 ) +
          ( $_->{formulaLines} ? @{ $_->{formulaLines} } : 0 );
        $self->{maxLines} = $lines
          unless $self->{maxLines} && $self->{maxLines} > $lines;
    }
    return;
}

sub wsWrite {
    my ( $self, $wb, $wsheet ) = @_;
    return if $self->{$wb};

    while (1) {
        $self->{rows}->wsPrepare( $wb, $wsheet ) if $self->{rows};
        my $col = $self->{location}->nextColumn($wb);
        foreach ( @{ $self->{columns} } ) {
            $_->{cols}->wsPrepare( $wb, $wsheet ) if $_->{cols};
            die "$_->{name} $_->{debug}"
              . ' is already in the workbook and'
              . ' cannot be written again as part of '
              . "$self->{name} $self->{debug}"
              if $_->{$wb};
            $_->wsPrepare( $wb, $wsheet );
            $_->{$wb} ||= {};    # Placeholder
        }
        last if $col == $self->{location}->nextColumn($wb);
        delete $_->{$wb} foreach @{ $self->{columns} };
    }
    my $ncol = $self->nCol;
    my ( $ws, $col, $deco, $titlesRow, $docRow, $dataRow ) =
      $self->{location}->positionNextGroup( $wb, $wsheet, $ncol );

    if ( $wb->{logger} ) {
        foreach ( @{ $self->{columns} } ) {
            $_->addTableNumber( $wb, $ws );
            $wb->{logger}->log($_);
        }
    }

    if ( $ncol == 1 ) {
        $ws->write_string( $titlesRow, $col, "$self->{name}",
            $wb->getFormat( 'captionca', $deco || (), 'tlttr' ) );
    }
    elsif ( $wb->{mergedRanges} ) {    # merged cell range
        $ws->merge_range( $titlesRow, $col, $titlesRow, $col + $ncol - 1,
            "$self->{name}",
            $wb->getFormat( 'captionca', $deco || (), 'tlttr' ) );
    }
    else {    # center-across formatting; buggy in Excel 2013?
        my $captionFormat = $wb->getFormat( 'captionca', $deco || () );
        $ws->write( $titlesRow, $col,      "$self->{name}", $captionFormat );
        $ws->write( $titlesRow, $col + $_, undef,           $captionFormat )
          foreach 1 .. $ncol - 2;
        $ws->write( $titlesRow, $col + $ncol - 1,
            undef, $wb->getFormat( 'captionca', $deco || (), 'tlttr' ) );
    }

    my $c4 = $col;
    foreach my $column ( @{ $self->{columns} } ) {

        @{ $column->{$wb} }{qw(worksheet row col)} = ( $ws, $dataRow, $c4 );

        if ( $column->{lines}
            or !( $wb->{noLinks} && $wb->{noLinks} == 1 )
            and $column->{formulaLines}
            || $column->{name} && $column->{sourceLines} )
        {
            my $lcol = $column->{cols} ? $#{ $column->{cols}{list} } : 0;
            my @decorations = $c4 + $lcol == $col + $ncol - 1 ? 'tlttr' : ();
            my $textFormat = $wb->getFormat( 'text', 'wrapca', @decorations );
            my $linkFormat = $wb->getFormat( 'link', 'wrapca', @decorations );
            my $xc         = 0;
            my @allLines   = (
                $column->{lines} ? @{ $column->{lines} } : (),
                !( $wb->{noLinks} && $wb->{noLinks} == 1 )
                  && $column->{sourceLines} && @{ $column->{sourceLines} }
                ? ( 'Data sources:', @{ $column->{sourceLines} } )
                : (),
                !( $wb->{noLinks} && $wb->{noLinks} == 1 )
                  && $column->{formulaLines} ? @{ $column->{formulaLines} }
                : ()
            );
            my $row = $docRow;
            foreach (@allLines) {

                if ( UNIVERSAL::isa( $_, 'SpreadsheetModel::Object' ) ) {
                    my $na = 'x' . ( ++$xc ) . " = $_->{name}";
                    if ( my $url = $_->wsUrl($wb) ) {
                        $ws->write_url( ++$row, $c4, $url, $na, $linkFormat );
                        $ws->write( $row, $c4 + $_, undef, $linkFormat )
                          foreach 1 .. $lcol;
                        (
                            $_->{location}
                              && ref $_->{location} eq
                              'SpreadsheetModel::Columnset'
                            ? $_->{location}
                            : $_
                          )->addForwardLink($column)
                          if $wb->{findForwardLinks};
                    }
                    else {
                        $ws->write_string( ++$row, $c4, $na, $textFormat );
                        $ws->write( $row, $c4 + $_, undef, $textFormat )
                          foreach 1 .. $lcol;
                    }
                }
                else {
                    $ws->write_string( ++$row, $c4, "$_", $textFormat );
                    $ws->write( $row, $c4 + $_, undef, $textFormat )
                      foreach 1 .. $lcol;
                }
            }
            while ( ++$row < $dataRow - 3 ) {
                $ws->write_string( $row, $c4, ' ', $textFormat );
                $ws->write( $row, $c4 + $_, undef, $textFormat )
                  foreach 1 .. $lcol;
            }
        }

        my $cell = $column->wsPrepare( $wb, $ws );
        foreach my $x ( $column->colIndices ) {
            foreach my $y ( $column->rowIndices ) {
                my ( $value, $format, $formula, @more ) = $cell->( $x, $y );
                $format = $wb->getFormat( $format, 'tlttr' )
                  if $c4 + $x == $col + $ncol - 1;
                if (@more) {
                    $ws->repeat_formula( $dataRow + $y,
                        $c4 + $x, $formula, $format, @more );
                }
                elsif ($formula) {
                    $ws->write_formula( $dataRow + $y,
                        $c4 + $x, $formula, $format );
                }
                else {
                    $ws->write( $dataRow + $y, $c4 + $x, $value, $format );
                }
            }
        }

        my $co = $column->{cols};
        if ( $co and $#{ $co->{list} } || -1 == index lc $column->{name},
            lc _shortNameCol( $co->{list}[0] ) )
        {
            if ( $#{ $co->{list} } && $wb->{mergedRanges} ) {
                my @decorations =
                  $c4 + $#{ $co->{list} } == $col + $ncol - 1 ? 'tlttr' : ();
                $ws->merge_range(
                    $dataRow - 2,
                    $c4,
                    $dataRow - 2,
                    $c4 + $#{ $co->{list} },
                    "$column->{name}",
                    $wb->getFormat( 'thca', @decorations )
                );
            }
            foreach ( 0 .. $#{ $co->{list} } ) {
                my @decorations = $c4 + $_ == $col + $ncol - 1 ? 'tlttr' : ();
                $ws->write(
                    $dataRow - 2,
                    $c4 + $_,
                    $_ ? undef : "$column->{name}",
                    $wb->getFormat( 'thca', @decorations )
                ) unless $#{ $co->{list} } && $wb->{mergedRanges};
                $ws->write(
                    $dataRow - 1,
                    $c4 + $_,
                    _shortNameCol( $co->{list}[$_] ),
                    $wb->getFormat(
                        !$co->{groups}
                          || defined $co->{groupid}[$_] ? 'thc' : 'thg',
                        @decorations
                    )
                );
            }
            $c4 += @{ $co->{list} };
        }
        else {
            my @decorations = $c4 == $col + $ncol - 1 ? 'tlttr' : ();
            $ws->write(
                $dataRow - 2,
                $c4,
                $column->{number} || $column->{numbered},
                $wb->getFormat( 'thca', @decorations )
            );
            $ws->write(
                $dataRow - 1,
                $c4,
                _shortName( $column->{name} ),
                $wb->getFormat( 'thc', @decorations )
            );
            ++$c4;
        }

        $_->( $column, $wb, $ws, undef, undef )
          foreach map { @{ $column->{postWriteCalls}{$_} }; }
          grep { $column->{postWriteCalls}{$_} } 'obj', $wb;

    }

}

sub nCol {
    my ($self) = @_;
    my $c = 0;
    $c +=
      $_->{cols}
      ? @{ $_->{cols}{list} }
      : 1
      foreach @{ $self->{columns} };
    $c;
}

1;
