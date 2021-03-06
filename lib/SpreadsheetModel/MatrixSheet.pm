﻿package SpreadsheetModel::MatrixSheet;

# Copyright 2015-2020 Franck Latrémolière, Reckon LLP and others.
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

use Spreadsheet::WriteExcel::Utility;
use SpreadsheetModel::Object ':_util';

sub new {
    my ( $class, %pairs ) = @_;
    my $matrixSheet = bless {
        captionDecorations => $pairs{captionDecorations} || [undef],
        dataRow            => $pairs{dataRow},
        noLines            => $pairs{noLines},
        noDoubleNames      => $pairs{noDoubleNames},
        noNumbers          => $pairs{noNumbers},
        titlesRow => defined $pairs{titlesRow} ? $pairs{titlesRow} : 2,
        verticalSpace => 3 +
          ( $pairs{noDoubleNames} ? 0 : 1 ) +
          ( $pairs{noNumbers}     ? 0 : 1 ),
    }, $class;
    $matrixSheet;
}

sub maxLines {
    my ($matrixSheet) = @_;
    $matrixSheet->{noLines} ? 0 : $matrixSheet->{maxLines};
}

sub positionNextGroup {
    my ( $matrixSheet, $wb, $ws, $ncol ) = @_;
    my $col       = $matrixSheet->{$wb}{nextColumn};
    my $titlesRow = $matrixSheet->{$wb}{titlesRow} ||=
      $matrixSheet->{titlesRow};
    my $dataRow = $matrixSheet->{$wb}{dataRow};
    unless ( defined $dataRow ) {
        $dataRow = $matrixSheet->{dataRow};
        my $minDataRow =
          $titlesRow + $matrixSheet->{verticalSpace} + $matrixSheet->maxLines;
        $dataRow = $minDataRow unless $dataRow && $dataRow > $minDataRow;
        $matrixSheet->{$wb}{dataRow} = $dataRow;
    }
    my $docRow =
      $dataRow - $matrixSheet->{verticalSpace} - $matrixSheet->maxLines;
    if ( $matrixSheet->{$wb}{worksheet} ) {
        $ws = $matrixSheet->{$wb}{worksheet};
    }
    else {
        $matrixSheet->{$wb}{worksheet} = $ws;
        $col = 0;
        my $row = $dataRow;
        if (   $wb->{noLinks}
            && $wb->{noLinks} == 2
            && $dataRow - $titlesRow > 3 )
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
        if ( $matrixSheet->{rows} ) {
            my $thFormat =
              $wb->getFormat( $matrixSheet->{rows}{defaultFormat} || 'th' );
            my $thgFormat = $wb->getFormat('thg');
            $ws->write(
                $row++,
                $col,
                _shortNameRow( $matrixSheet->{rows}{list}[$_] ),
                !$matrixSheet->{rows}{groups}
                  || defined $matrixSheet->{rows}{groupid}[$_]
                ? $thFormat
                : $thgFormat
            ) foreach 0 .. $#{ $matrixSheet->{rows}{list} };
        }
        else {
            ++$row;
        }
        $ws->{nextFree} = $row unless $ws->{nextFree} > $row;
        ++$col;
    }
    $matrixSheet->{$wb}{nextColumn} = $col + $ncol;
    my $deco =
      $matrixSheet->{captionDecorations}
      [ $matrixSheet->{$wb}{nextDecoration} ||= 0 ];
    ++$matrixSheet->{$wb}{nextDecoration} if $ncol;
    $matrixSheet->{$wb}{nextDecoration} %=
      @{ $matrixSheet->{captionDecorations} };
    $ws, $col, $deco, $titlesRow, $docRow, $dataRow;
}

sub nextColumn {
    my ( $matrixSheet, $wb ) = @_;
    $matrixSheet->{$wb}{nextColumn} || 1;
}

sub addDatasetGroup {
    my $matrixSheet = shift;
    my $group =
      SpreadsheetModel::MatrixSheet::DatasetGroup->new( @_,
        location => $matrixSheet );
    return unless @{ $group->{columns} };
    if ( defined $matrixSheet->{rows} ) {
        die <<ERR
Mismatch in MatrixSheet $matrixSheet->{name} $matrixSheet->{debug}
Rows in MatrixSheet: $matrixSheet->{rows}
Rows in $_->{name} $_->{debug}: $_->{rows}
ERR
          unless $matrixSheet->{rows} == $group->{rows};
    }
    else {
        $matrixSheet->{rows} = $group->{rows};
    }
    $matrixSheet->{maxLines} = $group->{maxLines}
      unless $matrixSheet->{maxLines}
      && $matrixSheet->{maxLines} > $group->{maxLines};
    $matrixSheet;
}

package SpreadsheetModel::MatrixSheet::DatasetGroup;
use SpreadsheetModel::Object ':_util';
our @ISA = qw(SpreadsheetModel::Object);

sub check {
    my ($dsGroup) = @_;
    return "Broken DatasetGroup in $dsGroup->{name}"
      unless 'ARRAY' eq ref $dsGroup->{columns};
    foreach ( @{ $dsGroup->{columns} } ) {
        if ( defined $dsGroup->{rows} ) {
            unless ( !$_->{rows} && !$dsGroup->{rows}
                || $_->{rows} == $dsGroup->{rows} )
            {
                return <<ERR ;
Mismatch in DatasetGroup $dsGroup->{name} $dsGroup->{debug}
Rows in DatasetGroup: $dsGroup->{rows}
Rows in $_->{name} $_->{debug}: $_->{rows}
ERR
            }
        }
        else {
            $dsGroup->{rows} = $_->{rows} || 0;
        }
        $_->{location} = $dsGroup;
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
        $dsGroup->{maxLines} = $lines
          unless $dsGroup->{maxLines} && $dsGroup->{maxLines} > $lines;
    }
    return;
}

sub wsWrite {
    my ( $dsGroup, $wb, $wsheet ) = @_;
    return if $dsGroup->{$wb};

    while (1) {
        $dsGroup->{rows}->wsPrepare( $wb, $wsheet ) if $dsGroup->{rows};
        my $col = $dsGroup->{location}->nextColumn($wb);
        foreach ( @{ $dsGroup->{columns} } ) {
            $_->{cols}->wsPrepare( $wb, $wsheet ) if $_->{cols};
            die "$_->{name} $_->{debug}"
              . ' is already in the workbook and'
              . ' cannot be written again as part of '
              . "$dsGroup->{name} $dsGroup->{debug}"
              if $_->{$wb};
            $_->wsPrepare( $wb, $wsheet );
            $_->{$wb} ||= {};    # Placeholder
        }
        last if $col == $dsGroup->{location}->nextColumn($wb);
        delete $_->{$wb} foreach @{ $dsGroup->{columns} };
    }

    my $ncol = 0;
    $ncol +=
      $_->{cols}
      ? @{ $_->{cols}{list} }
      : 1
      foreach @{ $dsGroup->{columns} };

    my ( $ws, $col, $deco, $titlesRow, $docRow, $dataRow ) =
      $dsGroup->{location}->positionNextGroup( $wb, $wsheet, $ncol );

    if ( $wb->{logger} ) {
        foreach ( @{ $dsGroup->{columns} } ) {
            $_->addTableNumber( $wb, $ws );
            $wb->{logger}->log($_);
        }
    }
    my $showNumbers;
    unless ( $dsGroup->{location}->{noNumbers} ) {
        foreach ( @{ $dsGroup->{columns} } ) {
            if ( $_->{number}
                || UNIVERSAL::isa( $_->{location}, 'SpreadsheetModel::Object' )
                && $_->{location}{number} )
            {
                $showNumbers = 1;
                last;
            }
        }
    }

    if ( $ncol == 1 ) {
        $ws->write_string( $titlesRow, $col, "$dsGroup->{name}",
            $wb->getFormat( 'captionca', $deco || (), 'tlttr' ) );
    }
    elsif ( $wb->{mergedRanges} ) {    # merged cell range
        $ws->merge_range( $titlesRow, $col, $titlesRow, $col + $ncol - 1,
            "$dsGroup->{name}",
            $wb->getFormat( 'captionca', $deco || (), 'tlttr' ) );
    }
    else {    # center-across formatting; might be buggy in Microsoft Excel 2013
        my $captionFormat = $wb->getFormat( 'captionca', $deco || () );
        $ws->write( $titlesRow, $col, "$dsGroup->{name}", $captionFormat );
        $ws->write( $titlesRow, $col + $_, undef, $captionFormat )
          foreach 1 .. $ncol - 2;
        $ws->write( $titlesRow, $col + $ncol - 1,
            undef, $wb->getFormat( 'captionca', $deco || (), 'tlttr' ) );
    }

    my $c4 = $col;
    foreach my $column ( @{ $dsGroup->{columns} } ) {

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
            foreach ( $dsGroup->{location}{noLines} ? () : @allLines ) {

                if ( UNIVERSAL::isa( $_, 'SpreadsheetModel::Object' ) ) {
                    my $na = 'x' . ( ++$xc ) . " = $_->{name}";
                    if ( my $url = $_->wsUrl($wb) ) {
                        $ws->write_url( ++$row, $c4, $url, $na, $linkFormat );
                        $ws->write( $row, $c4 + $_, undef, $linkFormat )
                          foreach 1 .. $lcol;
                        (
                            $_->{location}
                              && UNIVERSAL::can( $_->{location}, 'wsWrite' )
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
            if (   $#{ $co->{list} }
                && $wb->{mergedRanges} )
            {
                my @decorations =
                  $c4 + $#{ $co->{list} } == $col + $ncol - 1 ? 'tlttr' : ();
                $ws->merge_range(
                    $dataRow - 2,
                    $c4,
                    $dataRow - 2,
                    $c4 + $#{ $co->{list} },
                    "$column->{name}",
                    $wb->getFormat( 'thca', @decorations )
                ) unless $dsGroup->{location}{noDoubleNames};
            }
            foreach ( 0 .. $#{ $co->{list} } ) {
                my @decorations = $c4 + $_ == $col + $ncol - 1 ? 'tlttr' : ();
                $ws->write(
                    $dataRow - 2,
                    $c4 + $_,
                    $_ ? undef : "$column->{name}",
                    $wb->getFormat( 'thca', @decorations )
                  )
                  unless $dsGroup->{location}{noDoubleNames}
                  || $#{ $co->{list} } && $wb->{mergedRanges};
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
            if ($showNumbers) {
                if (
                    my $number =
                    $column->{number} ? $column->{number}
                    : UNIVERSAL::isa( $column->{location},
                        'SpreadsheetModel::Object' )
                    ? $column->{location}{number}
                    : undef
                  )
                {
                    $ws->write( $dataRow - 2,
                        $c4, $number, $wb->getFormat( 'thca', @decorations ) );
                }
            }
            $ws->write(
                $dataRow - 1,
                $c4,
                _shortName( $column->{name} ),
                $wb->getFormat( 'thc', @decorations )
            ) unless $dsGroup->{location}{noNames};
            ++$c4;
        }

        $_->( $column, $wb, $ws, undef, undef )
          foreach map { @{ $column->{postWriteCalls}{$_} }; }
          grep { $column->{postWriteCalls}{$_} } 'obj', $wb;

    }

}

1;
