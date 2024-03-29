﻿package SpreadsheetModel::Columnset;

# Copyright 2008-2021 Franck Latremoliere and others.
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

use SpreadsheetModel::Object ':_util';
our @ISA = qw(SpreadsheetModel::Objectset);

use SpreadsheetModel::Label;
use SpreadsheetModel::Stack;

sub wsUrl {
    my $self = shift;
    $self->{columns}[0]->wsUrl(@_);
}

sub populateCore {
    my ($self) = @_;
    $self->{core}{columns} = [ map { $_->getCore } @{ $self->{columns} } ];
}

sub check {
    my ($self) = @_;
    return "Broken column list in Columnset $self->{name}"
      unless 'ARRAY' eq ref $self->{columns};
    my @columns = @{ $self->{columns} };
    return "No columns in Columnset $self->{name}" unless @columns;
    $self->{lines} = [ SpreadsheetModel::Object::splitLines( $self->{lines} ) ]
      if $self->{lines};
    my $rows;
    my $colOffset = 0;

    foreach (@columns) {
        if ( defined $rows ) {
            unless ( !$_->{rows} && !$rows
                || $_->{rows} == $rows
                || $_->{rows} && $_->{rows}{accepts} && grep { $rows == $_ }
                @{ $_->{rows}{accepts} } )
            {
                return <<ERR ;
Mismatch in Columnset $self->{name} $self->{debug}
Rows in Columnset: $rows
Rows in $_->{name} $_->{debug}: $_->{rows}
ERR
            }
        }
        else {
            $rows = $_->{rows} || 0;
        }
        $_->{location} = $self;
        $_->{name} =
          new SpreadsheetModel::Label( $_->{name},
            "$_->{name} (in $self->{name})" )
          if $self->{name}
          && $self->{name} !~ /#/
          && index( $_->{name}, " (in $self->{name})" ) < 0;
        $_->{dataset} ||= $self->{dataset}
          if ref $_ eq 'SpreadsheetModel::Dataset';
        $_->{number} ||= $self->{number};
        $_->{colOffset} = $colOffset;
        $colOffset += 1 + $_->lastCol;
    }
    $self->{rows} = $rows;
    return;
}

sub objectType {
    my @col = @{ $_[0]{columns} };
    my $ty  = ( shift @col )->objectType;
    foreach (@col) {
        return 'Composite' if $ty ne $_->objectType;
    }
    $ty;
}

sub wsAdopt {
    my ( $self, $wbook, $wsheet, $source ) = @_;
    return unless UNIVERSAL::isa( $source, __PACKAGE__ );
    my $lastColumn = $#{ $self->{columns} };
    return unless $lastColumn == $#{ $source->{columns} };
    return
      if grep {
        !$self->{columns}[$_]
          ->wsAdopt( $wbook, $wsheet, $source->{columns}[$_] );
      } 0 .. $lastColumn;
    $self->{$wbook}{$wsheet} = 1;
}

sub wsWrite {

    my ( $self, $wb, $ws, $row, $col, $noCopy ) = @_;

    if ( $self->{$wb} ) {
        return values %{ $self->{$wb} }
          if !$wb->{copy} || $noCopy || $self->{$wb}{$ws};
        return $self->{$wb}{$ws} = new SpreadsheetModel::Columnset(
            name    => Label( $self->{name}, "$self->{name} (copy)" ),
            columns => [
                map {
                    $_->{$wb}{$ws} =
                      new SpreadsheetModel::Stack( sources => [$_] );
                } @{ $self->{columns} }
            ]
        )->wsWrite( $wb, $ws, $row, $col );
    }

    {
        my $wsWanted;
        $wsWanted = $wb->{ $self->{location} } if $self->{location};
        $wsWanted = $wb->{dataSheet}
          if !$wsWanted
          && !$self->{ignoreDatasheet}
          && !grep { ref $_ ne 'SpreadsheetModel::Dataset' }
          @{ $self->{columns} };
        return $self->wsWrite( $wb, $wsWanted, undef, undef, 1 )
          if $wsWanted && $wsWanted != $ws;
    }

    if (
           $wb->{dataSheet}
        && !$self->{ignoreDatasheet}
        && (
            my @dataColumns =
            grep { ref $_ eq 'SpreadsheetModel::Dataset' } @{ $self->{columns} }
        )
      )
    {
        if ( $ws != $wb->{dataSheet} && !$self->{doNotCopyInputColumns} ) {
            my $data = bless {%$self}, __PACKAGE__;
            $data->{columns} = \@dataColumns;
            $_->{location} = $data foreach @dataColumns;
            delete $self->{number};
            delete $self->{lines};
            my %mapping = map {
                $_ => new SpreadsheetModel::Stack(
                    sources  => [$_],
                    location => $self
                  )
            } @dataColumns;
            $self->{columns} = [
                map {
                    if ( $mapping{$_} ) {
                        $mapping{$_};
                    }
                    else {
                        if ( $_->{arguments} ) {
                            foreach my $k ( keys %{ $_->{arguments} } ) {
                                my $x = $mapping{ $_->{arguments}{$k} };
                                $_->{arguments}{$k} = $x if $x;
                            }
                        }
                        $_;
                    }
                } @{ $self->{columns} }
            ];    # This assumes Arithmetic objects
            $data->wsWrite( $wb, $wb->{dataSheet} );
        }
    }

    my @cell;

    my $lastCol = $#{ $self->{columns} };

    my $lastRow = $self->{rows} ? $#{ $self->{rows}{list} } : 0;

    my $headerCols =
         $self->{rows}
      || !exists $self->{singleRowName}
      || defined $self->{singleRowName} ? 1 : 0;

    my @sourceLines;
    @sourceLines = @{ $self->{sourceLines} } if $self->{sourceLines};

    {
        my @formulas =
          map {
               !$_->{deferWritingTo} && $_->{arguments}
              ? $_->{arithmetic}
              : ''
          } @{ $self->{columns} };
        if ( grep { $_ } @formulas ) {
            @sourceLines = (
                'Data sources:',
                _rewriteFormulas(
                    \@formulas,
                    [ map { $_->{arguments} } @{ $self->{columns} } ]
                )
            ) unless $wb->{noLinks} && $wb->{noLinks} == 1;
            unless ( $self->{formulasDone} ) {
                $self->{formulasDone} = 1;
                unless ( $wb->{noLinks} && $wb->{noLinks} == 1 ) {
                    push @{ $self->{lines} },
                      [
                        $headerCols ? ('Kind:') : (),
                        map { $_->objectType } @{ $self->{columns} }
                      ];
                    push @{ $self->{lines} },
                      [ $headerCols ? ('Formula:') : (), @formulas ];
                }
            }
        }
    }

    my $dualHeaded = 0;
    foreach ( @{ $self->{columns} } ) {
        if (   $_->{cols}
            && $#{ $_->{cols}{list} } )
        {
            $dualHeaded = 1;
            last;
        }
    }

    # Blank line
    my $headerLines = $self->{name} || $self->{lines} || @sourceLines ? 1 : 0;

    $headerLines += $dualHeaded ? 2 : 1 unless $self->{noHeaders};

    ++$headerLines if $self->{name};
    $headerLines += @{ $self->{lines} } if $self->{lines};
    $headerLines += @sourceLines;

    $self->{$wb}{$ws} = $ws;

    while (1) {
        unless ( defined $row && defined $col ) {
            ( $row, $col ) = ( ( $ws->{nextFree} ||= -1 ) + 1, 0 );
        }
        for ( 0 .. $lastCol ) {
            my $thecol = $self->{columns}[$_];
            foreach (qw(rows cols)) {
                $thecol->{$_}->wsPrepare( $wb, $ws ) if $thecol->{$_};
            }
            $thecol->{name} .= " ($thecol->{debug})"
              if $wb->{debug} && 0 > index $thecol->{name}, $thecol->{debug};
            die "$thecol->{name} $thecol->{debug}"
              . ' is already in the workbook and'
              . ' cannot be written again as part of '
              . "$self->{name} $self->{debug}"
              if $thecol->{$wb};
            $thecol->wsPrepare( $wb, $ws ) unless $thecol->{deferWritingTo};
            $thecol->{$wb} ||= {};    # Placeholder for other columns
        }
        last if !$ws->{nextFree} || $ws->{nextFree} < $row;
        delete $_->{$wb} for @{ $self->{columns} };
        undef $row;
    }

    --$row if $self->{noSpacing};

    {
        my $c2 = $col;
        for ( 0 .. $lastCol ) {
            my $thecol = $self->{columns}[$_];
            $cell[$_] = $thecol->wsPrepare( $wb, $ws )
              unless $thecol->{deferWritingTo};
            @{ $thecol->{$wb} }{qw(worksheet row col)} =
              ( $ws, $row + $headerLines, $c2 + $headerCols );
            $c2 +=
              $thecol->{cols}
              ? @{ $thecol->{cols}{list} }
              : 1;
        }
    }

    my $number = $wb->{logger}
      && $self->{name} ? $self->addTableNumber( $wb, $ws ) : undef;

    $self->{name} .= " ($self->{debug})"
      if $wb->{debug} && 0 > index $self->{name}, $self->{debug};

    return if $self->{rows} && !@{ $self->{rows}{list} };

    if ( $self->{name} ) {
        $ws->set_row( $row, $wb->{captionRowHeight} );
        $ws->write_string( $row++, $col, "$self->{name}",
            $wb->getFormat('caption') );
    }

    if ( $self->{lines} || @sourceLines ) {
        my $hideFormulas = $wb->{noLinks} && @sourceLines;
        my $textFormat   = $wb->getFormat('text');
        my $linkFormat   = $wb->getFormat('link');
        my $xc           = 0;
        my @arrayLines;
        foreach ( $self->{lines} ? @{ $self->{lines} } : (), @sourceLines, ) {
            if ( !defined $_ ) {
                $ws->set_row( $row, undef, undef, 1, 1 )
                  if $hideFormulas;
                ++$row;
            }
            elsif ( ref $_ eq 'ARRAY' ) {
                push @arrayLines, $_;
            }
            elsif ( UNIVERSAL::isa( $_, 'SpreadsheetModel::Object' ) ) {
                my $na = 'x' . ( ++$xc ) . " = $_->{name}";
                if ( my $url = $_->wsUrl($wb) ) {
                    $ws->set_row( $row, undef, undef, 1, 1 )
                      if $hideFormulas;
                    $ws->write_url( $row++, $col, $url, $na, $linkFormat );
                    (
                        $_->{location}
                          && UNIVERSAL::can( $_->{location}, 'wsWrite' )
                        ? $_->{location}
                        : $_
                      )->addForwardLink($self)
                      if $wb->{findForwardLinks};
                }
                else {
                    $ws->set_row( $row, undef, undef, 1, 1 )
                      if $hideFormulas;
                    $ws->write_string( $row++, $col, $na, $textFormat );
                }
            }
            elsif (/^(https?|mailto:)/) {
                $ws->set_row( $row, undef, undef, 1, 1 )
                  if $hideFormulas;
                $ws->write_url( $row++, $col, "$_", "$_", $linkFormat );
            }
            else {
                $ws->set_row( $row, undef, undef, 1, 1 )
                  if $hideFormulas;
                $ws->write_string( $row++, $col, "$_", $textFormat );
            }
        }
        if (@arrayLines) {
            foreach (@arrayLines) {
                my $c3 = $col;
                $ws->write_string( $row, $c3++, "$_->[0]",
                    $wb->getFormat('colnotecenter') )
                  if $headerCols;

                foreach my $cn ( 0 .. $lastCol ) {
                    my $wrapFormat = $wb->getFormat(
                        $self->{columns}[$cn]{cols}
                          && $#{ $self->{columns}[$cn]{cols}{list} } > 2
                        ? 'colnoteleft'
                        : 'colnotecenter'
                    );
                    $ws->write_string( $row, $c3++,
                        "$_->[ $cn + $headerCols ]", $wrapFormat );
                    $ws->write( $row, $c3++, undef, $wrapFormat )
                      foreach 1 .. $self->{columns}[$cn]{cols}
                      && $#{ $self->{columns}[$cn]{cols}{list} };
                }
                $ws->set_row( $row, undef, undef, 1, 1 )
                  if $hideFormulas;
                ++$row;
            }
        }
        if ($hideFormulas) {
            $ws->set_row( $row, undef, undef, 1, 1 );
            $ws->set_row( $row + 1, undef, undef, undef, 0, 0, 1 );
        }

    }

    # Blank line
    ++$row if $self->{name} || $self->{lines} || @sourceLines;

    my ( @dataDecorationsByColumn, @headerDecorationsByColumn );
    if ( my $columnDecorationsClosure = $wb->{columnDecorations} ) {
        my $c = 0;
        foreach ( 0 .. $lastCol ) {
            if ( $self->{columns}[$_]->lastCol ) {
                if (
                    my @decos = $columnDecorationsClosure->(
                        @{ $self->{columns}[$_]{cols}{list} }
                    )
                  )
                {
                    foreach ( 0 .. $#decos ) {
                        if ( my $deco = $decos[$_] ) {
                            $headerDecorationsByColumn[ $c + $_ ] = $deco->[0];
                            $dataDecorationsByColumn[ $c + $_ ]   = $deco->[1];
                        }
                    }
                }
                $c += @{ $self->{columns}[$_]{cols}{list} };
            }
            else {
                if (
                    my ($deco) = $columnDecorationsClosure->(
                        $self->{columns}[$_]->objectShortName
                    )
                  )
                {
                    $headerDecorationsByColumn[$c] = $deco->[0];
                    $dataDecorationsByColumn[$c]   = $deco->[1];
                }
                ++$c;
            }
        }
    }

    unless ( $self->{noHeaders} ) {

        my $c4 = $col + $headerCols;
        $row++ if $dualHeaded;
        foreach ( 0 .. $lastCol ) {
            my $deferred = $self->{columns}[$_]{deferWritingTo};
            if ( !$deferred && $wb->{logger} ) {
                if ($number) {
                    my $n = $self->{columns}[$_]{name};
                    $self->{columns}[$_]{name} =
                      new SpreadsheetModel::Label( $n, $number . $n );
                }
                elsif ( $self->{name} ) {
                    $self->{columns}[$_]->addTableNumber( $wb, $ws, 1 );
                }
                $wb->{logger}->log( $self->{columns}[$_] );
            }
            my $colShortName = _shortNameCol $self->{columns}[$_]{name};
            my $co           = $self->{columns}[$_]{cols};
            if ( $co && $#{ $co->{list} } ) {
                if ( $wb->{mergedRanges} ) {
                    $ws->merge_range( $row - 1, $c4, $row - 1,
                        $c4 + $#{ $co->{list} },
                        $colShortName, $wb->getFormat('thca') );
                }
                else {
                    my $caFormat =
                      $wb->getFormat(
                        $#{ $co->{list} } > 2 ? 'thcaleft' : 'thca' );
                    $ws->write( $row - 1, $c4 + $_, $_ ? undef : $colShortName,
                        $caFormat )
                      foreach 0 .. $#{ $co->{list} };
                }
                $ws->write(
                    $row,
                    $c4 + $_,
                    _shortNameCol( $co->{list}[$_] ),
                    $wb->getFormat(
                        !$co->{groups} || defined $co->{groupid}[$_]
                        ? (
                            'thc',
                            @headerDecorationsByColumn
                              && $headerDecorationsByColumn[ $c4 + $_ - $col -
                              $headerCols ]
                            ? $headerDecorationsByColumn[ $c4 + $_ - $col -
                              $headerCols ]
                            : ()
                          )
                        : 'thg'
                    )
                ) foreach 0 .. $#{ $co->{list} };
                $c4 += @{ $co->{list} };
            }
            elsif ($deferred) {
                my $row_clos  = $row;
                my $col_clos  = $c4;
                my $name_clos = $colShortName;
                my $target =
                  UNIVERSAL::can( $deferred->{location}, 'wsWrite' )
                  ? $deferred->{location}
                  : $deferred;
                push @{ $target->{postWriteCalls}{$wb} }, sub {
                    my ( $obj_pwc, $wb_pwc, $ws_pwc, $rowref_pwc, $col_pwc ) =
                      @_;
                    $ws->write(
                        $row_clos,
                        $col_clos,
                        $deferred->wsUrl($wb_pwc),
                        $name_clos,
                        $wb_pwc->getFormat(
                            [ base => 'thc', underline => 1, ]
                        )
                    );
                };
            }
            else {
                $ws->write(
                    $row, $c4,
                    $colShortName,
                    $wb->getFormat(
                        'thc',
                        @headerDecorationsByColumn
                          && $headerDecorationsByColumn[ $c4 - $col -
                          $headerCols ]
                        ? $headerDecorationsByColumn[ $c4 - $col - $headerCols ]
                        : ()
                    )
                );
                ++$c4;
            }

        }

        ++$row;

    }

    $col += $headerCols;

    if ( $self->{rows} ) {
        my $thformat = $wb->getFormat( $self->{rows}{defaultFormat} || 'th' );
        my $thgformat = $wb->getFormat('thg');
        for ( my $r = 0 ; $r <= $lastRow ; ++$r ) {
            if ( !$self->{rows}{groups}
                || defined $self->{rows}{groupid}[$r] )
            {
                $ws->write( $row + $r,
                    $col - 1, _shortNameRow( $self->{rows}{list}[$r] ),
                    $thformat );
            }
            else {
                $ws->write( $row + $r,
                    $col - 1, _shortNameRow( $self->{rows}{list}[$r] ),
                    $thgformat );
            }
        }
    }

    elsif ( !exists $self->{singleRowName} ) {
        $ws->write(
            $row, $col - 1,
            _shortNameRow( $self->{name} ),
            $wb->getFormat('th')
        );
    }

    elsif ( $self->{singleRowName} ) {
        $ws->write(
            $row, $col - 1,
            _shortNameRow( $self->{singleRowName} ),
            $wb->getFormat('th')
        );
    }

    my $c2 = $col;

    my $comment = $self->{comment};
    $comment = $self->{lines}
      if !defined $comment
      && $wb->{linesAsComment}
      && $self->{lines};

    foreach my $c ( 0 .. $lastCol ) {

        $comment = $self->{columns}[$c]{comment}
          if $self->{columns}[$c]{comment};
        $comment = $self->{columns}[$c]{lines}
          if !defined $comment
          && $wb->{linesAsComment}
          && $self->{columns}[$c]{lines};

        if ( my $co = $self->{columns}[$c]{cols} ) {

            if ( my $deferred = $self->{columns}[$c]{deferWritingTo} ) {
                _deferWritingData( $self->{columns}[$c],
                    $deferred, $wb, $ws, $row, $c2 );
            }
            else {
                foreach my $y ( $self->{columns}[$c]->rowIndices ) {
                    foreach my $x ( $self->{columns}[$c]->colIndices ) {
                        my ( $value, $format, $formula, @more ) =
                          $cell[$c]->( $x, $y );
                        if (@dataDecorationsByColumn) {
                            if ( my $dd =
                                $dataDecorationsByColumn[ $c2 + $x - $col ] )
                            {
                                $format = $wb->getFormat( $format, $dd );
                            }
                        }
                        if (@more) {
                            $ws->repeat_formula(
                                $row + $y, $c2 + $x, $formula,
                                $format,   @more
                            );
                        }
                        elsif ($formula) {
                            $ws->write_formula(
                                $row + $y, $c2 + $x, $formula,
                                $format,   $value
                            );
                        }
                        else {
                            $value = "=$value"
                              if $value
                              and $value eq '#VALUE!' || $value eq '#N/A'
                              and $wb->formulaHashValues;
                            $ws->write( $row + $y, $c2 + $x, $value, $format );
                            if ($comment) {
                                $ws->write_comment(
                                    $row + $y,
                                    $c2 + $x,
                                    (
                                        map {
                                            ref $_ eq 'ARRAY'
                                              ? join "\n", @$_
                                              : $_;
                                          } ref $comment eq 'HASH'
                                        ? $comment->{text}
                                        : $comment
                                    ),
                                    x_scale => ref $comment eq 'HASH'
                                      && $comment->{x_scale}
                                    ? $comment->{x_scale}
                                    : 3,
                                );
                                undef $comment;
                            }
                        }
                    }
                }
            }

            $self->{columns}[$c]
              ->dataValidation( $wb, $ws, $row, $c2, $row + $lastRow )
              if $self->{columns}[$c]{validation};

            $self->{columns}[$c]->conditionalFormatting(
                $wb, $ws, $row, $c2,
                $row + $lastRow,
                $c2 + @{ $co->{list} }
            ) if $self->{columns}[$c]{conditionalFormatting};

            $c2 += @{ $co->{list} };

        }

        else {

            if ( my $deferred = $self->{columns}[$c]{deferWritingTo} ) {
                _deferWritingData( $self->{columns}[$c],
                    $deferred, $wb, $ws, $row, $c2 );
            }
            else {
                foreach my $y ( $self->{columns}[$c]->rowIndices ) {
                    my ( $value, $format, $formula, @more ) =
                      $cell[$c]->( 0, $y );
                    if (@dataDecorationsByColumn) {
                        if ( my $dd = $dataDecorationsByColumn[ $c2 - $col ] ) {
                            $format = $wb->getFormat( $format, $dd );
                        }
                    }
                    if (@more) {
                        $ws->repeat_formula( $row + $y, $c2, $formula,
                            $format, @more );
                    }
                    elsif ($formula) {
                        $ws->write_formula( $row + $y, $c2, $formula, $format,
                            $value );
                    }
                    else {
                        $value = "=$value"
                          if $value
                          and $value eq '#VALUE!' || $value eq '#N/A'
                          and $wb->formulaHashValues;
                        $ws->write( $row + $y, $c2, $value, $format );
                        if ($comment) {
                            $ws->write_comment(
                                $row + $y,
                                $c2,
                                (
                                    map {
                                        ref $_ eq 'ARRAY'
                                          ? join "\n", @$_
                                          : $_;
                                      } ref $comment eq 'HASH'
                                    ? $comment->{text}
                                    : $comment
                                ),
                                x_scale => ref $comment eq 'HASH'
                                  && $comment->{x_scale} ? $comment->{x_scale}
                                : 3,
                            );
                            undef $comment;
                        }
                    }
                }
            }

            if (    $self->{columns}[$c]{validation}
                and ( my $l = $self->{lines} || $self->{columns}[$c]{lines} )
                and $wb->{validation}
                and $wb->{validation} =~ /withlinesmsg/i )
            {
                $self->{columns}[$c]{validation}{input_message} ||=
                  ref $l eq 'ARRAY' ? join "\n", @$l : $l;
            }

            $self->{columns}[$c]
              ->dataValidation( $wb, $ws, $row, $c2, $row + $lastRow, $c2 )
              if $self->{columns}[$c]{validation};

            $self->{columns}[$c]
              ->conditionalFormatting( $wb, $ws, $row, $c2, $row + $lastRow,
                $c2 )
              if $self->{columns}[$c]{conditionalFormatting};

            ++$c2;
        }

    }

    {
        my $dataset;
        $dataset = $self->{dataset}{ $self->{number} } if $self->{number};
        $dataset =
          $self->{dataset}{defaultClosure}->( $self->{number}, $wb, $ws )
          if !$dataset && $self->{dataset}{defaultClosure};
        $dataset = $dataset->( $self->{number}, $wb, $ws )
          if ref $dataset eq 'CODE';
        my $scribbleFormat = $wb->getFormat('scribbles');
        foreach ( 1 .. 1 ) {    # Scribble columns
            my @note;
            if ($dataset) {
                my $nd = $dataset->[$c2];
                if ( ref $nd eq 'HASH' ) {
                    @note =
                      map {
                        local $_ = $_;
                        s/.*\n//s;
                        s/[^A-Za-z0-9 -]/ /g;
                        s/- / /g;
                        s/ +/ /g;
                        s/^ //;
                        s/ $//;
                        $nd->{$_};
                      } $self->{rows}
                      ? @{ $self->{rows}{list} }
                      : ( $self->{singleRowName}
                          || _shortNameRow( $self->{name} ) );
                }
                elsif ( ref $nd eq 'ARRAY' ) {
                    @note = @$nd;
                    shift @note;
                }
            }
            foreach my $y ( 0 .. $lastRow ) {
                $ws->write( $row + $y, $c2, $note[$y], $scribbleFormat );
            }
            ++$c2;
        }
    }

    $row += $lastRow;
    $_->( $self, $wb, $ws, \$row, $col )
      foreach map { @{ $self->{postWriteCalls}{$_} }; }
      grep { $self->{postWriteCalls}{$_} } 'obj', $wb;
    $self->requestForwardLinks( $wb, $ws, \$row, $col ) if $wb->{forwardLinks};
    ++$row unless $self->{noSpaceBelow};
    $ws->{nextFree} = $row unless $ws->{nextFree} > $row;

    $ws;

}

sub _deferWritingData {
    my ( $obj, $deferred, $wb, $ws, $row, $col ) = @_;
    my $target =
      UNIVERSAL::can( $deferred->{location}, 'wsWrite' )
      ? $deferred->{location}
      : $deferred;
    my $lc = $obj->lastCol;
    my $lr = $obj->lastRow;
    push @{ $target->{postWriteCalls}{$wb} }, sub {
        my ( $obj_pwc, $wb_pwc, $ws_pwc, $rowref_pwc, $col_pwc ) = @_;
        my $cell = $obj->wsPrepare( $wb_pwc, $ws );
        for ( my $x = 0 ; $x <= $lc ; ++$x ) {
            for ( my $y = 0 ; $y <= $lr ; ++$y ) {
                my ( $value, $format, $formula, @more ) = $cell->( $x, $y );
                $ws->repeat_formula( $row + $y, $col + $x, $formula, $format,
                    @more );
            }
        }
    };
}

1;
