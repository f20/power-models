package SpreadsheetModel::Columnset;

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

use SpreadsheetModel::Object ':_util';
our @ISA = qw(SpreadsheetModel::Object);

use constant {
    OLD_STYLE_SCRIBBLES  => undef,
    NUM_SCRIBBLE_COLUMNS => 1,
    BLANK_LINE           => 1,
};

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
    return 'Broken columnset' unless 'ARRAY' eq ref $self->{columns};
    my @columns = @{ $self->{columns} };
    return 'Empty columnset' unless @columns;
    $self->{lines} = [ SpreadsheetModel::Object::splitLines( $self->{lines} ) ]
      if $self->{lines};
    my $rows;
    my $colOffset = 0;
    $self->{anonRow} = 0;

    foreach (@columns) {
        0 and warn "$self->{name} $self->{debug} $_->{name} $self->{debug}";
        if ( defined $rows ) {
            return <<ERR unless !$_->{rows} && !$rows || $_->{rows} == $rows;
Mismatch in Columnset
$self->{name} $self->{debug} $rows
$_->{name} $_->{debug} $_->{rows}
ERR
        }
        else {
            if ( $_->{rows} ) {
                $rows = $_->{rows};
            }
            else {
                $self->{anonRow} = 1;
                $rows = 0;
            }
        }
        $_->{location} = $self;
        $_->{name} =
          new SpreadsheetModel::Label( $_->{name},
            "$_->{name} (in $self->{name})" )
          if index( $_->{name}, " (in $self->{name})" ) < 0;
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

sub wsWrite {

    my ( $self, $wb, $ws, $row, $col, $noCopy ) = @_;

    if ( $self->{$wb} ) {
        return if !$wb->{copy} || $noCopy || $self->{$wb}{$ws};
        return (
            $self->{$wb}{$ws} = new SpreadsheetModel::Columnset(
                name    => "$self->{name} (copy)",
                columns => [
                    map {
                        $_->{$wb}{$ws} =
                          new SpreadsheetModel::Stack( sources => [$_] );
                    } @{ $self->{columns} }
                ]
            )
        )->wsWrite( $wb, $ws, $row, $col );
    }

    if (   $self->{location}
        && $wb->{ $self->{location} }
        && $wb->{ $self->{location} } ne $ws )
    {
        return $self->wsWrite( $wb, $wb->{ $self->{location} }, undef, undef,
            1 );
    }

    if (   !$self->{location}
        and $wb->{dataSheet}
        and $wb->{dataSheet} ne $ws
        and !grep { ref $_ ne 'SpreadsheetModel::Dataset' }
        @{ $self->{columns} } )
    {
        return $self->wsWrite( $wb, $wb->{dataSheet}, undef, undef, 1 );
    }

    my @dataColumns =
      grep { ref $_ eq 'SpreadsheetModel::Dataset' } @{ $self->{columns} };

    if ( ( my $inSheet = $wb->{dataSheet} || $wb->{inputSheet} )
        && @dataColumns )
    {
        if ( $ws != $inSheet && !$self->{doNotCopyInputColumns} ) {
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
                    if ( $mapping{$_} ) { $mapping{$_}; }
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
            ];

            $data->wsWrite( $wb, $inSheet );

        }
    }

    my @cell;

    my $lastCol = $#{ $self->{columns} };

    my $lastRow = $self->{rows} ? $#{ $self->{rows}{list} } : 0;

    my $headerCols =
         $self->{rows}
      || !exists $self->{singleRowName}
      || $self->{singleRowName} ? 1 : 0;

    my @sourceLines;

    {
        my @formulas =
          map { $_->{arguments} ? $_->{arithmetic} : '' } @{ $self->{columns} };
        if ( grep { $_ } @formulas ) {
            @sourceLines =
              _rewriteFormulae( \@formulas,
                [ map { $_->{arguments} } @{ $self->{columns} } ] );
            unless ( $self->{formulasDone} ) {
                $self->{formulasDone} = 1;
                unless ( $wb->{noLinks} ) {
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

    my $headerLines = $self->{name} ? BLANK_LINE : 0;

    $headerLines += $dualHeaded ? 2 : 1 unless $self->{noHeaders};

    if ( $self->{name} ) {
        ++$headerLines if OLD_STYLE_SCRIBBLES;
        ++$headerLines;
        $headerLines += @{ $self->{lines} } if $self->{lines};
        $headerLines += 1 + @sourceLines if !$wb->{noLinks} && @sourceLines;
    }

    $self->{$wb}{$ws} = 1;

    my @hideRange;

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
            $thecol->wsPrepare( $wb, $ws );

# This is a placeholder assignment, only OK for other columns of the same columnset to use.
            @{ $thecol->{$wb} }{qw(worksheet row col)} = ( 0, -666, -666 );
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
            $cell[$_] = $thecol->wsPrepare( $wb, $ws );
            @{ $thecol->{$wb} }{qw(worksheet row col)} = (
                $ws,
                $row + $headerLines +
                  ( $thecol->{rows} ? $self->{anonRow} : 0 ),
                $c2 + $headerCols
            );
            $c2 +=
              $thecol->{cols}
              ? @{ $thecol->{cols}{list} }
              : 1;
        }
    }

    if (@hideRange) {
        $ws->set_row( $_, undef, undef, 1 )
          foreach $hideRange[0] .. $hideRange[1] - 1;
    }

    my $number = $wb->{logger}
      && $self->{name} ? $self->addTableNumber( $wb, $ws ) : undef;

    $self->{name} .= " ($self->{debug})"
      if $wb->{debug} && 0 > index $self->{name}, $self->{debug};

    return if $self->{rows} && !@{ $self->{rows}{list} };

    my $dataset;
    $dataset = $self->{dataset}{ $self->{number} } if $self->{number};

    if ( $self->{name} ) {

        # $ws->set_row( $row, 24 );
        $ws->write_string( $row++, $col, "$self->{name}",
            $wb->getFormat('caption') );

        if (OLD_STYLE_SCRIBBLES) {
            my $note;
            $note = $dataset->[0]{_note} if $dataset;
            $ws->write_string(
                $row++, $col,
                $note || '',
                $wb->getFormat('scribbles')
            );
        }

        if ( $self->{lines} || !$wb->{noLinks} && @sourceLines ) {
            my $textFormat = $wb->getFormat('text');
            my $linkFormat = $wb->getFormat('link');
            my $xc         = 0;
            my @arrayLines;
            foreach (
                $self->{lines} ? @{ $self->{lines} } : (),
                !$wb->{noLinks} && @sourceLines
                ? ( 'Data sources:', @sourceLines )
                : ()
              )
            {
                if ( ref $_ eq 'ARRAY' ) {
                    push @arrayLines, $_;
                }
                elsif ( ref($_) =~ /^SpreadsheetModel::/ ) {
                    my $na = 'x' . ( ++$xc ) . " = $_->{name}";
                    if ( my $url = $_->wsUrl($wb) ) {
                        $ws->write_url( $row++, $col, $url, $na, $linkFormat );
                        (
                            $_->{location}
                              && ref $_->{location} eq
                              'SpreadsheetModel::Columnset'
                            ? $_->{location}
                            : $_
                          )->addForwardLink($self)
                          if $wb->{findForwardLinks};
                    }
                    else {
                        $ws->write_string( $row++, $col, $na, $textFormat );
                    }
                }
                elsif (/^(https?|mailto:)/) {
                    $ws->write_url( $row++, $col, "$_", "$_", $linkFormat );
                }
                else {
                    $ws->write_string( $row++, $col, "$_", $textFormat );
                }
            }
            if (@arrayLines) {
                foreach (@arrayLines) {
                    my $c3 = $col;
                    $ws->write_string( $row, $c3++, "$_->[0]",
                        $wb->getFormat('textwrap') )
                      if $headerCols;

                    foreach my $cn ( 0 .. $lastCol ) {
                        my $wrapFormat = $wb->getFormat(
                            $self->{columns}[$cn]{cols}
                              && $#{ $self->{columns}[$cn]{cols}{list} } > 2
                            ? 'textlrap'
                            : 'textwrap'
                        );
                        $ws->write_string( $row, $c3++,
                            "$_->[ $cn + $headerCols ]", $wrapFormat );
                        $ws->write( $row, $c3++, undef, $wrapFormat )
                          foreach 1 .. $self->{columns}[$cn]{cols}
                          && $#{ $self->{columns}[$cn]{cols}{list} };
                    }
                    ++$row;
                }
            }

        }

    }

    ++$row if BLANK_LINE && $self->{name};

    unless ( $self->{noHeaders} ) {

        my $c4 = $col + $headerCols;
        $row++ if $dualHeaded;
        foreach ( 0 .. $lastCol ) {

            my $colShortName = _shortNameCol $self->{columns}[$_]{name};

            if ( ( my $co = $self->{columns}[$_]{cols} )
                && $#{ $self->{columns}[$_]{cols}{list} } )
            {

                foreach ( 0 .. $#{ $co->{list} } ) {
                    $ws->write(
                        $row - 1,
                        $c4 + $_,
                        $_ ? undef : $colShortName,
                        $wb->getFormat(
                            $#{ $co->{list} } > 2 ? 'thla' : 'thca'
                        )
                    );
                    $ws->write(
                        $row,
                        $c4 + $_,
                        _shortNameCol( $co->{list}[$_] ),
                        $wb->getFormat(
                            !$co->{groups}
                              || defined $co->{groupid}[$_] ? 'thc' : 'thg'
                        )
                    );
                }
                $c4 += @{ $co->{list} };
            }
            else {
                $ws->write( $row, $c4++, $colShortName, $wb->getFormat('thc') );
            }

            if ( $wb->{logger} ) {
                if ($number) {
                    my $n = $self->{columns}[$_]{name};
                    $self->{columns}[$_]{name} =
                      new SpreadsheetModel::Label( $n, $number . $n );
                }
                $wb->{logger}->log( $self->{columns}[$_] );
            }

        }

        ++$row;

    }

    $col += $headerCols;

    if ( $self->{rows} ) {
        my $thformat = $wb->getFormat( $self->{rows}{defaultFormat} || 'th' );
        my $thgformat = $wb->getFormat('thg');
        $ws->write( $row, $col - 1, '', $thgformat ) if $self->{anonRow};
        for ( my $r = 0 ; $r <= $lastRow ; ++$r ) {
            if ( !$self->{rows}{groups}
                || defined $self->{rows}{groupid}[$r] )
            {
                $ws->write( $row + $self->{anonRow} + $r,
                    $col - 1, _shortNameRow( $self->{rows}{list}[$r] ),
                    $thformat );
            }
            else {
                $ws->write( $row + $self->{anonRow} + $r,
                    $col - 1, _shortNameRow( $self->{rows}{list}[$r] ),
                    $thgformat );
            }
        }
    }

    elsif ( !exists $self->{singleRowName} ) {
        my $srn = _shortNameRow $self->{name};
        $srn =~ s/^[0-9]+[a-z]*\.\s+//i;
        $srn =~ s/\s*\(copy\)$//i;

=head comment

Perhaps whoever auto-creates table numbers and names of Stacks
that are just copies should use SpreadsheetModel::Label and then we could
use ->shortName here.

=cut

        $ws->write( $row, $col - 1, $srn, $wb->getFormat('th') );
    }

    elsif ( $self->{singleRowName} ) {
        $ws->write(
            $row, $col - 1,
            _shortNameRow( $self->{singleRowName} ),
            $wb->getFormat('th')
        );
    }

    my $c2 = $col;
    foreach my $c ( 0 .. $lastCol ) {
        if ( my $co = $self->{columns}[$c]{cols} ) {
            foreach my $y ( $self->{columns}[$c]->rowIndices ) {
                foreach my $x ( $self->{columns}[$c]->colIndices ) {
                    my ( $value, $format, $formula, @more ) =
                      $cell[$c]->( $x, $y );
                    if (@more) {
                        $ws->repeat_formula(
                            $row + $y + (
                                  $self->{columns}[$c]{rows}
                                ? $self->{anonRow}
                                : 0
                            ),
                            $c2 + $x,
                            $formula, $format, @more
                        );
                    }
                    elsif ($formula) {
                        $ws->write_formula(
                            $row + $y + (
                                  $self->{columns}[$c]{rows}
                                ? $self->{anonRow}
                                : 0
                            ),
                            $c2 + $x,
                            $formula, $format, $value
                        );
                    }
                    else {
                        $ws->write(
                            $row + $y + (
                                  $self->{columns}[$c]{rows}
                                ? $self->{anonRow}
                                : 0
                            ),
                            $c2 + $x,
                            $value, $format
                        );
                    }
                }
            }
            $self->{columns}[$c]->dataValidation( $wb, $ws, $row, $c2,
                $row + $lastRow +
                  ( $self->{columns}[$c]{rows} ? $self->{anonRow} : 0 ) )
              if $self->{columns}[$c]{validation};
            $c2 += @{ $co->{list} };
        }
        else {
            foreach my $y ( $self->{columns}[$c]->rowIndices ) {
                my ( $value, $format, $formula, @more ) = $cell[$c]->( 0, $y );
                if (@more) {
                    $ws->repeat_formula( $row + $y +
                          ( $self->{columns}[$c]{rows} ? $self->{anonRow} : 0 ),
                        $c2, $formula, $format, @more );
                }
                elsif ($formula) {
                    $ws->write_formula( $row + $y +
                          ( $self->{columns}[$c]{rows} ? $self->{anonRow} : 0 ),
                        $c2, $formula, $format, $value );
                }
                else {
                    $ws->write( $row + $y +
                          ( $self->{columns}[$c]{rows} ? $self->{anonRow} : 0 ),
                        $c2, $value, $format );
                }
            }
            $self->{columns}[$c]->dataValidation(
                $wb,
                $ws,
                $row,
                $c2,
                $row + $lastRow +
                  ( $self->{columns}[$c]{rows} ? $self->{anonRow} : 0 ),
                $c2
            ) if $self->{columns}[$c]{validation};
            ++$c2;
        }

    }

    unless ( $self->{noHeaders} ) {
        my $scribbleFormat = $wb->getFormat('scribbles');
        foreach ( 1 .. NUM_SCRIBBLE_COLUMNS ) {
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
                        $_ = "a$_" if /^[0-9]/;
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

    $row += $lastRow ? ( $lastRow + $self->{anonRow} ) : 0;
    $self->requestForwardLinks( $wb, $ws, \$row, $col ) if $wb->{forwardLinks};
    ++$row;
    $ws->{nextFree} = $row unless $ws->{nextFree} > $row;
    if ( $self->{postWriteCalls}{$wb} ) {
        $_->($self) foreach @{ $self->{postWriteCalls}{$wb} };
    }

}

1;
