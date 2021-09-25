package SpreadsheetModel::CalcBlock;

# Copyright 2015-2021 Franck Latrémolière and others.
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

use Exporter qw(import);
our @EXPORT = qw(CalcBlock);

sub CalcBlock {
    unshift @_, 'SpreadsheetModel::CalcBlock';
    goto &SpreadsheetModel::Object::new;
}

use Spreadsheet::WriteExcel::Utility;
use SpreadsheetModel::Object ':_util';
our @ISA = qw(SpreadsheetModel::Objectset);

sub wsUrl {
    my $self = shift;
    $self->{items}[0]->wsUrl(@_);
}

sub check {

    my ($self) = @_;
    return 'No item list for CalcBlock' unless 'ARRAY' eq ref $self->{items};
    my $defaultFormat = $self->{defaultFormat} || '0soft';
    my ( @items, %knownMap, $addItems );
    my $addDataset;
    $addDataset = sub {
        ( local $_ ) = @_;
        return $knownMap{ 0 + $_ } =
          $addDataset->( SpreadsheetModel::Stack->new( sources => [$_] ) )
          if $_->{location};
        my $args;
        $args = $_->{arguments};
        my %argMap;
        %argMap = map { ( 0 + $_ => $_ ); } values %$args if $args;
        my @externalArgKeys = grep { !$knownMap{$_}; } keys %argMap;
        if (@externalArgKeys) {

            if ( keys %argMap == 1 ) {
                $_->{singleExternalSource} = 1;
            }
            elsif (
                   $self->{consolidate}
                && UNIVERSAL::isa( $_, 'SpreadsheetModel::Arithmetic' )
                && !grep {
                         $_->{rows}
                      or $_->{cols} || $self->{cols}
                      and $_->{cols} != $self->{cols};
                } values %argMap
              )
            {
                $addItems->( $argMap{$_} ) foreach @externalArgKeys;
            }
            else {
                return $knownMap{ 0 + $_ } = $addDataset->(
                    SpreadsheetModel::Stack->new( sources => [$_] ) );
            }
        }
        if ($args) {
            while ( my ( $k, $v ) = each %$args ) {
                $args->{$k} = $knownMap{ 0 + $v }
                  if defined $knownMap{ 0 + $v };
            }
        }
        $_->{location} = $self unless $self->{virtual};
        push @items, $_;
        $knownMap{ 0 + $_ } = $_;
    };
    $addItems = sub {
        local @_ = @_;    # to avoid side effects
        $_[$#_] = { name => $_[$#_] }
          if !ref $_[$#_]
          || UNIVERSAL::isa( $_[$#_], 'SpreadsheetModel::Label' );
        my ( $key, @arglist, %argmap );
        foreach (@_) {
            if ( !ref $_ ) {
                $key = $_;
                next;
            }
            if ( ref $_ eq 'HASH' ) {
                if ( $_->{arithmetic} ) {
                    $_->{arguments} =
                      { %argmap, $_->{arguments} ? %{ $_->{arguments} } : () };
                }
                else {
                    my @terms;
                    my %formulaArg = map {
                        my $cell = 'A' . ( $_ + 1 );
                        push @terms, $cell;
                        $cell => $arglist[$_];
                    } 0 .. $#arglist;
                    my $formula = join '+', @terms;
                    $formula = "ROUND($formula,$_->{rounding})"
                      if defined $_->{rounding};
                    $_->{arguments}  = \%formulaArg;
                    $_->{arithmetic} = '=' . $formula;
                }
                $_->{defaultFormat} ||= $defaultFormat;
                $_->{cols} = $self->{cols} ||= $_->{arguments}{A1}{cols};
                $addDataset->( $_ = SpreadsheetModel::Arithmetic->new(%$_) );
            }
            elsif ( ref $_ eq 'ARRAY' ) {
                $_ = $addItems->(@$_);
            }
            else {
                die "$_->{name} not allowed in"
                  . " CalcBlock $self->{name} $self->{debug}"
                  . " because it needs row labels ($_->{rows})"
                  if $_->{rows};
                if ( defined $self->{cols} ) {
                    unless ( !$_->{cols} && !$self->{cols}
                        || $_->{cols} == $self->{cols} )
                    {
                        die join "\n",
                          "Mismatch in CalcBlock $self->{name} $self->{debug}",
                          "Columns in CalcBlock: $self->{cols}",
                          "Columns in $_->{name} $_->{debug}: $_->{cols}";
                    }
                }
                else {
                    $self->{cols} = $_->{cols} || 0;
                }
                $addDataset->($_);
            }
            if ($key) {
                $self->{$key} = $argmap{$key} = $_;
                undef $key;
            }
            push @arglist, $_;
        }
        $arglist[$#arglist];
    };

    eval { $addItems->( @{ $self->{items} } ); };
    return $@ if $@;
    $self->{items} = \@items;
    return;

}

sub wsWrite {

    my ( $self, $wb, $ws, $row, $col ) = @_;

    return values %{ $self->{$wb} } if $self->{$wb};
    $self->{$wb} ||= {};

    $self->{cols}->wsPrepare( $wb, $ws ) if $self->{cols};

    while (1) {

        unless ( defined $row && defined $col ) {
            ( $row, $col ) = ( ( $ws->{nextFree} ||= -1 ) + 1, 0 );
        }

        foreach ( @{ $self->{items} } ) {
            die "$_ $_->{name} is already in the workbook"
              . " and cannot be written as part of $self $self->{name}"
              if $_->{$wb};
            $_->wsPrepare( $wb, $ws );
            $_->{$wb} ||= {};    # Placeholder for other rows
        }

        last if !$ws->{nextFree} || $ws->{nextFree} < $row;
        delete $_->{$wb} foreach @{ $self->{items} };
        undef $row;

    }

    $self->{$wb}{$ws} = $ws;

    if ( $wb->{logger} and my $oldName = $self->{name} ) {
        my $number = $self->addTableNumber( $wb, $ws );
        foreach ( @{ $self->{items} } ) {
            my $n = $_->{name};
            $_->{name} =
              new SpreadsheetModel::Label( $n, "$number$n (in $oldName)" );
            $wb->{logger}->log($_);
        }
    }

    $ws->set_row( $row, $wb->{captionRowHeight} );
    $ws->write( $row++, $col, "$self->{name}", $wb->getFormat('caption') );

    ++$row;    # Blank line
    ++$col;

    if ( $self->{cols} ) {
        $ws->write(
            $row,
            $col + $_,
            _shortNameCol( $self->{cols}{list}[$_], $wb, $ws ),
            $wb->getFormat(
                !$self->{cols}{groups} || defined $self->{cols}{groupid}[$_]
                ? ( $self->{cols}{defaultFormat} || 'thc' )
                : 'thg'
            )
        ) foreach $self->{cols}->indices;
        ++$row;
    }
    elsif ( !exists $self->{singleColName} ) {
        $ws->write(
            $row, $col,
            _shortNameRow( $self->{name} ),
            $wb->getFormat('thc')
        );
        ++$row;
    }
    elsif ( $self->{singleColName} ) {
        $ws->write(
            $row, $col,
            _shortNameRow( $self->{singleColName} ),
            $wb->getFormat('th')
        );
        ++$row;
    }

    my @cellClosures;

    {
        my $r = $row;
        foreach my $item ( @{ $self->{items} } ) {
            push @cellClosures, $item->wsPrepare( $wb, $ws );
            @{ $item->{$wb} }{qw(worksheet row col)} =
              ( $ws, $r++, $col );
        }
    }

    my $scribbleFormat = $wb->getFormat('scribbles');
    my $scribbleColumn =
      1 + $col + ( $self->{cols} ? $#{ $self->{cols}{list} } : 0 );

    my @sourceLines;
    my @formulas =
      map {
        $_->{singleExternalSource} || !$_->{arguments}
          ? 'A0='
          : "A0$_->{arithmetic}"
      } @{ $self->{items} };
    if ( grep { $_ } @formulas ) {
        @sourceLines = (
            _rewriteFormulas(
                \@formulas,
                [
                    map {
                        $_->{singleExternalSource} || !$_->{arguments}
                          ? { A0 => $_ }
                          : { A0 => $_, %{ $_->{arguments} }, };
                    } @{ $self->{items} }
                ]
            )
        );
    }

    my $thFormat = $wb->getFormat('th');
    for ( my $i = 0 ; $i < @{ $self->{items} } ; ++$i ) {
        my $item = $self->{items}[$i];
        $ws->write_string( $row, $col - 1, _shortNameRow( $item->{name} ),
            $thFormat );
        local $_ = $formulas[$i]
          . ( $item->{singleExternalSource} ? $item->{arithmetic} : '' );
        s/(x[0-9]+)[=\s]+/$1 = /;
        s/ = $/ = constant/;
        if ( $item->{singleExternalSource} ) {
            unless (s/(=\s*)A[0-9]+$/$1$item->{sourceLines}[0]{name}/) {
                s/([(,])A[0-9]+([),])/$1$item->{sourceLines}[0]{name}$2/g;
                s/\bA[0-9]+\b/($item->{sourceLines}[0]{name})/g;
            }
            $ws->write_url( $row, $scribbleColumn,
                $item->{sourceLines}[0]->wsUrl($wb),
                $_, $wb->getFormat('link') );
            (
                $item->{sourceLines}[0]{location} && UNIVERSAL::can(
                    $item->{sourceLines}[0]{location}, 'wsWrite'
                  )
                ? $item->{sourceLines}[0]{location}
                : $item->{sourceLines}[0]
              )->addForwardLink($item)
              if $wb->{findForwardLinks};
        }
        else {
            $ws->write_string( $row, $scribbleColumn, $_,
                $wb->getFormat('text') );
        }
        my $cell = $cellClosures[$i];
        @{ $item->{$wb} }{qw(worksheet row col)} = ( $ws, $row, $col );
        foreach my $y ( $item->rowIndices ) {
            foreach my $x ( $item->colIndices ) {
                my ( $value, $format, $formula, @more ) = $cell->( $x, $y );
                if (@more) {
                    $ws->repeat_formula( $row,
                        $col + $x, $formula, $format, @more );
                }
                elsif ($formula) {
                    $ws->write_formula( $row, $col + $x, $formula, $format );
                }
                else {
                    $ws->write( $row, $col + $x, $value, $format );
                }
            }
            ++$row;
        }
        $_->( $item, $wb, $ws, \$row, $col )
          foreach map { @{ $item->{postWriteCalls}{$_} }; }
          grep { $item->{postWriteCalls}{$_} } 'obj', $wb;
    }
    if ( $wb->{forwardLinks} ) {
        --$row;
        $self->requestForwardLinks( $wb, $ws, \$row, $col );
        ++$row;
    }
    $ws->{nextFree} = $row unless $ws->{nextFree} > $row;

}

1;
