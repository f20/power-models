package SpreadsheetModel::CalcBlock;

=head Copyright licence and disclaimer

Copyright 2015-2017 Franck Latrémolière, Reckon LLP and others.

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

use Exporter qw(import);
our @EXPORT = qw(CalcBlock);

sub CalcBlock {
    unshift @_, 'SpreadsheetModel::CalcBlock';
    goto &SpreadsheetModel::Object::new;
}

use Spreadsheet::WriteExcel::Utility;
use SpreadsheetModel::Object ':_util';
our @ISA = qw(SpreadsheetModel::Objectset);

sub check {

    my ($self) = @_;
    return 'No item list for CalcBlock' unless 'ARRAY' eq ref $self->{items};
    my $defaultFormat = $self->{defaultFormat} || '0boldsoft';
    my %inBlock;
    my @items;
    my $add;
    $add = sub {
        local @_ = @_;    # to avoid side effects
        $_[$#_] = { name => $_[$#_] }
          if !ref $_[$#_]
          || UNIVERSAL::isa( $_[$#_], 'SpreadsheetModel::Label' );
        my ( $key, @args, %args, %redirected );
        foreach (@_) {
            if ( !ref $_ ) {
                $key = $_;
                next;
            }
            if ( ref $_ eq 'HASH' ) {
                if ( $_->{arithmetic} ) {
                    $_->{arguments} =
                      { %args, $_->{arguments} ? %{ $_->{arguments} } : () };
                }
                else {
                    my @terms;
                    my %formulaArg = map {
                        my $cell = 'A' . ( $_ + 1 );
                        push @terms, $cell;
                        $cell => $args[$_];
                    } 0 .. $#args;
                    my $formula = join '+', @terms;
                    $formula = "ROUND($formula,$_->{rounding})"
                      if defined $_->{rounding};
                    $_->{arguments}  = \%formulaArg;
                    $_->{arithmetic} = '=' . $formula;
                }
                $_->{defaultFormat} ||= $defaultFormat;
                $_->{cols} = $self->{cols};
                push @items, $_ = SpreadsheetModel::Arithmetic->new(%$_);
            }
            elsif ( ref $_ eq 'ARRAY' ) {
                $_ = $add->(@$_);
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
                $_ = SpreadsheetModel::Stack->new(
                    singleExternalSource => 1,
                    sources              => [$_],
                ) if $_->{location};
                my %uniques;
                %uniques = map { ( 0 + $_ => $_ ); } values %{ $_->{arguments} }
                  if $_->{arguments};
                if ( grep { !exists $inBlock{$_}; } keys %uniques ) {
                    if ( keys %uniques > 1 ) {
                        if (
                            $self->{consolidate}
                            && UNIVERSAL::isa(
                                $_, 'SpreadsheetModel::Arithmetic'
                            )
                            && !grep {
                                     $_->{rows}
                                  or $_->{cols} || $self->{cols}
                                  and $_->{cols} != $self->{cols};
                            } values %uniques
                          )
                        {
                            foreach my $k ( sort keys %{ $_->{arguments} } ) {
                                my $v = $_->{arguments}{$k};
                                next if exists $inBlock{ 0 + $v };
                                $_->{arguments}{$k} = $redirected{ 0 + $v } ||=
                                  $add->($v);
                            }
                        }
                        elsif ( !grep { exists $inBlock{$_}; } keys %uniques ) {
                            $_ = SpreadsheetModel::Stack->new(
                                singleExternalSource => 1,
                                sources              => [$_]
                            );
                        }
                        else {
                            die join "\n",
                              "Cannot use $_->{name}"
                              . ' due to external dependencies:',
                              map { $_->{name} } @uniques{
                                grep { !exists $inBlock{$_}; }
                                  keys %uniques
                              };
                        }
                    }
                    else {
                        $_->{singleExternalSource} = 1;
                    }
                }
                push @items, $_;
            }
            $_->{location} = $self;
            if ($key) {
                $self->{$key} = $args{$key} = $_;
                undef $key;
            }
            push @args, $_;
            undef $inBlock{ 0 + $_ };
        }
        $args[$#args];
    };

    eval { $add->( @{ $self->{items} } ); };
    return $@ if $@;
    $self->{items} = \@items;
    return;

}

sub wsWrite {

    my ( $self, $wb, $ws, $row, $col ) = @_;

    return if $self->{$wb};
    0
      and warn join "\n", $self->{name},
      map { "\t$_->{name}" } @{ $self->{items} };
    $self->{cols}->wsPrepare( $wb, $ws ) if $self->{cols};

    while (1) {

        unless ( defined $row && defined $col ) {
            ( $row, $col ) = ( ( $ws->{nextFree} ||= -1 ) + 1, 0 );
        }

        foreach ( @{ $self->{items} } ) {
            die "$_->{name} is already in the workbook"
              . " and cannot be written as part of $self->{name}"
              if $_->{$wb};
            $_->wsPrepare( $wb, $ws );
            $_->{$wb} ||= {};    # Placeholder for other rows
        }

        last if !$ws->{nextFree} || $ws->{nextFree} < $row;
        delete $_->{$wb} for @{ $self->{items} };
        undef $row;

    }

    $self->{$wb}{$ws} = 1;

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

    my @cells;

    {
        my $r = $row;
        foreach my $item ( @{ $self->{items} } ) {
            push @cells, $item->wsPrepare( $wb, $ws );
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
                $item->{sourceLines}[0]{location} && UNIVERSAL::isa(
                    $item->{sourceLines}[0]{location},
                    'SpreadsheetModel::Columnset'
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
        my $cell = $cells[$i];
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
            $row++;
        }
        $_->( $item, $wb, $ws, \$row, $col )
          foreach map { @{ $item->{postWriteCalls}{$_} }; }
          grep { $item->{postWriteCalls}{$_} } 'obj', $wb;
    }
    $ws->{nextFree} = $row unless $ws->{nextFree} > $row;

}

1;
