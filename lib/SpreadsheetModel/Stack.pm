
=head Development note

This file contains SpreadsheetModel::Stack, SpreadsheetModel::View and SpreadsheetModel::Constant.

=cut

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

require SpreadsheetModel::Dataset;

package SpreadsheetModel::Stack;
our @ISA = qw(SpreadsheetModel::Dataset);

use SpreadsheetModel::Label;
use Spreadsheet::WriteExcel::Utility;

sub objectType {
    @{ $_[0]{sources} } > 1 ? 'Combine tables' : 'Copy cells';
}

sub populateCore {
    my ($self) = @_;
    $self->{core}{sources} = [ map { $_->getCore } @{ $self->{sources} } ];
}

sub findxy {
    my ( $self, $z, $x, $y ) = @_;
    unless ( defined $x ) {
        ($x) =
          defined $z
          ? (
            $self->{cols}
            ? grep { $z eq $self->{cols}{list}[$_] } $self->colIndices
            : ()
          )
          : $self->{cols} ? ()
          :                 (0);
        return ( $x, $y ) if defined $x;
    }
    return if defined $y;
    ($y) =
      defined $z
      ? (
        $self->{rows}
        ? grep { $z eq $self->{rows}{list}[$_] } $self->rowIndices
        : ()
      )
      : $self->{rows} ? ()
      :                 (0);
    return ( $x, $y ) if defined $y;
    return;
}

sub check {
    my ($self) = @_;
    return 'Broken stack'
      unless 'ARRAY' eq ref $self->{sources};
    $self->{rows} = $#{ $self->{sources} } ? 0 : $self->{sources}[0]->{rows}
      unless defined $self->{rows};
    $self->{cols} = $#{ $self->{sources} } ? 0 : $self->{sources}[0]->{cols}
      unless defined $self->{cols};
    push @{ $self->{sourceLines} }, @{ $self->{sources} };
    $self->{arithmetic} =
      '= ' . join( ' or ', map { "IV1$_" } 0 .. $#{ $self->{sources} } );
    $self->{arguments} =
      { map { ( "IV1$_" => $self->{sources}[$_] ); }
          0 .. $#{ $self->{sources} } };
    if ( !$#{ $self->{sources} } ) {

        if ( !$self->{name} || $self->{name} =~ /^Untitled/ ) {
            0 and warn $self->{name} . ' ' . $self->{debug};
            my $n = $self->{sources}[0]{name};
            $self->{name} =
                 $self->{cols} == $self->{sources}[0]{cols}
              && $self->{rows} == $self->{sources}[0]{rows}
              ? new SpreadsheetModel::Label( "$n (copy)", $n )
              : $n;
        }
        if ( !$self->{defaultFormat}
            && ( my $df = $self->{sources}[0]{defaultFormat} ) )
        {
            $df =~ s/(soft|hard)/copy/g unless ref $df;
            $self->{defaultFormat} = $df;
        }
    }
    my @map;
    $self->{sources} = [ grep { $_ } @{ $self->{sources} } ];
    for my $source ( @{ $self->{sources} } ) {
        my $xabs = $source->{cols} != $self->{cols}
          && ( !$self->{cols}{accepts}
            || !grep { $_ == $source->{cols} } @{ $self->{cols}{accepts} } );
        my $yabs = $source->{rows} != $self->{rows};
        for my $sx ( $source->colIndices ) {
            my @xy0;
            if ($xabs) {
                my $scol;
                $scol = $source->{cols}{list}[$sx] if $source->{cols};
                next unless @xy0 = $self->findxy($scol);
            }
            else {
                @xy0 = ($sx);
            }
            for my $sy ( $source->rowIndices ) {
                my @xy;
                if ($yabs) {
                    my $srow;
                    $srow = $source->{rows}{list}[$sy] if $source->{rows};
                    next unless @xy = $self->findxy( $srow, @xy0 );
                }
                else {
                    @xy = ( $xy0[0], $sy );
                }
                next if $map[ $xy[0] ][ $xy[1] ];
                $map[ $xy[0] ][ $xy[1] ] = [ $source, $sx, $sy, $xabs, $yabs ];
            }
        }
    }

    $self->{map} = [@map];
    $self->SUPER::check;
}

sub wsPrepare {
    my ( $self, $wb, $ws ) = @_;

    my %formula;
    my %rowcol;
    my $broken;

    for ( @{ $self->{sources} } ) {
        if ( ref $_ eq 'SpreadsheetModel::Constant'
            && !$_->{$wb} )
        {
            $rowcol{ 0 + $_ } = [ $_, $wb->getFormat( $_->{defaultFormat} ) ];
        }
        elsif ( ref $_ eq 'SpreadsheetModel::View'
            && !$_->{$wb} )
        {
            $formula{ 0 + $_ } = $_->wsPrepare( $wb, $ws );
            $_->{name} = $1 . $_->{name}
              if $_->{name} !~ /^[0-9]+[0-9a-z]*\./i
              && $_->{sources}[0]{name} =~ /^([0-9]+[0-9a-z]*\.\s*)/i;
        }
        else {
            my ( $srcsheet, $srcr, $srcc ) =
              $_->wsWrite( $wb, $ws, undef, undef, 1 );
            $broken =
              "UNFEASIBLE LINK to source for $self->{name} $self->{debug}"
              unless $srcsheet;
            $formula{ 0 + $_ } = $ws->store_formula(
                !$srcsheet || $ws == $srcsheet
                ? '=IV1'
                : '=' . q"'" . $srcsheet->get_name . q"'!IV1"
            );
            $rowcol{ 0 + $_ } = [ $srcr, $srcc ];
        }
    }

    return sub { die $broken; }
      if $broken;

    my $format = $wb->getFormat( $self->{defaultFormat} || '0.000copy' );
    my $unavailable = $wb->getFormat('unavailable');

    sub {

        return ( '', $unavailable )
          unless $self->{map}[ $_[0] ] && $self->{map}[ $_[0] ][ $_[1] ];

        my ( $source, $sx, $sy, $xabs, $yabs ) =
          @{ $self->{map}[ $_[0] ][ $_[1] ] };

        $xabs = 1 if $_[2];
        $yabs = 1 if $_[3];

        return $formula{ 0 + $source }->( $sx, $sy, $xabs, $yabs )
          if 'CODE' eq ref $formula{ 0 + $source };

        my ( $row, $col ) = @{ $rowcol{ 0 + $source } };

        return $row->{byrow}
          ? $row->{data}[$sy][$sx]
          : $row->{data}[$sx][$sy], $col
          unless $formula{ 0 + $source };

        '', $format, $formula{ 0 + $source },
          IV1 => xl_rowcol_to_cell( $row + $sy, $col + $sx, $yabs, $xabs );

    };
}

#

package SpreadsheetModel::View;
our @ISA = qw/SpreadsheetModel::Stack/;

sub check {
    $_[0]->SUPER::check;
    return 'Broken view' if $#{ $_[0]->{sources} };
    $_[0]->{name} = "$_[0]->{sources}[0]{name} (extracts)"
      if !$_[0]->{name} || $_[0]->{name} eq $_[0]->{debug};
    return;
}

sub objectType {
    'Select from table';
}

sub wsUrl {
    my $self = shift;
    $self->SUPER::wsUrl(@_) || $self->{sources}[0]->wsUrl(@_);
}

sub addForwardLink {
    my $self = shift;
    $self->SUPER::addForwardLink(@_);
    $self->{sources}[0]->addForwardLink(@_);
}

#

package SpreadsheetModel::Constant;
our @ISA = qw/SpreadsheetModel::Dataset/;

sub populateCore {
    my ($self) = @_;
    $self->{core}{$_} = $self->{$_}
      foreach grep { exists $self->{$_}; } qw(data);
}

sub check {
    $_[0]{defaultFormat} ||= '0.000con';
    return "No data in constant $_[0]{name}" unless 'ARRAY' eq ref $_[0]{data};
    $_[0]{arithmetic} = '[ ' . join(
        ( $_[0]{byrow} ? ",\n" : ', ' ),
        map {
                !defined $_ ? 'undef'
              : $_     eq ''      ? q['']
              : ref $_ eq 'ARRAY' ? (
                '[ '
                  . join( ', ',
                    map { !defined $_ ? 'undef' : $_ eq '' ? q[''] : "$_" }
                      @$_ )
                  . ' ]'
              )
              : "$_"
        } @{ $_[0]{data} }
    ) . ' ]';
    $_[0]->SUPER::check;
}

sub objectType {
    $_[0]{specialType} || 'Fixed data';
}

1;
