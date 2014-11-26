package SpreadsheetModel::Notes;

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

require SpreadsheetModel::Object;
our @ISA = qw(SpreadsheetModel::Object);
use Spreadsheet::WriteExcel::Utility;

sub objectType {
    'Notes';
}

sub check {
    my ($self) = @_;
    $self->{lines} = [ SpreadsheetModel::Object::splitLines( $self->{lines} ) ];
    ( $self->{name} ) = splice @{ $self->{lines} }, 0, 2
      if $self->{lines}[0] && !$self->{lines}[1]
      and !defined $self->{name} || $self->{name} =~ /^Untitled/;
    return;
}

sub lastCol {
    0;
}

sub lastRow {
    $#{ $_[0]{lines} };
}

sub wsUrl {
    my ( $self, $wb ) = @_;
    return unless $self->{$wb};
    my ( $wo, $ro, $co ) = @{ $self->{$wb} }{qw(worksheet row col)};
    my $ce = xl_rowcol_to_cell( $ro + 1, $co );
    my $wn =
        $wo
      ? $wo->get_name
      : die( join "\n", "No worksheet for $self->{name}",
        $self->{debug}, "$self->{rows} x $self->{cols}" );
    "internal:'$wn'!$ce";
}

sub wsWrite {

    my ( $self, $wb, $ws, $row, $col ) = @_;

    return @{ $self->{$wb} }{qw(worksheet row col)} if $self->{$wb};

    if (   $self->{location}
        && $wb->{ $self->{location} } )
    {
        return $self->wsWrite( $wb, $wb->{ $self->{location} }, undef, undef,
            1 )
          if $wb->{ $self->{location} } != $ws;
    }

    ( $row, $col ) = ( ( $ws->{nextFree} ||= -1 ) + 1, 0 )
      unless defined $row && defined $col;

    my @result = @{ $self->{$wb} }{qw(worksheet row col)} = ( $ws, $row, $col );

    my $lastRow = $self->lastRow;

    if ( $self->{name} ) {
        my $n = $self->{name};
        $n = "$wb->{titlePrefix}: $n" if $wb->{titlePrefix};
        if ( local $_ = $wb->{titleAppend} ) {
            use bytes;
            s/^"/="$n/s or s/^/="$n"&/s;
            $n = $_;
        }

        # 'Not calculated' thing not working with .xlsx
        $ws->write( $row, $col, $n, $wb->getFormat('notes'),
            1
            ? ()
            : 'Not calculated'
              . ' (need to open in spreadsheet app and/or calculate now)' );
        $ws->set_row( $row, 21 );
        ++$row;
    }

    my $defaultFormat = $wb->getFormat( $self->{defaultFormat} || 'text' );

    for ( 0 .. $lastRow ) {
        my $rf = $self->{rowFormats}[$_];
        $ws->set_row( $row, 21 ) if $rf && $rf eq 'caption';
        $ws->write( $row++, $col, "$self->{lines}[$_]",
              $rf
            ? $wb->getFormat($rf)
            : $defaultFormat );
    }

    if ( !$wb->{noLinks} && $self->{sourceLines} ) {
        my $linkFormat = $wb->getFormat('link');
        foreach ( @{ $self->{sourceLines} } ) {
            my @cells = ref $_ eq 'ARRAY' ? @$_ : $_;
            for ( my $c = 0 ; $c < @cells ; ++$c ) {
                next unless defined( local $_ = $cells[$c] );
                if ( ref($_) =~ /^SpreadsheetModel::/ ) {
                    if ( my $url = $_->wsUrl($wb) ) {
                        $ws->write_url( $row, $col + $c, $url,
                            "$_->{name}" || $_->{lines}[0], $linkFormat );
                    }
                    else {
                        $ws->write_string( $row, $col + $c,
                            "$_->{name}" || $_->{lines}[0],
                            $defaultFormat );
                    }
                }
                else {    # formulas allowed here
                    $ws->write( $row, $col + $c, "$_", $defaultFormat );
                }
            }
            ++$row;
        }
    }

    $ws->{nextFree} = $row unless $ws->{nextFree} > $row;

    @result;

}

1;
