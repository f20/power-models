package SpreadsheetModel::Notes;

=head Copyright licence and disclaimer

Copyright 2008-2011 Reckon LLP and others. All rights reserved.

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

sub objectType {
    'Notes';
}

sub check {
    my ($self) = @_;
    $self->{lines} = [ SpreadsheetModel::Object::splitLines( $self->{lines} ) ];
    ( $self->{name} ) = splice @{ $self->{lines} }, 0, 2
      if $self->{lines}[0] && !$self->{lines}[1]
      and !$self->{name} || $self->{name} =~ /^Untitled/;
    return;
}

sub lastCol {
    0;
}

sub lastRow {
    $#{ $_[0]{lines} };
}

sub wsWrite {

    my ( $self, $wb, $ws, $row, $col ) = @_;

    return @{ $self->{$wb} }{qw(worksheet row col)} if $self->{$wb};

    ( $row, $col ) = ( ( $ws->{nextFree} ||= -1 ) + 1, 0 )
      unless defined $row && defined $col;

    my @result = @{ $self->{$wb} }{qw(worksheet row col)} = ( $ws, $row, $col );

    my $lastRow = $self->lastRow;

    if ( $self->{name} ) {
        my $n = $self->{name};
        $n = "$wb->{titlePrefix}: $n" if $wb->{titlePrefix};
        if ( $wb->{titleAppend} ) {
            use bytes;
            $n = qq%="$n"&$wb->{titleAppend}%;
        }
        $ws->write( $row, $col, $n, $wb->getFormat('notes') );
        ++$row;
    }

    my $defaultFormat = $wb->getFormat( $self->{defaultFormat} || 'text' );

    for ( 0 .. $lastRow ) {
        $ws->write( $row++, $col, "$self->{lines}[$_]",
              $self->{rowFormats}[$_]
            ? $wb->getFormat( $self->{rowFormats}[$_] )
            : $defaultFormat );
    }

    if ( !$wb->{noLinks} && $self->{sourceLines} ) {
        my $linkFormat = $wb->getFormat('link');
        foreach ( @{ $self->{sourceLines} } ) {
            if ( ref($_) =~ /^SpreadsheetModel::/ ) {
                if ( my $url = $_->wsUrl($wb) ) {
                    $ws->write_url( $row++, $col, $url, "$_->{name}",
                        $linkFormat );
                }
                else {
                    $ws->write_string( $row++, $col, "$_->{name}",
                        $defaultFormat );
                }
            }
            else {
                $ws->write_string( $row++, $col, "$_", $defaultFormat );
            }
        }
    }

    $ws->{nextFree} = $row unless $ws->{nextFree} > $row;

    @result;

}

1;
