package SpreadsheetModel::Chartset;

=head Copyright licence and disclaimer

Copyright 2017 Franck Latrémolière, Reckon LLP and others.

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
    'Chart set';
}

sub check {
    my ($self) = @_;
    return
        "Broken chartset $self->{name} $self->{debug}: charts is "
      . ( ref $self->{charts} )
      . ' but must be ARRAY'
      unless ref $self->{charts} eq 'ARRAY';
    return "Broken chartset $self->{name} $self->{debug}: no rows"
      unless ref $self->{rows};
    $self->{width} ||= 320;
    return;
}

sub wsWrite {
    my ( $self, $wb, $ws, $row, $col ) = @_;
    return $self->{$wb} if $self->{$wb};
    ( $row, $col ) = ( ( $ws->{nextFree} ||= -1 ) + 1, 0 )
      unless defined $row && defined $col;

    if ( $self->{name} ) {
        $ws->write( $row, $col, "$self->{name}", $wb->getFormat('notes') );
        $ws->set_row( $row, 21 );
        ++$row;
    }

    for ( my $c = 0 ; $c < @{ $self->{charts} } ; ++$c ) {
        my $chart = $wb->add_chart(
            %{ $self->{charts}[$c] },
            embedded => 1,
            name     => $self->{charts}[$c]->objectShortName,
        );
        $ws->insert_chart(
            $row + 1, $col + 1, $chart, $self->{width} * $c,
            0,
            $self->{width} / 480.0,
            @{ $self->{rows}{list} } * 20 / 288.0
        );
        $self->{charts}[$c]->applyInstructions( $chart,
            $wb, $ws, $self->{charts}[$c]->{instructions} );
    }

    for ( my $r = 0 ; $r < @{ $self->{rows}{list} } ; ++$r ) {
        $ws->write(
            ++$row,
            $col,
            "$self->{rows}{list}[$r]",
            $wb->getFormat(
                $self->{rowFormats}
                ?  $self->{rowFormats}[$r]
                                  : 'th'
            )
        );
    }

    $row += 2;
    $ws->{nextFree} = $row unless $ws->{nextFree} > $row;
    $self->{$wb} = $ws;
}

1;
