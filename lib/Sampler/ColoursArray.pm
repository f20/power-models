package Sampler;

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
use Data::Dumper;
use SpreadsheetModel::Shortcuts ':all';

sub hexrgb {
    '#' . unpack 'H*', pack 'C3', map { int( 0.5 + 255 * $_ ); } @{ $_[0] };
}

sub writeColourMatrix {

    my ( $model, $wbook, $wsheet ) = @_;

    Notes( name => 'Colour matrix, full value, 20 per cent saturation pastels' )
      ->wsWrite( $wbook, $wsheet );
    my $row = $wsheet->{nextFree} || -1;
    ++$row;
    my @saturated = (
        [ 1, 0, 0 ],
        [ 1, 1, 0 ],
        [ 0, 1, 0 ],
        [ 0, 1, 1 ],
        [ 0, 0, 1 ],
        [ 1, 0, 1 ]
    );
    my @pastel = map {
        [ map { 0.8 + 0.2 * $_; } @$_ ];
    } @saturated;
    my @hex = map { hexrgb($_) } @pastel, @saturated;

    for ( my $x = 0 ; $x < @pastel ; ++$x ) {
        $wsheet->write_string(
            $row,
            $x + 1,
            $hex[ 6 + $x ],
            $wbook->getFormat(
                [
                    base       => 'thc',
                    num_format => '@',
                    color      => $hex[ 6 + ( 3 + $x ) % 6 ],
                    bg_color   => $hex[ 6 + $x ],
                ]
            )
        );
    }
    for ( my $y = 0 ; $y < @pastel ; ++$y ) {
        $wsheet->write_string(
            ++$row,
            0,
            $hex[$y],
            $wbook->getFormat(
                [
                    base     => 'puretextcon',
                    bg_color => $hex[$y],
                ]
            )
        );
        for ( my $x = 0 ; $x < @pastel ; ++$x ) {
            my $col = hexrgb(
                [ map { 0.5 * ( $pastel[$x][$_] + $pastel[$y][$_] ); } 0 .. 2 ]
            );
            $wsheet->write_string(
                $row,
                $x + 1,
                $col,
                $wbook->getFormat(
                    [
                        base     => 'puretextcon',
                        bg_color => $col,
                    ]
                )
            );
        }
    }
    $wsheet->{nextFree} = $row + 1;

}

1;
