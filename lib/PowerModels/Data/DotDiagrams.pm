package PowerModels::Data::DotDiagrams;

=head Copyright licence and disclaimer

Copyright 2013 Franck Latrémolière and others. All rights reserved.

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
use IPC::Open2;

sub _runDot {
    my $dot = shift;
    my $pid = open2(
        my $dhout, my $dh,    # 'env', 'GDFONTPATH=/Reckon/usr/fonts',
        '/usr/local/bin/dot', @_
    );
    print $dh $dot;
    close $dh;
    local undef $/;
    my $out = <$dhout>;
    close $dhout;
    waitpid $pid, 0;
    $out;
}

sub writeDotDiagrams {

    my @diagrams = @_;

    mkdir 'O_Diagrams';

    my $html    = '';
    my $single  = '';
    my $counter = 0;

    foreach (@diagrams) {
        my $id = 'dot' . ( ++$counter );
        my $name = $id;
        $name = $2 if s#(graph\s*\[.*?),\s*label\s*=\s*"(.*?)"#$1#;
        my ( $h, $v );

        my $dot2 = _runDot($_);

        # This is not scaled to fit into "size"
        if (undef) {
            ( $v, $h ) = $dot2 =~ /graph\s*\[bb="0,0,([0-9]+),([0-9]+)"\]/;
        }

        if (1) {
            my $svg = _runDot( $dot2, '-Tsvg' );
            ($h) = $svg =~ /height="([0-9.]+)/;
            ($v) = $svg =~ /width="([0-9.]+)/;
            $_ = int( 0.4999 + $_ * 96 / 72 ) foreach $h, $v;
            open my $fh, '>', "O_Diagrams/$id.svg";
            print $fh $svg;
            $single .= $svg;
        }

      # The PNG is buggered by setlinewidth on Mac, but can be used to get size.
        if (undef) {
            my $res = 96;
            my $png = _runDot( $dot2, '-Tpng', "-Gdpi=$res" );
            ( $v, $h ) = unpack( 'x' x 16 . 'NN', $png );    # big-endian
            $_ = int( 0.4999 + $_ * 96 / $res ) foreach $h, $v;
            open my $fh, '>', "$id.png";
            print $fh $png;
        }

        if (1) {
            my $imap = _runDot( $dot2, '-Tcmap' );
            $imap =~ s/>/ \/>/g;
            my $style = $counter > 1 ? ' style="page-break-before:always"' : '';
            $html .= <<EOX ;
<h2$style>$name</h2><p>
<img style="border:0;width:${v}px;height:${h}px" src="$id.svg" alt="[diagram]" usemap="#$id" />
<map id="$id" name="$id">$imap</map>
</p>
EOX
        }

        if (undef) {
            open my $fh, '>', "O_Diagrams/$id-raw-dot.txt";
            print $fh $_;
        }
        if (undef) {
            open my $fh, '>', "O_Diagrams/$id-processed-dot.txt";
            print $fh $dot2;
        }

    }
    if (1) {
        open my $fh, '>', 'O_Diagrams/index.html';
        print $fh '<!DOCTYPE html><html><head><meta charset="UTF-8">',
          '<meta name="viewport" content="width=650,initial-scale=1.0" />',
          '<title>Dot diagrams</title>', '</head><body>', $html,
          '</body></html>';
    }
    if (undef) {
        open my $fh, '>', '~$single.html';
        print $fh $single;
    }

}

1;
