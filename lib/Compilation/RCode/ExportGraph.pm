package Compilation::RCode::ExportGraph;

=head Copyright licence and disclaimer

Copyright 2015 Franck Latrémolière, Reckon LLP and others.

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

local undef $/;
binmode DATA, ':utf8';
my $rCode = <DATA>;

sub rCode {
    my ( $self, $script ) = @_;
    return '' if $script->{$self};
    $script->{$self} = 1;
    $rCode;
}

1;

__DATA__
exportGraph <- function ( p, file.name='Graph', file.type='pdf' ) {
    if ( file.type == 'pdf' | file.type == 'PDF' ) {
        filename<-paste(file.name, '.pdf', sep='');
        pdf(filename, width=11.69, height=8.27);
        if (is.null(p$layers)) {
            for (a in 1:length(p)) print(p[[a]]);
        } else {
            print(p);
        }
    } else if (is.null(p$layers)) {
        for (a in 1:length(p)) exportGraph(p[[a]], paste(file.name, a, sep='-'), file.type);
    } else {
         if ( file.type == 'svg' | file.type == 'SVG' ) {
            filename<-paste(file.name, '.svg', sep='');
            svg(filename, width=11, height=7.58);
        } else if ( file.type == 'jpg' | file.type == 'jpeg' | file.type == 'JPG' | file.type == 'JPEG' ) {
            file.type <- 2048;
            filename<-paste(file.name, '.jpeg', sep='');
            jpeg(filename, width=file.type, height=file.type/1.5, res=file.type/10.0);
        } else {
            file.type <- as.integer(file.type);
            filename<-paste(file.name, '.png', sep='');
            png(filename, width=file.type, height=file.type/1.5, res=file.type/10.0);
        }
        print(p);
    }
    graphics.off();
};
