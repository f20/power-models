#!/usr/bin/env perl

=head Copyright licence and disclaimer

Copyright 2012-2013 Franck Latrémolière, Reckon LLP and others.

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

foreach (@ARGV) {
    next unless -f $_;
    open my $f, '<:utf8', $_ or next;
    local undef $/;
    my $source = <$f>;
    close $f;
    my $destination;
    my $messages;
    my $error;

    if (/\.(t|p[lm])$/) {
        $source =~ s/^\N{U+FEFF}//s;
        require Perl::Tidy;
        local $/ = "\n";    # for perltidy
        $error = Perl::Tidy::perltidy(
            source      => \$source,
            errorfile   => \$messages,
            destination => \$destination,
            argv        => [],
        );
    }
    elsif (/\.ya?ml/) {
        require YAML;
        eval { $destination = YAML::Dump( YAML::Load($source) ); };
        $messages = $@;
    }
    else {
        warn "Ignored: $_";
        next;
    }
    if ( $error || $messages ) {
        warn $messages;
    }
    if ( defined $destination && $source ne $destination ) {
        unlink "$_.tdy";
        link $_, "$_.tdy";
        rename "$_.tdy", "$_.bak";
        open my $g, '>:utf8', "$_.tdy";
        print $g /\.pm$/ ? "\N{U+FEFF}" : (), $destination;
        close $g;
        chmod( 07777 & ( stat $_ )[2], "$_.tdy" );
        rename "$_.tdy", $_;
    }

}
