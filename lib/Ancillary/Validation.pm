package Ancillary::Validation;

=head Copyright licence and disclaimer

Copyright 2009-2013 Reckon LLP and others.

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
use base 'Exporter';
our @EXPORT_OK = qw(sha1File sourceCodeSha1);

sub sha1File {
    my ($file) = @_;
    return 'no file' unless -f $file;
    my $sha1 = eval {
        require Digest::SHA1;
        my $sha1Machine = new Digest::SHA1;
        open my $fh, '<', $file;
        $sha1Machine->addfile($fh)->hexdigest;
    };
    warn $@ if $@;
    $sha1;
}

sub sourceCodeSha1 {
    my ($perl5dir) = @_;
    my $l = length $perl5dir;
    my %hash;
    eval {
        require Digest::SHA1;
        my $sha1Machine = new Digest::SHA1;
        %hash =
          map {
            substr( $INC{$_}, 0, $l ) eq $perl5dir
              ? do {
                open my $fh, '<', $INC{$_};
                ( $_ => $sha1Machine->addfile($fh)->hexdigest );
              }
              : ();
          } keys %INC;
    };
    \%hash;
}

1;
