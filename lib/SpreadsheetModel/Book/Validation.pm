package SpreadsheetModel::Book::Validation;

# Copyright 2009-2017 Franck Latrémolière, Reckon LLP and others.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice,
# this list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
# this list of conditions and the following disclaimer in the documentation
# and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
# EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
# THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

use warnings;
use strict;
use utf8;

sub digestMachine {
    foreach (qw(Digest::SHA Digest::SHA1 Digest::SHA::PurePerl)) {
        eval "require $_";
        my $digestMachine = eval { $_->new; };
        return $digestMachine if $digestMachine;
    }
}

sub sourceCodeDigest {
    my ($validatedLibs) = @_;
    my @libs = map { [ $_, length $_ ]; } @$validatedLibs;
    my %hash;
    eval {
        my $digestMachine = digestMachine();
        while ( my ( $key, $file ) = each %INC ) {
            next if $key =~ m#^SpreadsheetModel/(?:CLI|Data)/#s;
            next unless grep { substr( $file, 0, $_->[1] ) eq $_->[0]; } @libs;
            open my $fh, '<', $file;
            $hash{$key} = $digestMachine->addfile($fh)->hexdigest;
        }
    };
    \%hash;
}

1;
