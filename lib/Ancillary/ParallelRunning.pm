package Ancillary::ParallelRunning;

=head Copyright licence and disclaimer

Copyright 2009-2015 Reckon LLP and others.

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
our @EXPORT_OK = qw(registerpid waitanypid);

my %names;
my %conts;
my %status;

sub registerpid {
    my ( $pid, $name, $continuation ) = @_;
    $names{$pid} = $name;
    $conts{$pid} = $continuation if $continuation;
    warn "$name started ($pid)\n";
}

sub backgroundrun {
    my ( $module, $method, $firstArg, $otherArgsRef, $continuation ) = @_;
    my $pid = fork;
    return registerpid( $pid, $firstArg, $continuation ) if $pid;
    $0 = "perl: $firstArg" if defined $pid;
    my $status = $module->$method( $firstArg, @$otherArgsRef );
    if ( defined $pid ) {
        exit $status;

        # If you need to clean up spreadsheet generation
        # without calling exit, use something like:
        #   eval { File::Temp::cleanup(); };
        #   require POSIX and POSIX::_exit($status);

    }
    elsif ($continuation) {
        $continuation->();
    }
}

sub waitanypid {
    my ($limit) = @_;
    while ( keys %names > $limit ) {
        my $pid = waitpid -1, 0;    # WNOHANG
        $status{$pid} = $?;
        my $name = delete $names{$pid};
        if ( my $continuation = delete $conts{$pid} ) {
            $continuation->($name);
        }
        else {
            warn "$name complete ($pid)\n";
        }
    }
    grep { $_ } values %status;
}

1;
