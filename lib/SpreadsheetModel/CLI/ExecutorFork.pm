package SpreadsheetModel::CLI::ExecutorFork;

=head Copyright licence and disclaimer

Copyright 2009-2016 Reckon LLP and others.

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

# Singleton
my $maxThreadsMinusOne;
my %processNameByPid;
my %continuationByPid;
my @nameStatusPairs;

sub new {
    my ( $class, $threads ) = @_;
    unless ($threads) {
        if (    $^O !~ /win32/i
            and $threads = `sysctl -n hw.ncpu 2>/dev/null`
            || `nproc 2>/dev/null` )
        {
            chomp $threads;
        }
    }
    return 0 unless $threads && $threads > 1;
    $maxThreadsMinusOne = $threads - 1;
    $class;
}

sub setThreads {
    $maxThreadsMinusOne = $_[1] - 1;
}

sub complete {
    my ( $executor, $maxThreadsLeft ) = @_;
    $maxThreadsLeft ||= 0;
    while ( keys %processNameByPid > $maxThreadsLeft ) {
        my $pid = waitpid -1, 0;
        my $name = delete $processNameByPid{$pid};
        push @nameStatusPairs, [ $name, $? ];
        if ( my $continuation = delete $continuationByPid{$pid} ) {
            $continuation->( $name, $executor );
        }
        else {
            warn "finished $name ($pid)\n";
        }
    }
    return grep { $_->[1]; } @nameStatusPairs unless $maxThreadsLeft;
}

sub run {
    my ( $executor, $module, $method, $firstArg, $otherArgsRef, $continuation )
      = @_;
    $executor->complete($maxThreadsMinusOne);
    local $_ = $firstArg;
    $_ = $1 if m#[/\\].*[/\\]([^/\\]+)#s;
    my $pid = fork;
    if ($pid) {
        $processNameByPid{$pid} = $_;
        $continuationByPid{$pid} = $continuation if $continuation;
        warn "$method $_ ($pid)\n";
    }
    else {
        $0 = "perl: $method $_" if defined $pid;
        my $status = $module->$method( $firstArg, @$otherArgsRef );
        if ( defined $pid ) {
            exit $status;
        }
        elsif ($continuation) {
            $continuation->($firstArg);
        }
    }
}

1;
