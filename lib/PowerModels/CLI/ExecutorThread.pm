﻿package PowerModels::CLI::ExecutorThread;

=head Copyright licence and disclaimer

Copyright 2009-2017 Reckon LLP and others.

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
use threads;
use threads::shared;
use Thread::Queue;

my $chdir_lock = 0;
share $chdir_lock;

sub new {
    my ( $class, $threads ) = @_;
    unless ($threads) {
        if ( $^O =~ /win32/i ) {
            $threads = $ENV{'NUMBER_OF_PROCESSORS'}
              if $ENV{'NUMBER_OF_PROCESSORS'};
        }
        elsif ( $threads = `sysctl -n hw.ncpu 2>/dev/null`
            || `nproc 2>/dev/null` )
        {
            chomp $threads;
        }
    }
    return 0 unless $threads && $threads > 1;
    bless {
        maxThreadsMinusOne => $threads - 1,
        completedThreads   => Thread::Queue->new,
        lastTid            => -1,
    }, $class;
}

sub setThreads {
    my ( $executor, $maxThreads ) = @_;
    $executor->{maxThreadsMinusOne} = $maxThreads - 1;
}

sub complete {
    my ( $executor, $maxThreadsLeft ) = @_;
    $maxThreadsLeft ||= 0;
    while (keys %{ $executor->{nameByTid} } > $maxThreadsLeft
        || keys %{ $executor->{threadByTid} } > 5 * $maxThreadsLeft )
    {
        my $dequeued = $executor->{completedThreads}->dequeue;
        if ( ref $dequeued ) {
            my ( $tid, $status ) = @$dequeued;
            my $name = delete $executor->{nameByTid}{$tid};
            push @{ $executor->{nameStatusPairs} }, [ $name, $status ];
            if ( my $continuation =
                delete $executor->{continuationByTid}{$tid} )
            {
                $continuation->( $name, $executor );
            }
            else {
                warn "finished $name ($tid)\n";
            }
        }
        else {
            ( delete $executor->{threadByTid}{$dequeued} )->join;
        }
    }
    return grep { $_->[1]; } @{ $executor->{nameStatusPairs} }
      unless $maxThreadsLeft;
}

sub thread_run {
    my ( $completionQueue, $myid, $module, $method, $firstArg, $otherArgsRef,
        $protectionFlag )
      = @_;
    my ( $status, @hazardousWaste );
    if ($protectionFlag) {
        lock $chdir_lock;
        ++$chdir_lock;
    }
    eval {
        ( $status, @hazardousWaste ) =
          $module->$method( $firstArg, @$otherArgsRef );
    };
    if ($protectionFlag) {
        lock $chdir_lock;
        --$chdir_lock;
        cond_signal $chdir_lock;
    }
    if ($@) {
        warn "$@ in thread $myid";
        $status = $@;
    }
    $completionQueue->enqueue( [ $myid, $status ] );
    if (@hazardousWaste) {
        lock $chdir_lock;
        cond_wait $chdir_lock while $chdir_lock > 0;
        $_->DESTROY foreach @hazardousWaste;
        cond_signal $chdir_lock;
    }
    $completionQueue->enqueue($myid);
}

sub run {
    my ( $executor, $module, $method, $firstArg, $otherArgsRef, $continuation,
        $protectionFlag )
      = @_;
    $executor->complete( $executor->{maxThreadsMinusOne} );
    my $tid = ++$executor->{lastTid};
    local $_ = $firstArg;
    $_ = $1 if m#[/\\].*[/\\]([^/\\]+)#s;
    warn "$method $_ ($tid)\n";
    $executor->{nameByTid}{$tid} = $_;
    $executor->{continuationByTid}{$tid} = $continuation if $continuation;
    $executor->{threadByTid}{$tid} =
      threads->new( \&thread_run, $executor->{completedThreads},
        $tid, $module, $method, $firstArg, $otherArgsRef, $protectionFlag );
}

1;