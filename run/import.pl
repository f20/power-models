#!/usr/bin/env perl

=head Copyright licence and disclaimer

Copyright 2008-2014 Reckon LLP and others.

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
use Carp;
$SIG{__DIE__} = \&Carp::confess;
use File::Spec::Functions qw(rel2abs catdir);
use File::Basename 'dirname';
use Cwd;
my ( $cwd, $homedir );

BEGIN {
    $cwd = getcwd();
    $homedir = dirname( rel2abs( -l $0 ? ( readlink $0, dirname $0) : $0 ) );
    while (1) {
        last if -d catdir( $homedir, 'lib', 'SpreadsheetModel' );
        my $parent = dirname $homedir;
        last if $parent eq $homedir;
        $homedir = $parent;
    }
    chdir $homedir or die "chdir $homedir: $!";
    $homedir = getcwd();    # to resolve any /../ in the path
    chdir $cwd;
}
use lib map { catdir( $homedir, $_ ); } qw(cpan lib);
use Compilation::ImportDumpers;
use Compilation::ImportCalcSqlite;
use Ancillary::ParallelRunning;

my ( $sheetFilter, $writer, $settings, $postProcessor );

my $threads;
$threads = `sysctl -n hw.ncpu 2>/dev/null` || `nproc` unless $^O =~ /win32/i;
chomp $threads if $threads;
$threads ||= 1;

foreach (@ARGV) {
    if (/^-+([0-9]+)$/i) {
        $threads = $1 if $1 > 0;
        next;
    }
    if (/^-+(ya?ml.*)/i) {
        $writer = Compilation::ImportDumpers::ymlWriter($1);
        next;
    }
    if (/^-+(json.*)/i) {
        $writer = Compilation::ImportDumpers::jsonWriter($1);
        next;
    }
    if (/^-+sqlite3?(=.*)?$/i) {
        if ( my $wantedSheet = $1 ) {
            $wantedSheet =~ s/^=//;
            $sheetFilter = sub { $_[0] eq $wantedSheet; };
        }
        $writer =
          Compilation::ImportCalcSqlite::makeSQLiteWriter( undef,
            $sheetFilter );
        next;
    }
    if (/^-+prune=(.*)$/i) {
        $writer->( undef, $1 );
        next;
    }
    if (/^-+xls$/i) {
        $writer = Compilation::ImportDumpers::xlsWriter();
        next;
    }
    if (/^-+flat/i) {
        $writer = Compilation::ImportDumpers::xlsFlattener();
        next;
    }
    if (/^-+(tsv|txt|csv)$/i) {
        $writer = Compilation::ImportDumpers::tsvDumper($1);
        next;
    }
    if (/^-+tall(csv)?$/i) {
        $writer = Compilation::ImportDumpers::tallDumper( $1 || 'xls' );
        next;
    }
    if (/^-+cat$/i) {
        $threads = 1;
        $writer  = Compilation::ImportDumpers::tsvDumper( \*STDOUT );
        next;
    }
    if (/^-+split$/i) {
        $writer = Compilation::ImportDumpers::xlsSplitter();
        next;
    }
    if (/^-+(calc|convert.*)/i) {
        $settings = $1;
        next;
    }

    (
        $postProcessor ||= Compilation::ImportCalcSqlite::makePostProcessor(
            $threads, $writer, $settings
        )
    )->($_);

}

Ancillary::ParallelRunning::waitanypid(0) if $threads > 1;
