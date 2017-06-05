#!/usr/bin/env perl

=head Copyright licence and disclaimer

Copyright 2011-2016 Franck Latrémolière and others. All rights reserved.

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
binmode STDIN,  ':utf8';
binmode STDOUT, ':utf8';
binmode STDERR, ':utf8';
use File::Spec::Functions qw(rel2abs catdir);
use File::Basename qw(dirname);
use Cwd qw(getcwd);

my ( @homes, @validatedLibs, @otherLibs );

BEGIN {
    my @scriptPaths = ( getcwd(), dirname( rel2abs($0) ) );
    push @scriptPaths, dirname( rel2abs( readlink $0, dirname $0) ) if -l $0;
    foreach my $folder (@scriptPaths) {
        while (1) {
            push @homes, $folder if -e catdir( $folder, 'models' );
            my $lib = catdir( $folder, 'lib' );
            if ( -d $lib ) {
                push @validatedLibs, $lib;
                $lib = catdir( $folder, 'cpan' );
                push @otherLibs, $lib if -d $lib;
                last;
            }
            my $parent = dirname $folder;
            last if $parent eq $folder;
            $folder = $parent;
        }
    }
}
use lib @validatedLibs, @otherLibs;

use SpreadsheetModel::CLI::CommandParser;
my $parser = SpreadsheetModel::CLI::CommandParser->new;
@ARGV ? $parser->acceptCommand(@ARGV) : $parser->acceptScript( \*STDIN );

use SpreadsheetModel::CLI::CommandRunner;
my $runner =
  SpreadsheetModel::CLI::CommandRunner->new( \@homes, \@validatedLibs );
$parser->run($runner);
$runner->finish;
