#!/usr/bin/env perl

=head Copyright licence and disclaimer

Copyright 2011-2014 Franck Latrémolière, Reckon LLP and others.

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
use Encode qw(decode_utf8);
binmode STDIN,  ':utf8';
binmode STDOUT, ':utf8';
binmode STDERR, ':utf8';
use File::Spec::Functions qw(rel2abs abs2rel catfile catdir);
use File::Basename 'dirname';
my ( $homedir, $perl5dir );

BEGIN {
    $homedir = dirname( rel2abs( -l $0 ? ( readlink $0, dirname $0) : $0 ) );
    while (1) {
        $perl5dir = catdir( $homedir, 'lib' );
        last if -d catdir( $perl5dir, 'SpreadsheetModel' );
        my $parent = dirname $homedir;
        last if $parent eq $homedir;
        $homedir = $parent;
    }
}
use lib catdir( $homedir, 'cpan' ), $perl5dir;

use Ancillary::Manufacturing;
my $maker =
  Ancillary::Manufacturing->factory(
    validate => [ $perl5dir, grep { -d $_ } catdir( $homedir, 'X_Revisions' ) ]
  );

if ( $^O !~ /win32/i ) {
    if ( my $threads = `sysctl -n hw.ncpu 2>/dev/null` || `nproc` ) {
        chomp $threads;
        $maker->{threads}->($threads);
    }
}

foreach ( map { decode_utf8 $_} @ARGV ) {
    if (/^-/s) {
        if (/^-+$/s) { $maker->{processStream}->( \*STDIN ); }
        elsif (/^-+(?:carp|confess)/is) {
            require Carp;
            $SIG{__DIE__} = \&Carp::confess;
        }
        elsif (/^-+check/is) {
            $maker->{setRule}->(
                activeSheets => 'Result|Tariff',
                checksums    => 'Tariff checksum 5; Model checksum 7'
            );
        }
        elsif (/^-+debug/is)   { $maker->{setRule}->( debug        => 1 ); }
        elsif (/^-+forward/is) { $maker->{setRule}->( forwardLinks => 1 ); }
        elsif (
            /^-+( graphviz|
                  html|
                  perl|
                  rtf|
                  text|
                  yaml
                )/xis
          )
        {
            $maker->{setting}->( 'Export' . ucfirst( lc($1) ), 1 );
        }
        elsif (/^-+lib=(\S+)/is) {
            my $d = catdir( $perl5dir, $1 );
            if ( -d $d ) {
                lib->import($d);
            }
            else {
                die "Special lib $d not found";
            }
        }
        elsif (
            /^-+( numExtraLocations|
                  numExtraTariffs|
                  numLocations|
                  numSampleTariffs|
                  numTariffs
                )=([0-9]+)/xis
          )
        {
            $maker->{setRule}->( $1 => $2 );
        }
        elsif (/^-+orange/is) { $maker->{setRule}->( colour => 'orange' ); }
        elsif (/^-+gold/is) {
            srand();
            $maker->{setRule}->( colour => 'gold', password => rand() );
        }
        elsif (/^-+pickbest/is) {
            $maker->{setting}->( pickBestRules => 1 );
        }
        elsif (/^-+password=(.+)/is) { $maker->{setRule}->( password => $1 ); }
        elsif (/^-+(no|skip)protect/is) { $maker->{setRule}->( protect => 0 ); }
        elsif (/^-+(right.*)/is) { $maker->{setRule}->( alignment => $1 ); }
        elsif (/^-+single/is)    { $maker->{threads}->(1); }
        elsif (/^-+(sqlite.*)/is) {
            my $settings = "convert$1";
            require Compilation::ImportCalcSqlite;
            $maker->{setting}->(
                PostProcessing =>
                  Compilation::ImportCalcSqlite::makePostProcessor(
                    $maker->{threads}->(),
                    Compilation::ImportCalcSqlite::makeSQLiteWriter($settings),
                    $settings
                  )
            );
        }
        elsif (/^-+stats/is) {
            $maker->{setRule}->( summary => 'statistics', illustrative => 1 );
        }
        elsif (/^-+([0-9]+)/is) { $maker->{threads}->($1); }
        elsif (/^-+template(?:=(.+))?/is) {
            $maker->{setRule}->( template => $1 || ( time . "-$$" ) );
        }
        elsif (/^-+xdata=?(.*)/is) {
            if ($1) {
                $maker->{xdata}->($1);
            }
            else {
                local undef $/;
                print "Enter xdata:\n";
                $maker->{xdata}->(<STDIN>);
            }
        }
        elsif (/^-+xls$/is)  { $maker->{setting}->( xls => 1 ); }
        elsif (/^-+xlsx$/is) { $maker->{setting}->( xls => 0 ); }
        elsif (/^-+new(data|rules|settings)/is) {
            $maker->{fileList}->();
            $maker->{ 'reset' . ucfirst( lc($1) ) }->();
        }
        else {
            warn "Unrecognised option: $_";
        }
    }
    elsif ( -f $_ ) {
        my $dh;
        if (/\.(ygz|ybz|bz2|gz)$/si) {
            local $_ = $_;
            s/'/'"'"'/g;
            open $dh, join ' ', ( $1 =~ /bz/ ? 'bzcat' : qw(gunzip -c) ),
              "'$_'", '|';
        }
        else {
            open $dh, '<', $_;
        }
        unless ($dh) {
            warn "Could not open file: $_";
            next;
        }
        $maker->{processStream}->( $dh, abs2rel($_) );
    }
    else {
        warn "Cannot handle this argument: $_";
    }
}

$maker->{fileList}->();
$maker->{run}->();
