#!/usr/bin/env perl

=head Copyright licence and disclaimer

Copyright 2011-2013 Franck Latrémolière, Reckon LLP and others.

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
require Carp;
$SIG{__DIE__} = \&Carp::confess;
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
my $maker = Ancillary::Manufacturing->factory;
my %override;
my $xdata = '';
my $threads;
$threads = `sysctl -n hw.ncpu 2>/dev/null` || `nproc` unless $^O =~ /win32/i;
chomp $threads if $threads;
$threads ||= 1;

foreach (@ARGV) {

    if (/^-/s) {
        if    (/^-+$/s)          { $maker->{processStream}->( \*STDIN ); }
        elsif (/^-+xlsx/is)      { $maker->{useXLSX}->(); }
        elsif (/^-+(right.*)/is) { $override{alignment} = $1; }
        elsif (/^-+(no|skip)protect/is) { $override{protect} = 0; }
        elsif (/^-+(html|text|perl|yaml)/is) {
            $override{ 'Export' . ucfirst( lc($1) ) } = 1;
        }
        elsif (/^-+defaultcol/is) { $override{defaultColours} = 1; }
        elsif (/^-+gold/is) {
            $override{colour}   = 'gold';
            $override{password} = rand();
        }
        elsif (/^-+orange/is)        { $override{colour}   = 'orange'; }
        elsif (/^-+debug/is)         { $override{debug}    = 1; }
        elsif (/^-+password=(.+)/is) { $override{password} = $1; }
        elsif (/^-+single/is)        { $threads            = 1; }
        elsif (/^-+([0-9]+)/is)      { $threads            = $1; }
        elsif (/^-+onefile/is)       { $override{template} = time . "-$$"; }
        elsif (
/^-+(numExtraLocations|numExtraTariffs|numLocations|numSampleTariffs|numTariffs)=([0-9]+)/is
          )
        {
            $override{$1} = $2;
        }
        elsif (/^-+xdata=?(.*)/is) {
            $xdata .= "$1\n";
            unless ($xdata) {
                local undef $/;
                print "Enter xdata:\n";
                $xdata .= <STDIN>;
            }
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
        elsif (/\.(?:ya?ml|json)$/si) {
            open $dh, '<', $_;
        }
        $maker->{processStream}->( $dh, abs2rel($_) ) if $dh;
    }
    else {
        warn "Cannot handle this argument: $_";
    }
}
$maker->{overrideRules}->(%override) if %override;
$maker->{overrideData}->($xdata)     if $xdata;
$maker->{setThreads}->($threads);
$maker->{validate}
  ->( $perl5dir, grep { -e $_ } catdir( $homedir, 'X_Revisions' ) );
$maker->{ $threads > 1 ? 'runParallel' : 'run' }
  ->( $maker->{prepare}->( $maker->{fileList}->() ) );
