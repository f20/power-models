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
my $maker = Ancillary::Manufacturing->factory;
my %dataAccumulator;
my %override;
my $xdata = '';

my $threads;
$threads = `sysctl -n hw.ncpu 2>/dev/null` || `nproc` unless $^O =~ /win32/i;
chomp $threads if $threads;
$threads ||= 1;

my $processArg = sub {
    local $_ = $_[0];
    if (/^-/s) {
        if (/^-+$/s) { $maker->{processStream}->( \*STDIN ); }
        elsif (/^-+(?:carp|confess)/is) {
            require Carp;
            $SIG{__DIE__} = \&Carp::confess;
        }
        elsif (/^-+check/is) {
            $override{activeSheets} = 'Result|Tariff';
            $override{checksums}    = 'Tariff checksum 5; Model checksum 7';
        }
        elsif (/^-+debug/is)   { $override{debug}        = 1; }
        elsif (/^-+forward/is) { $override{forwardLinks} = 1; }
        elsif (/^-+(html|text|rtf|perl|yaml|graphviz)/is) {
            $maker->{addOptions}->( 'Export' . ucfirst( lc($1) ), 1 );
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
            $override{$1} = $2;
        }
        elsif (/^-+orange/is) { $override{colour} = 'orange'; }
        elsif (/^-+gold/is) {
            $override{colour} = 'gold';
            srand();
            $override{password} = rand();
        }
        elsif (/^-+output=?(.*)/is) {
            $maker->{setSettings}->( output => $1 );
        }
        elsif (/^-+pickbest/is) {
            $maker->{setSettings}->( pickBestRules => 1 );
        }
        elsif (/^-+password=(.+)/is)    { $override{password}  = $1; }
        elsif (/^-+(no|skip)protect/is) { $override{protect}   = 0; }
        elsif (/^-+(right.*)/is)        { $override{alignment} = $1; }
        elsif (/^-+single/is)           { $threads             = 1; }
        elsif (/^-+(sqlite.*)/is) {
            my $settings = "convert$1";
            require Compilation::Import;
            $maker->{setSettings}->(
                PostProcessing => Compilation::Import::makePostProcessor(
                    $threads, Compilation::Import::makeSQLiteWriter($settings),
                    $settings
                )
            );
        }
        elsif (/^-+stats/is) {
            $override{summary}      = 'statistics';
            $override{illustrative} = 1;
        }
        elsif (/^-+([0-9]+)/is) { $threads = $1; }
        elsif (/^-+template(?:=(.+))?/is) {
            $override{template} = $1 || ( time . "-$$" );
        }
        elsif (/^-+xdata=?(.*)/is) {
            $xdata .= "$1\n";
            unless ($xdata) {
                local undef $/;
                print "Enter xdata:\n";
                $xdata .= <STDIN>;
            }
        }
        elsif (/^-+xls$/is)  { $maker->{useXLS}->(); }
        elsif (/^-+xlsx$/is) { $maker->{useXLSX}->(); }
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
        elsif (/\.(?:ya?ml|json|dta)$/si) {
            open $dh, '<', $_;
        }
        unless ($dh) {
            warn "Cannot open $_";
            return;
        }
        if (/\.dta$/is) {
            require Parse::Stata::DtaReader;
            warn "Reading $_\n";
            my $dta = Parse::Stata::DtaReader->new($dh);
            my ( @table, @column );
            for ( my $i = 1 ; $i < $dta->{nvar} ; ++$i ) {
                if ( $dta->{varlist}[$i] =~ /t([0-9]+)c([0-9]+)/ ) {
                    $table[$i]  = $1;
                    $column[$i] = $2;
                }
            }
            while ( my @row = $dta->readRow ) {
                my $book = $row[0];
                my $line = $table[1] ? 'Single-line CSV' : $row[1];
                $dataAccumulator{$book}{ $table[$_] }[ $column[$_] ]{$line} =
                  $row[$_]
                  foreach grep { $table[$_] } 1 .. $#table;
            }
            return;
        }
        $maker->{processStream}->( $dh, abs2rel($_) );
    }
    else {
        warn "Cannot handle this argument: $_";
    }
};

$processArg->($_) foreach @ARGV;

if (%dataAccumulator) {
    require YAML;
    while ( my ( $book, $data ) = each %dataAccumulator ) {
        $book =~ s/(?:-LRIC|-LRICsplit|-FCP)?(-r[0-9]+)?$//is;
        my $yml = "$book.yml";
        if ( 0 && -e $yml ) {
            my $no = 0;
            $yml = $book . --$no . '.yml' while -e $yml;
        }
        warn "Writing $book data\n";
        open my $h, '>', $yml;
        binmode $h, ':utf8';
        print {$h} YAML::Dump $data;
        close $h;
        $processArg->($yml);
    }
}

$maker->{overrideRules}->(%override) if %override;
$maker->{overrideData}->($xdata)     if $xdata;
$maker->{setThreads}->($threads);
$maker->{validate}
  ->( $perl5dir, grep { -e $_ } catdir( $homedir, 'X_Revisions' ) );
$maker->{ $threads > 1 ? 'runParallel' : 'run' }
  ->( $maker->{prepare}->( $maker->{fileList}->() ) );
