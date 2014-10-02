package Ancillary::CommandLine;

=head Copyright licence and disclaimer

Copyright 2011-2014 Franck Latrémolière and others. All rights reserved.

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
use File::Spec::Functions qw(abs2rel catdir);
use File::Basename 'dirname';

use constant {
    PERL5DIR => 0,
    HOMEDIR  => 1,
};

sub factory {
    my ( $class, $perl5dir, $homedir ) = @_;
    bless [ $perl5dir, $homedir ], $class;
}

sub run {
    goto \&make if grep { /\.(?:ya?ml|json|dta)$/si } @_;
    shift;
    goto \&import if grep { /\.xl[a-z]\S+$/si } @_;
    goto \&export if -s '~$database.sqlite';
    warn "pmod.pl: nothing to do\n";
}

sub make {
    my $self = shift;
    require Ancillary::Manufacturing;
    my $maker = Ancillary::Manufacturing->factory(
        validate => [
            $self->[PERL5DIR],
            grep { -d $_ } catdir( $self->[HOMEDIR], 'X_Revisions' )
        ]
    );

    if ( $^O !~ /win32/i ) {
        if ( my $threads = `sysctl -n hw.ncpu 2>/dev/null` || `nproc` ) {
            chomp $threads;
            $maker->{threads}->($threads);
        }
    }

    foreach ( map { decode_utf8 $_} @_ ) {
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
                my $d = catdir( $self->[PERL5DIR], $1 );
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
            elsif (/^-+password=(.+)/is) {
                $maker->{setRule}->( password => $1 );
            }
            elsif (/^-+(no|skip)protect/is) {
                $maker->{setRule}->( protect => 0 );
            }
            elsif (/^-+(right.*)/is) { $maker->{setRule}->( alignment => $1 ); }
            elsif (/^-+single/is) { $maker->{threads}->(1); }
            elsif (/^-+(sqlite.*)/is) {
                my $settings = "convert$1";
                require Compilation::ImportCalcSqlite;
                $maker->{setting}->(
                    PostProcessing =>
                      Compilation::ImportCalcSqlite::makePostProcessor(
                        $maker->{threads}->(),
                        Compilation::ImportCalcSqlite::makeSQLiteWriter(
                            $settings),
                        $settings
                      )
                );
            }
            elsif (/^-+stats/is) {
                $maker->{setRule}
                  ->( summary => 'statistics', illustrative => 1 );
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
}

sub export {

    if ( grep { /extract1076from1001/i } @_ ) {
        require Compilation::Database;
        Compilation->extract1076from1001;
    }

    if ( grep { /chedam/i } @_ ) {
        require Compilation::ChedamMaster;
        Compilation::Chedam->runFromDatabase;
    }

    require Compilation::Database;
    my $db = Compilation->new;

    if ( grep { /\bcsv\b/i } @_ ) {
        require Compilation::ExportCsv;
        $db->csvCreateEdcm( grep { /all/i } @_ );
        exit 0;
    }

    my $workbookModule =
      ( grep { /^-+xls$/i } @_ )
      ? 'SpreadsheetModel::Workbook'
      : 'SpreadsheetModel::WorkbookXLSX';
    eval "require $workbookModule" or die $@;

    my $options = ( grep { /right/i } @_ ) ? { alignment => 'right' } : {};

    foreach
      my $modelsMatching ( map { /\ball\b/i ? '.' : m#^/(.+)/$# ? $1 : (); }
        @_ )
    {
        require Compilation::ExportTabs;
        my $tablesMatching = join '|', map { /^([0-9]+)$/ ? "^$1" : (); } @_;
        $tablesMatching ||= '.';
        local $_ = "Compilation $modelsMatching$tablesMatching";
        s/ *[^ a-zA-Z0-9-^]//g;
        $db->tableCompilations( $workbookModule, $options, $_, $modelsMatching,
            $tablesMatching );
        return;
    }

    if ( grep { /csv/i } @_ ) {
        require Compilation::ExportCsv;
        $options->{tablesMatching} =
          [qw(^11 ^911$ ^913$ ^935$ ^4501$ ^4601$ ^47)]
          unless grep { /all/i } @_;
        $db->csvCompilation( $options, );
        return;
    }

    if ( grep { /\btscs/i } @_ ) {
        require Compilation::ExportTscs;
        my @tablesMatching = map { /^([0-9]+)$/ ? "^$1" : (); } @_;
        @tablesMatching = ('.') unless @tablesMatching;
        $options->{tablesMatching} = \@tablesMatching;
        $db->tscsCreateIntermediateTables($options)
          unless grep { /norebuild/i } @_;
        $db->tscsCompilation( $workbookModule, $options, );
        return;
    }

    if ( my ($dcp) = map { /^(dcp\S*)/i ? $1 : /-dcp=(.+)/i ? $1 : (); } @_ ) {
        my $options = {};
        if ( my ($yml) = grep { /\.ya?ml$/is } @_ ) {
            require YAML;
            $options = YAML::LoadFile($yml);
        }
        $options->{colour} = 'orange' if grep { /^-*orange$/ } @_;
        my ($name) = map { /-+name=(.*)/i ? $1 : (); } @_;
        my ($base) = map { /-+base=(.+)/i ? $1 : (); } @_;
        if ($base) { $name ||= $dcp . ' v ' . $base; }
        else {
            $name ||= $dcp;
            $base = qr/original|clean|after|master|mini|F201|L201|F600|L600/i;
        }
        require Compilation::ExportImpact;
        my @arguments = (
            $workbookModule,
            dcpName   => $name,
            basematch => sub { $_[0] =~ /$base/i; },
            dcpmatch  => sub { $_[0] =~ /[-+]$dcp/i; },
            %$options,
        );
        my @outputs = map { /^-+(cdcm\S*|edcm\S*)/ ? $1 : (); } @_;
        @outputs = qw(cdcm) unless @outputs;
        $db->$_(@arguments) foreach map {
            $_ eq 'cdcm'
              ? qw(cdcmTariffImpact cdcmPpuImpact cdcmRevenueMatrixImpact cdcmUserImpact)
              : $_ eq 'edcm' ? qw(edcmTariffImpact edcmRevenueMatrixImpact)
              :                $_;
        } @outputs;
        return;
    }

}

sub import {

    require Compilation::ImportDumpers;
    require Compilation::ImportCalcSqlite;
    require Ancillary::ParallelRunning;

    my ( $writer, $settings, $postProcessor );

    my $threads;
    $threads = `sysctl -n hw.ncpu 2>/dev/null` || `nproc`
      unless $^O =~ /win32/i;
    chomp $threads if $threads;
    $threads ||= 1;

    foreach (@_) {
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
            my $sheetFilter;
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
}

1;
