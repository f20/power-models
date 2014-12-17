package Ancillary::CommandRunner;

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
use File::Spec::Functions qw(abs2rel catdir catfile rel2abs);
use File::Basename 'dirname';

use constant {
    C_PERL5DIR => 0,
    C_HOMEDIR  => 1,
    C_FOLDER   => 2,
    C_LOG      => 3,
};

sub factory {
    my ( $class, $perl5dir, $homedir ) = @_;
    bless [ $perl5dir, $homedir ], $class;
}

sub makeFolder {
    my ( $self, $folder ) = @_;
    if ( $self->[C_FOLDER] ) {
        return if $folder && $folder eq $self->[C_FOLDER];
        if ( $self->[C_LOG] ) {
            open my $h, '>', '~$tmptxt' . $$;
            print {$h} @{ $self->[C_LOG] };
            close $h;
            rename '~$tmptxt' . $$, "$self->[C_FOLDER].txt";
            delete $self->[C_LOG];
        }
        chdir '..';
        my $tmp = '~$tmp-' . $$ . ' ' . $self->[C_FOLDER];
        return if rmdir $tmp;
        rename $self->[C_FOLDER], $tmp . '/~$old-' . $$
          if -e $self->[C_FOLDER];
        rename $tmp, $self->[C_FOLDER];
        delete $self->[C_FOLDER];
    }
    if ($folder) {
        my $tmp = '~$tmp-' . $$ . ' ' . ( $self->[C_FOLDER] = $folder );
        mkdir $tmp;
        chdir $tmp;
    }
}

sub finish {
    my ($self) = @_;
    $self->makeFolder;
}

sub log {
    my ( $self, $verb, @objects ) = @_;
    return if $verb eq 'makeFolder';
    push @{ $self->[C_LOG] },
      join( "\n", $verb, map { "\t$_"; } @objects ) . "\n\n";
}

sub R {
    my $self = shift;
    open my $r, '| R --vanilla --slave';
    binmode $r, ':utf8';
    print {$r} (
          /^\s*#\s*include\s*<\s*(.*\S)\s*>/
        ? qq%source("$self->[C_HOMEDIR]/other/R/$1");%
        : $_
      )
      . "\n"
      foreach @_;
    close $r;
}

sub makeModels {

    my $self = shift;

    require Ancillary::Manufacturing;
    my $maker = Ancillary::Manufacturing->factory(
        validate => [
            $self->[C_PERL5DIR],
            grep { -d $_ } catdir( $self->[C_HOMEDIR], 'X_Revisions' )
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
            elsif (/^-+(auto)?check/is) {
                $maker->{setRule}->(
                    activeSheets => 'Result|Tariff',
                    checksums    => 'Tariff checksum 5; Model checksum 7'
                );
                if (/^-+autocheck/is) {
                    require Compilation::DataExtraction;
                    $maker->{setting}->(
                        PostProcessing => _makePostProcessor(
                            $maker->{threads}->(),
                            Compilation::DataExtraction::checksumWriter(),
                            'convert'
                        )
                    );
                }
            }
            elsif (/^-+debug/is)   { $maker->{setRule}->( debug        => 1 ); }
            elsif (/^-+edcm/is)    { $maker->{setRule}->( edcmTables   => 1 ); }
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
                my $d = catdir( $self->[C_PERL5DIR], $1 );
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
                require Compilation::DataExtraction;
                $maker->{setting}->(
                    PostProcessing => _makePostProcessor(
                        $maker->{threads}->(),
                        Compilation::DataExtraction::databaseWriter(),
                        "convert$1"
                    )
                );
            }
            elsif (/^-+stats=?(.*)/is) {
                $maker->{setRule}->(
                    summary      => 'statistics',
                    illustrative => 1,
                    $1 ? ( statistics => $1 ) : (),
                );
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
            $maker->{addFile}->( abs2rel($_) );
        }
        else {
            s/^\s+//s;
            s/\s+$//s;
            if ( -f $_ ) {
                $maker->{addFile}->( abs2rel($_) );
            }
            else {
                my $file = catfile( $self->[C_HOMEDIR], $_ );
                if ( -f $file ) {
                    $maker->{addFile}->( abs2rel($file) );
                }
                elsif ( my @list = <"$file"> ) {
                    $maker->{addFile}->( abs2rel($_) ) foreach @list;
                }
                else {
                    warn "Cannot handle this argument: $_";
                }
            }
        }
    }

    my @files = $maker->{fileList}->();
    mkdir '~$models' if @files > 3 and !-e '~$models';
    $maker->{run}->();

}

sub useDatabase {

    my $self = shift;

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

sub fillDatabase {

    my $self = shift;

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
            require Compilation::DataExtraction;
            $writer = Compilation::DataExtraction::ymlWriter($1);
            next;
        }
        if (/^-+jbz/i) {
            require Compilation::DataExtraction;
            $writer = Compilation::DataExtraction::jbzWriter();
            next;
        }
        if (/^-+modelcheck/i) {
            require Compilation::DataExtraction;
            $writer = Compilation::DataExtraction::checksumWriter();
            next;
        }
        if (/^-+(json.*)/i) {
            require Compilation::DataExtraction;
            $writer = Compilation::DataExtraction::jsonWriter($1);
            next;
        }
        if (/^-+sqlite3?(=.*)?$/i) {
            my %settings;
            if ( my $wantedSheet = $1 ) {
                $wantedSheet =~ s/^=//;
                $settings{sheetFilter} = sub { $_[0]{Name} eq $wantedSheet; };
            }
            require Compilation::DataExtraction;
            $writer = Compilation::DataExtraction::databaseWriter( \%settings );
            next;
        }
        if (/^-+prune=(.*)$/i) {
            $writer->( undef, $1 );
            next;
        }
        if (/^-+xls$/i) {
            require Compilation::Dumpers;
            $writer = Compilation::Dumpers::xlsWriter();
            next;
        }
        if (/^-+flat/i) {
            require Compilation::Dumpers;
            $writer = Compilation::Dumpers::xlsFlattener();
            next;
        }
        if (/^-+(tsv|txt|csv)$/i) {
            require Compilation::Dumpers;
            $writer = Compilation::Dumpers::tsvDumper($1);
            next;
        }
        if (/^-+tall(csv)?$/i) {
            require Compilation::Dumpers;
            $writer = Compilation::Dumpers::tallDumper( $1 || 'xls' );
            next;
        }
        if (/^-+cat$/i) {
            $threads = 1;
            require Compilation::Dumpers;
            $writer = Compilation::Dumpers::tsvDumper( \*STDOUT );
            next;
        }
        if (/^-+split$/i) {
            require Compilation::Dumpers;
            $writer = Compilation::Dumpers::xlsSplitter();
            next;
        }
        if (/^-+(calc|convert.*)/i) {
            $settings = $1;
            next;
        }

        die 'Nothing to do' unless $writer;

        ( $postProcessor ||=
              _makePostProcessor( $threads, $writer, $settings ) )->($_);

    }

    Ancillary::ParallelRunning::waitanypid(0) if $threads > 1;

}

sub _makePostProcessor {

    my ( $threads1, $writer, $settings ) = @_;
    $threads1 = $threads1 && $threads1 > 1 ? $threads1 - 1 : 0;
    require Ancillary::ParallelRunning if $threads1;

    my ( $calculator_prefork, $calculator_postfork );
    if ( $settings && $settings =~ /calc|convert/i ) {
        if (`which osascript`) {
            if ( $settings =~ /calc/ ) {
                $calculator_prefork = sub {
                    my ($inname) = @_;
                    my $inpath = rel2abs($inname);
                    $inpath =~ s/\.(xls.?)$/-$$.$1/i;
                    rename $inname, $inpath;
                    open my $fh, '| osascript';
                    print $fh <<EOS;
tell application "Microsoft Excel"
	set theWorkbook to open workbook workbook file name POSIX file "$inpath"
	set calculate before save to true
	close theWorkbook saving yes
end tell
EOS
                    close $fh;
                    rename $inpath, $inname;
                    $inname;
                };
            }
            else {
                my $convert          = ' file format Excel98to2004 file format';
                my $convertExtension = '.xls';
                if ( $settings =~ /xlsx/i ) {
                    $convert          = '';
                    $convertExtension = '.xlsx';
                }
                $calculator_prefork = sub {
                    my ($inname) = @_;
                    my $inpath   = rel2abs($inname);
                    my $outpath  = $inpath;
                    $outpath =~ s/\.xls.?$/$convertExtension/i;
                    my $outname = abs2rel($outpath);
                    s/\.(xls.?)$/-$$.$1/i foreach $inpath, $outpath;
                    rename $inname, $inpath;
                    open my $fh, '| osascript';
                    print $fh <<EOS;
tell application "Microsoft Excel"
	set theWorkbook to open workbook workbook file name POSIX file "$inpath"
	set calculate before save to true
	set theFile to POSIX file "$outpath" as string
	save workbook as theWorkbook filename theFile$convert
	close active workbook saving no
end tell
EOS
                    close $fh;
                    rename $inpath,  $inname;
                    rename $outpath, $outname;
                    $outname;
                };
            }
        }
        else {
            if (`which ssconvert`) {
                warn 'Using ssconvert';
                $calculator_postfork = sub {
                    my ($inname) = @_;
                    my $inpath   = rel2abs($inname);
                    my $outpath  = $inpath;
                    $outpath =~ s/\.xls.?$/\.xls/i;
                    my $outname = abs2rel($outpath);
                    s/\.(xls.?)$/-$$.$1/i foreach $inpath, $outpath;
                    rename $inname, $inpath;
                    my @b = ( $inpath, $outpath );
                    s/'/'"'"'/g foreach @b;
                    system qq%ssconvert --recalc '$b[0]' '$b[1]' 2>/dev/null%;
                    rename $inpath,  $inname;
                    rename $outpath, $outname;
                    $outname;
                };
            }
            else {
                warn 'No calculator found';
            }
        }
    }

    sub {
        my ($inFile) = @_;
        unless ( -f $inFile ) {
            $inFile = '~$models/' . $inFile;
            return unless -f $inFile;
        }
        my $calcFile = $inFile;
        $calcFile = $calculator_prefork->($inFile) if $calculator_prefork;
        Ancillary::ParallelRunning::waitanypid($threads1) if $threads1;
        my $pid;
        if ( $threads1 && ( $pid = fork ) ) {
            Ancillary::ParallelRunning::registerpid( $pid, $calcFile );
        }
        else {
            $0 = "perl: $calcFile";
            $calcFile = $calculator_postfork->($inFile) if $calculator_postfork;
            my $workbook;
            eval {
                if ( $calcFile =~ /\.xlsx$/is ) {
                    require Spreadsheet::ParseXLSX;
                    my $parser = Spreadsheet::ParseXLSX->new;
                    $workbook = $parser->parse( $calcFile, 'NOOP_CLASS' );
                }
                else {
                    require Spreadsheet::ParseExcel;
                    my $parser = Spreadsheet::ParseExcel->new;
                    local %SIG;
                    $SIG{__WARN__} = sub { };
                    $workbook = $parser->Parse( $calcFile, 'NOOP_CLASS' );
                    delete $SIG{__WARN__};
                }
            };
            warn "$@ for $calcFile" if $@;
            if ($workbook) {
                eval { $writer->( $calcFile, $workbook ); };
                warn "$@ for $calcFile" if $@;
            }
            else {
                warn "Cannot parse $calcFile";
            }
            exit 0 if defined $pid;
        }
    };
}

our $AUTOLOAD;

sub comment { }

sub DESTROY { }

sub AUTOLOAD {
    no strict 'refs';
    warn "$AUTOLOAD not implemented";
    *{$AUTOLOAD} = sub { };
    return;
}

package NOOP_CLASS;
our $AUTOLOAD;

sub AUTOLOAD {
    no strict 'refs';
    *{$AUTOLOAD} = sub { };
    return;
}

1;
