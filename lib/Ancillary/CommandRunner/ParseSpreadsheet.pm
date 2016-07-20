package Ancillary::CommandRunner;

=head Copyright licence and disclaimer

Copyright 2011-2015 Franck Latrémolière and others. All rights reserved.

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
        if (/^-+(re-?build.*)/i) {
            require Compilation::DataExtraction;
            $writer = Compilation::DataExtraction::rebuildWriter( $1, $self );
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
        if (/^-+(?:auto|model)check/i) {
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

        ( $postProcessor ||=
              _makePostProcessor( $threads, $writer, $settings ) )->($_);

    }

    if ( $threads > 1 ) {
        my $errorCount = Ancillary::ParallelRunning::waitanypid(0);
        die "Some ($errorCount) things have gone wrong" if $errorCount;
    }

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
            warn "$inFile not found";
            return;
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
                local %SIG;

                #  $SIG{__WARN__} = sub { };
                if ( $calcFile =~ /\.xlsx$/is ) {
                    require Spreadsheet::ParseXLSX;
                    my $parser = Spreadsheet::ParseXLSX->new;
                    $workbook = $parser->parse( $calcFile, 'NOOP_CLASS' );
                }
                else {
                    require Spreadsheet::ParseExcel;
                    my $parser    = Spreadsheet::ParseExcel->new;
                    my $formatter = 'NOOP_CLASS';
                    eval {
                        require Spreadsheet::ParseExcel::FmtJapan;
                        $formatter = Spreadsheet::ParseExcel::FmtJapan->new;
                    };
                    $workbook = $parser->Parse( $calcFile, $formatter );
                }
                delete $SIG{__WARN__};
            };
            warn "$@ for $calcFile" if $@;
            if ($writer) {
                if ($workbook) {
                    eval { $writer->( $calcFile, $workbook ); };
                    warn "$@ for $calcFile" if $@;
                }
                else {
                    warn "Cannot parse $calcFile";
                }
            }
            exit 0 if defined $pid;
        }
    };
}

1;
