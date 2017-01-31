package SpreadsheetModel::CLI::CommandRunner;

=head Copyright licence and disclaimer

Copyright 2011-2017 Franck Latrémolière and others. All rights reserved.

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
use File::Glob qw(bsd_glob);
use File::Spec::Functions qw(abs2rel rel2abs);

use constant { C_HOMEDIR => 0, };

sub useModels {

    my $self = shift;

    my ( @writerAndParserOptions, $fillSettings, $postProcessor, $executor,
        @files );

    foreach (@_) {
        if (/^-+single/is) {
            $executor = 0;
            next;
        }
        if (/^-+([0-9]*)([tp])?$/is) {
            unless ($executor) {
                if ( $2 ? $2 eq 't' : $^O =~ /win32/i ) {
                    require SpreadsheetModel::CLI::ExecutorThread;
                    $executor = SpreadsheetModel::CLI::ExecutorThread->new;
                }
                else {
                    require SpreadsheetModel::CLI::ExecutorFork;
                    $executor = SpreadsheetModel::CLI::ExecutorFork->new;
                }
            }
            $executor->setThreads($1) if $1;
            next;
        }
        if (/^-+(re-?build.*)/i) {
            require SpreadsheetModel::Data::DataExtraction;
            @writerAndParserOptions =
              SpreadsheetModel::Data::DataExtraction::rebuildWriter( $1,
                $self );
            next;
        }
        if (/^-+(ya?ml.*)/i) {
            require SpreadsheetModel::Data::DataExtraction;
            @writerAndParserOptions =
              SpreadsheetModel::Data::DataExtraction::ymlWriter($1);
            next;
        }
        if (/^-+rules/i) {
            require SpreadsheetModel::Data::DataExtraction;
            @writerAndParserOptions =
              SpreadsheetModel::Data::DataExtraction::rulesWriter();
            next;
        }
        if (/^-+jbz/i) {
            require SpreadsheetModel::Data::DataExtraction;
            @writerAndParserOptions =
              SpreadsheetModel::Data::DataExtraction::jbzWriter();
            next;
        }
        if (/^-+(?:auto|model)check/i) {
            require SpreadsheetModel::Data::Autocheck;
            @writerAndParserOptions =
              SpreadsheetModel::Data::Autocheck->new( $self->[C_HOMEDIR] )
              ->makeWriterAndParserOptions;
            next;
        }
        if (/^-+(json.*)/i) {
            require SpreadsheetModel::Data::DataExtraction;
            @writerAndParserOptions =
              SpreadsheetModel::Data::DataExtraction::jsonWriter($1);
            next;
        }
        if (/^-+sqlite3?(=.*)?$/i) {
            my %settings;
            if ( my $wantedSheet = $1 ) {
                $wantedSheet =~ s/^=//;
                $settings{sheetFilter} = sub { $_[0]{Name} eq $wantedSheet; };
            }
            require SpreadsheetModel::Data::DataExtraction;
            @writerAndParserOptions =
              SpreadsheetModel::Data::DataExtraction::databaseWriter(
                \%settings );
            next;
        }
        if (/^-+prune=(.*)$/i) {
            unless (@writerAndParserOptions) {
                require SpreadsheetModel::Data::DataExtraction;
                @writerAndParserOptions =
                  SpreadsheetModel::Data::DataExtraction::databaseWriter( {} );
            }
            @writerAndParserOptions->( undef, $1 );
            next;
        }
        if (/^-+xls$/i) {
            require SpreadsheetModel::Data::Dumpers;
            @writerAndParserOptions =
              SpreadsheetModel::Data::Dumpers::xlsWriter();
            next;
        }
        if (/^-+flat/i) {
            require SpreadsheetModel::Data::Dumpers;
            @writerAndParserOptions =
              SpreadsheetModel::Data::Dumpers::xlsFlattener();
            next;
        }
        if (/^-+(tsv|txt|csv)$/i) {
            require SpreadsheetModel::Data::Dumpers;
            @writerAndParserOptions =
              SpreadsheetModel::Data::Dumpers::tsvDumper($1);
            next;
        }
        if (/^-+tall(csv)?$/i) {
            require SpreadsheetModel::Data::Dumpers;
            @writerAndParserOptions =
              SpreadsheetModel::Data::Dumpers::tallDumper( $1 || 'xls' );
            next;
        }
        if (/^-+cat$/i) {
            $executor = 0;
            require SpreadsheetModel::Data::Dumpers;
            @writerAndParserOptions =
              SpreadsheetModel::Data::Dumpers::tsvDumper( \*STDOUT );
            next;
        }
        if (/^-+split$/i) {
            require SpreadsheetModel::Data::Dumpers;
            @writerAndParserOptions =
              SpreadsheetModel::Data::Dumpers::xlsSplitter();
            next;
        }
        if (/^-+(calc|convert.*)/i) {
            $fillSettings = $1;
            next;
        }

        push @files, -f $_ ? $_ : grep { -f $_; } bsd_glob($_);

    }

    unless ( defined $executor ) {
        if ( $^O !~ /win32/i
            && eval 'require SpreadsheetModel::CLI::ExecutorFork' )
        {
            $executor = SpreadsheetModel::CLI::ExecutorFork->new;
        }
        elsif ( eval 'require SpreadsheetModel::CLI::ExecutorThread' ) {
            $executor = SpreadsheetModel::CLI::ExecutorThread->new;
        }
        else {
            warn "No multi-threading: $@";
        }
    }

    ( $postProcessor ||=
          makePostProcessor( $fillSettings, @writerAndParserOptions ) )
      ->( $_, $executor )
      foreach @files;

    if ($executor) {
        if ( my $errorCount = $executor->complete ) {
            die(
                (
                    $errorCount > 1
                    ? "$errorCount things have"
                    : 'Something has'
                )
                . ' gone wrong'
            );
        }
    }

}

sub makePostProcessor {

    my ( $processSettings, @writerAndParserOptions, ) = @_;

    my ( $calc_mainprocess, $calc_ownthread, $calc_worker );
    if ( $processSettings && $processSettings =~ /calc|convert/i ) {

        if ( $^O =~ /win32/i ) {

            # Control Microsoft Excel (not Excel Mobile) under Windows.
            # Each calculator runs in its own thread, run synchronously.
            # (Loading Win32::OLE in the mother thread causes a crash.)
            # It would have been better to set up a single worker thread
            # and some queues to handle Win32::OLE calculations.

            if ( $processSettings =~ /calc/ ) {
                $calc_ownthread = sub {
                    my ($inname) = @_;
                    my $inpath = $inname;
                    $inpath =~ s/\.(xls.?)$/-$$.$1/i;
                    rename $inname, $inpath;
                    require Win32::OLE;
                    if ( my $excelApp =
                           Win32::OLE->GetActiveObject('Excel.Application')
                        || Win32::OLE->new( 'Excel.Application', 'Quit' ) )
                    {
                        my $excelWorkbooks;
                        $excelWorkbooks = $excelApp->Workbooks
                          until $excelWorkbooks;
                        my $excelWorkbook;
                        $excelWorkbook = $excelWorkbooks->Open($inpath)
                          until $excelWorkbook;
                        $excelWorkbook->Save;
                        warn 'Waiting for Excel' until $excelWorkbook->Saved;
                        $excelWorkbook->Close;
                        $excelWorkbook->Dispose;
                    }
                    else {
                        warn 'Cannot find Microsoft Excel';
                    }
                    rename $inpath, $inname
                      or die "rename $inpath, $inname: $! in " . `pwd`;
                    $inname;
                };
            }
            else {
                my @convertIncantation = ( FileFormat => 39 );
                my $convertExtension = '.xls';
                if ( $processSettings =~ /xlsx/i ) {
                    @convertIncantation = ();
                    $convertExtension   = '.xlsx';
                }
                $calc_ownthread = sub {
                    my ($inname) = @_;
                    my $inpath   = $inname;
                    my $outpath  = $inpath;

                    $outpath =~ s/\.xls.?$/$convertExtension/i;
                    my $outname = $outpath;
                    s/\.(xls.?)$/-$$.$1/i foreach $inpath, $outpath;
                    rename $inname, $inpath;
                    require Win32::OLE;
                    if ( my $excelApp =
                           Win32::OLE->GetActiveObject('Excel.Application')
                        || Win32::OLE->new( 'Excel.Application', 'Quit' ) )
                    {
                        my $excelWorkbooks;
                        $excelWorkbooks = $excelApp->Workbooks
                          until $excelWorkbooks;
                        my $excelWorkbook;
                        $excelWorkbook = $excelWorkbooks->Open($inpath)
                          until $excelWorkbook;
                        $excelWorkbook->SaveAs(
                            { FileName => $outpath, @convertIncantation } );
                        warn 'Waiting for Excel' until $excelWorkbook->Saved;
                        $excelWorkbook->Close;
                        $excelWorkbook->Dispose;
                    }
                    else {
                        warn 'Cannot find Microsoft Excel';
                    }
                    rename $inpath,  $inname;
                    rename $outpath, $outname
                      or die "rename $outpath, $outname: $! in " . `pwd`;
                    $outname;
                };
            }
        }

        elsif (`which osascript`) {

            # Control Microsoft Excel under Apple macOS.

            if ( $processSettings =~ /calc/ ) {
                $calc_mainprocess = sub {
                    my ($inname) = @_;
                    my $inpath = $inname;
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
                    rename $inpath, $inname
                      or die "rename $inpath, $inname: $! in " . `pwd`;
                    $inname;
                };
            }
            else {
                my $convert          = ' file format Excel98to2004 file format';
                my $convertExtension = '.xls';
                if ( $processSettings =~ /xlsx/i ) {
                    $convert          = '';
                    $convertExtension = '.xlsx';
                }
                $calc_mainprocess = sub {
                    my ($inname) = @_;
                    my $inpath   = $inname;
                    my $outpath  = $inpath;
                    $outpath =~ s/\.xls.?$/$convertExtension/i;
                    my $outname = $outpath;
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
                    rename $outpath, $outname
                      or die "rename $outpath, $outname: $! in " . `pwd`;
                    $outname;
                };
            }
        }

        elsif (`which ssconvert`) {

            # Try to calculate workbooks using ssconvert

            warn 'Using ssconvert';
            $calc_worker = sub {
                my ($inname) = @_;
                my $inpath   = $inname;
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
            warn 'No automatic calculation attempted';
        }

    }

    require Cwd;
    my $wd = Cwd::getcwd();

    sub {
        my ( $inFile, $executor ) = @_;
        my $absFile = rel2abs( $inFile, $wd );
        unless ( -f $absFile ) {
            warn "$absFile not found";
            return;
        }
        $absFile = $calc_mainprocess->($absFile) if $calc_mainprocess;
        $absFile =
          $INC{'threads.pm'}
          ? threads->new( $calc_ownthread, $absFile )->join
          : $calc_ownthread->($absFile)
          if $calc_ownthread;
        if ($executor) {
            $executor->run( __PACKAGE__, 'parseModel', $absFile,
                [ $calc_worker, @writerAndParserOptions ] );
        }
        else {
            __PACKAGE__->parseModel( $absFile, $calc_worker,
                @writerAndParserOptions );
        }
    };

}

sub parseModel {
    my ( undef, $fileToParse, $calc_worker, $writer, %parserOptions ) = @_;
    $fileToParse = $calc_worker->($fileToParse) if $calc_worker;
    my $workbook;
    eval {
        my $parserModule;
        my $formatter = 'NOOP_CLASS';
        if ( $fileToParse =~ /\.xlsx$/is ) {
            require Spreadsheet::ParseXLSX;
            $parserModule = 'Spreadsheet::ParseXLSX';
        }
        else {
            require Spreadsheet::ParseExcel;
            eval
            { # The NOOP_CLASS produces warnings, the Japanese formatter does not
                require Spreadsheet::ParseExcel::FmtJapan;
                $formatter = Spreadsheet::ParseExcel::FmtJapan->new;
            };
            $parserModule = 'Spreadsheet::ParseExcel';
        }
        if ( my $setup = delete $parserOptions{Setup} ) {
            $setup->($fileToParse);
        }
        my $parser = $parserModule->new(%parserOptions);
        $workbook = $parser->parse( $fileToParse, $formatter );
    };
    warn "$@ for $fileToParse" if $@;
    if ($writer) {
        if ($workbook) {
            eval { $writer->( $fileToParse, $workbook ); };
            die "$@ for $fileToParse" if $@;
        }
        else {
            die "Cannot parse $fileToParse in " . `pwd`;
        }
    }
    0;
}

# Do-nothing cell content formatter for Spreadsheet::ParseExcel
package NOOP_CLASS;

our $AUTOLOAD;

sub AUTOLOAD {
    no strict 'refs';
    *{$AUTOLOAD} = sub { };
    return;
}

1;
