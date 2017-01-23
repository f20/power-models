package SpreadsheetModel::CLI::CommandRunner;

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
use File::Glob qw(bsd_glob);
use File::Spec::Functions qw(abs2rel rel2abs);

use constant { C_HOMEDIR => 0, };

sub fillDatabase {

    my $self = shift;

    my ( $writer, $fillSettings, $postProcessor );

    my $executor;
    $executor = SpreadsheetModel::CLI::ExecutorFork->new
      if eval 'require SpreadsheetModel::CLI::ExecutorFork';

    foreach (@_) {
        if (/^-+([0-9]+)$/i) {
            if   ( $1 > 1 ) { $executor->setThreads($1); }
            else            { $executor = 0; }
            next;
        }
        if (/^-+(re-?build.*)/i) {
            require SpreadsheetModel::Data::DataExtraction;
            $writer =
              SpreadsheetModel::Data::DataExtraction::rebuildWriter( $1,
                $self );
            next;
        }
        if (/^-+(ya?ml.*)/i) {
            require SpreadsheetModel::Data::DataExtraction;
            $writer = SpreadsheetModel::Data::DataExtraction::ymlWriter($1);
            next;
        }
        if (/^-+rules/i) {
            require SpreadsheetModel::Data::DataExtraction;
            $writer = SpreadsheetModel::Data::DataExtraction::rulesWriter();
            next;
        }
        if (/^-+jbz/i) {
            require SpreadsheetModel::Data::DataExtraction;
            $writer = SpreadsheetModel::Data::DataExtraction::jbzWriter();
            next;
        }
        if (/^-+(?:auto|model)check/i) {
            require SpreadsheetModel::Data::Autocheck;
            $writer =
              SpreadsheetModel::Data::Autocheck->new( $self->[C_HOMEDIR] )
              ->checker;
            next;
        }
        if (/^-+(json.*)/i) {
            require SpreadsheetModel::Data::DataExtraction;
            $writer = SpreadsheetModel::Data::DataExtraction::jsonWriter($1);
            next;
        }
        if (/^-+sqlite3?(=.*)?$/i) {
            my %settings;
            if ( my $wantedSheet = $1 ) {
                $wantedSheet =~ s/^=//;
                $settings{sheetFilter} = sub { $_[0]{Name} eq $wantedSheet; };
            }
            require SpreadsheetModel::Data::DataExtraction;
            $writer =
              SpreadsheetModel::Data::DataExtraction::databaseWriter(
                \%settings );
            next;
        }
        if (/^-+prune=(.*)$/i) {
            unless ($writer) {
                require SpreadsheetModel::Data::DataExtraction;
                $writer =
                  SpreadsheetModel::Data::DataExtraction::databaseWriter( {} );
            }
            $writer->( undef, $1 );
            next;
        }
        if (/^-+xls$/i) {
            require SpreadsheetModel::Data::Dumpers;
            $writer = SpreadsheetModel::Data::Dumpers::xlsWriter();
            next;
        }
        if (/^-+flat/i) {
            require SpreadsheetModel::Data::Dumpers;
            $writer = SpreadsheetModel::Data::Dumpers::xlsFlattener();
            next;
        }
        if (/^-+(tsv|txt|csv)$/i) {
            require SpreadsheetModel::Data::Dumpers;
            $writer = SpreadsheetModel::Data::Dumpers::tsvDumper($1);
            next;
        }
        if (/^-+tall(csv)?$/i) {
            require SpreadsheetModel::Data::Dumpers;
            $writer =
              SpreadsheetModel::Data::Dumpers::tallDumper( $1 || 'xls' );
            next;
        }
        if (/^-+cat$/i) {
            $executor = 0;
            require SpreadsheetModel::Data::Dumpers;
            $writer = SpreadsheetModel::Data::Dumpers::tsvDumper( \*STDOUT );
            next;
        }
        if (/^-+split$/i) {
            require SpreadsheetModel::Data::Dumpers;
            $writer = SpreadsheetModel::Data::Dumpers::xlsSplitter();
            next;
        }
        if (/^-+(calc|convert.*)/i) {
            $fillSettings = $1;
            next;
        }

        ( $postProcessor ||= makePostProcessor( $writer, $fillSettings ) )
          ->( $_, $executor )
          foreach -f $_ ? $_ : grep { -f $_; } bsd_glob($_);

    }

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

    my ( $writer, $processSettings ) = @_;

    my ( $calculator_beforefork, $calculator_afterfork );
    if ( $processSettings && $processSettings =~ /calc|convert/i ) {

        if ( $^O =~ /win32/i ) {

            # Quick and dirty code to control Microsoft Excel
            # (but not Excel Mobile) under Microsoft Windows.

            if ( $processSettings =~ /calc/ ) {
                $calculator_afterfork = sub {
                    my ($inname) = @_;
                    my $inpath = rel2abs($inname);
                    $inpath =~ s/\.(xls.?)$/-$$.$1/i;
                    rename $inname, $inpath;
                    require Win32::OLE;
                    if ( my $excelApp =
                           Win32::OLE->GetActiveObject('Excel.Application')
                        || Win32::OLE->new( 'Excel.Application', 'Quit' ) )
                    {
                        my $excelWorkbook = $excelApp->Workbooks->Open($inpath);
                        $excelWorkbook->Save;
                        warn 'Waiting for Excel' until $excelWorkbook->Saved;
                        $excelWorkbook->Close;
                    }
                    else {
                        warn 'Cannot find Microsoft Excel';
                    }
                    rename $inpath, $inname;
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
                $calculator_afterfork = sub {
                    my ($inname) = @_;
                    my $inpath   = rel2abs($inname);
                    my $outpath  = $inpath;

                    $outpath =~ s/\.xls.?$/$convertExtension/i;
                    my $outname = abs2rel($outpath);
                    s/\.(xls.?)$/-$$.$1/i foreach $inpath, $outpath;
                    rename $inname, $inpath;
                    require Win32::OLE;
                    if ( my $excelApp =
                           Win32::OLE->GetActiveObject('Excel.Application')
                        || Win32::OLE->new( 'Excel.Application', 'Quit' ) )
                    {
                        $excelApp->{Visible}       = 0;
                        $excelApp->{DisplayAlerts} = 0;
                        my $excelWorkbook = $excelApp->Workbooks->Open($inpath);
                        $excelWorkbook->SaveAs(
                            { FileName => $outpath, @convertIncantation } );
                        warn 'Waiting for Excel' until $excelWorkbook->Saved;
                        $excelWorkbook->Close;
                    }
                    else {
                        warn 'Cannot find Microsoft Excel';
                    }
                    rename $inpath, $inname;
                    rename $outpath, $outname or die $!;
                    $outname;
                };
            }
        }

        elsif (`which osascript`) {

            # Code to control Microsoft Excel under Apple macOS.
            if ( $processSettings =~ /calc/ ) {
                $calculator_beforefork = sub {
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
                if ( $processSettings =~ /xlsx/i ) {
                    $convert          = '';
                    $convertExtension = '.xlsx';
                }
                $calculator_beforefork = sub {
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

                # Experimental code to calculate workbooks using ssconvert
                warn 'Using ssconvert';
                $calculator_afterfork = sub {
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
        my ( $inFile, $executor ) = @_;
        unless ( -f $inFile ) {
            warn "$inFile not found";
            return;
        }
        my $fileToParse = $inFile;
        $fileToParse = $calculator_beforefork->($inFile)
          if $calculator_beforefork;
        if ($executor) {
            $executor->run( __PACKAGE__, 'parseModel', $fileToParse,
                [ $calculator_afterfork, $writer ] );
        }
        else {
            __PACKAGE__->parseModel( $fileToParse, $calculator_afterfork,
                $writer );
        }
    };

}

sub parseModel {
    my ( undef, $fileToParse, $calculator_afterfork, $writer ) = @_;
    $fileToParse = $calculator_afterfork->($fileToParse)
      if $calculator_afterfork;
    my $workbook;
    eval {
        if ( $fileToParse =~ /\.xlsx$/is ) {
            require Spreadsheet::ParseXLSX;
            my $parser = Spreadsheet::ParseXLSX->new;
            $workbook = $parser->parse( $fileToParse, 'NOOP_CLASS' );
        }
        else {
            require Spreadsheet::ParseExcel;
            my $parser    = Spreadsheet::ParseExcel->new;
            my $formatter = 'NOOP_CLASS';
            eval {
                require Spreadsheet::ParseExcel::FmtJapan;
                $formatter = Spreadsheet::ParseExcel::FmtJapan->new;
            };
            $workbook = $parser->Parse( $fileToParse, $formatter );
        }
    };
    warn "$@ for $fileToParse" if $@;
    if ($writer) {
        if ($workbook) {
            eval { $writer->( $fileToParse, $workbook ); };
            die "$@ for $fileToParse" if $@;
        }
        else {
            die "Cannot parse $fileToParse";
        }
    }
    0;
}

1;
