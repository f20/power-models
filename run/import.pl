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

BEGIN {
    my %pids;

    sub registerpid {
        my ( $pid, $name ) = @_;
        $pids{$pid} = $name;
        warn "$name started ($pid)\n";
    }

    sub waitanypid {
        my ($limit) = @_;
        while ( keys %pids > $limit ) {
            my $pid = waitpid -1, 0;    # WNOHANG
            warn "$pids{$pid} complete ($pid)\n";
            delete $pids{$pid};
        }
    }
}

my ( $sheetFilter, $writer, $calculate, $doNotImport );
our ( $WHICH_osascript, $WHICH_ssconvert );

my $threads1;    # one less than the number of import worker threads we want
$threads1 = `sysctl -n hw.ncpu 2>/dev/null` || `nproc` unless $^O =~ /win32/i;
chomp $threads1 if $threads1;
$threads1 ||= 2;
$threads1 -= 2;    # keep one thread free for Excel

foreach (@ARGV) {
    if (/^-+([0-9]+)$/i) {
        $threads1 = $1 - 1 if $1 > 0;
        next;
    }
    if (/^-+(ya?ml.*)/i) {
        $writer = ymlWriter($1);
        next;
    }
    if (/^-+(json.*)/i) {
        $writer = jsonWriter($1);
        next;
    }
    if (/^-+sqlite3?(=.*)?$/i) {
        if ( my $wantedSheet = $1 ) {
            $wantedSheet =~ s/^=//;
            $sheetFilter = sub { $_[0] eq $wantedSheet; };
        }
        require Compilation::Import;
        $writer = Compilation::Import::makeSQLiteWriter( undef, $sheetFilter );
        next;
    }
    if (/^-+prune=(.*)$/i) {
        $writer->( undef, $1 );
        next;
    }
    if (/^-+xls$/i) {
        $writer = xlsWriter();
        next;
    }
    if (/^-+flat/i) {
        $writer = xlsFlattener();
        next;
    }
    if (/^-+(tsv|txt|csv)$/i) {
        $writer = tsvDumper($1);
        next;
    }
    if (/^-+cat$/i) {
        $threads1 = 0;
        $writer   = tsvDumper( \*STDOUT );
        next;
    }
    if (/^-+split$/i) {
        $writer = xlsSplitter();
        next;
    }
    if (/^-+calconly/i) {
        warn "Calculating only";
        $calculate   = 1;
        $doNotImport = 1;
        next;
    }
    if (/^-+calc/i) {
        warn "Calculating before";
        $calculate = 1;
        next;
    }
    if (/^-+convertxlsx/i) {
        warn "Converting before";
        $calculate = 3;
        next;
    }
    if (/^-+convert/i) {
        warn "Converting before";
        $calculate = 2;
        next;
    }
    my $infile = $_;
    unless ( -f $infile ) {
        warn "No such file: $infile";
        next;
    }
    if ($calculate) {
        warn "Calculating $infile";
        my $originalLocation = $infile;
        my $originalFile     = rel2abs($infile);
        $originalFile =~ s/\.(xls.?)$/-$$.$1/i;
        rename $infile, $originalFile;
        if (
            defined $WHICH_osascript
            ? $WHICH_osascript
            : ( $WHICH_osascript = `which osascript` )
          )
        {
            my $savingCommands = 'close theWorkbook saving yes';
            if ( $calculate > 1 ) {
                $infile =~ s/\.xls.?$/\.xls/i;
                $infile .= 'x' unless $calculate == 2;
                my $calcFile = rel2abs($infile);
                $savingCommands = $calculate == 2
                  ? <<EOS
	set theFile to POSIX file "$calcFile" as string
	save workbook as theWorkbook filename theFile file format Excel98to2004 file format
	close active workbook saving no
EOS
                  : <<EOS;
	set theFile to POSIX file "$calcFile" as string
	save workbook as theWorkbook filename theFile
	close active workbook saving no
EOS
            }
            open my $fh, '| osascript';
            print $fh <<EOS;
tell application "Microsoft Excel"
	set theWorkbook to open workbook workbook file name POSIX file "$originalFile"
	set calculate before save to true
	$savingCommands
end tell
EOS
            close $fh;
            rename $originalFile, $originalLocation;
        }
        elsif (
            defined $WHICH_ssconvert
            ? $WHICH_ssconvert
            : ( $WHICH_ssconvert = `which ssconvert` )
          )
        {
            my $b1 = $infile;
            my $b2 = $originalFile;
            $b1 =~ s/'/'"'"'/g;
            $b2 =~ s/'/'"'"'/g;
            system qq%ssconvert --recalc '$b2' '$b1' 2>/dev/null%;
        }
        else {
            die 'I do not know how to calculate';
        }
    }
    next if $doNotImport;
    waitanypid($threads1);
    my $pid = fork;
    if ($pid) {
        registerpid( $pid, $_ );
        next;
    }
    my $workbook;
    eval {
        if ( $infile =~ /\.xlsx$/is ) {
            require Spreadsheet::ParseXLSX;
            my $parser = Spreadsheet::ParseXLSX->new;
            $workbook = $parser->parse( $infile, 'NOOP2_CLASS' );
        }
        else {
            require Spreadsheet::ParseExcel;
            my $parser = Spreadsheet::ParseExcel->new;
            local %SIG;
            $SIG{__WARN__} = sub { };
            $workbook = $parser->Parse( $infile, 'NOOP2_CLASS' );
            delete $SIG{__WARN__};
        }
    };
    warn "$@ for $infile" if $@;
    if ($workbook) {
        if ($writer) {
            eval { $writer->( $infile, $workbook ); };
            warn "$@ for $infile" if $@;
        }
        else {
            warn "No output specified for $infile";
        }
    }
    else {
        warn "Cannot parse $infile";
    }
    exit 0 if defined $pid;
}

waitanypid(0);

sub updateTree {
    my ( $workbook, $tree, $options ) = @_;
    $tree ||= {};
    my $sheetNumber = 0;
    for my $worksheet ( $workbook->worksheets() ) {
        next if $sheetFilter && !$sheetFilter->( $worksheet->{Name} );
        my ( $row_min, $row_max ) = $worksheet->row_range();
        my ( $col_min, $col_max ) = $worksheet->col_range();
        my $tableNumber = --$sheetNumber;
        my $columnHeadingsRow;
        my $to;
        for my $row ( $row_min .. $row_max ) {
            my $rowName;
            for my $col ( $col_min .. $col_max ) {
                my $cell = $worksheet->get_cell( $row, $col );
                my $v;
                $v = $cell->unformatted if $cell;
                next unless defined $v;
                if ( $col == 0 ) {
                    if ( !ref $cell->{Format} || $cell->{Format}{Lock} ) {
                        if ( $v && $v =~ /^([0-9]{2,})\. / ) {
                            $tableNumber = $1;
                            undef $columnHeadingsRow;
                            $to = $tree->{$tableNumber}
                              || [
                                $tableNumber !~ /00$/
                                  && ( $options->{preferArrays}
                                    || $tableNumber =~ /^(?:17)/ )
                                ? []
                                : {}
                              ];
                            if ( ref $to->[0] eq 'ARRAY' ) {
                                $to->[0][0] = $v;
                            }
                            else {
                                $to->[0]{'_table'} = $v
                                  unless $options->{minimum};
                            }
                        }
                        else {
                            if ($v) {
                                $v =~ s/[^A-Za-z0-9-]/ /g;
                                $v =~ s/- / /g;
                                $v =~ s/ +/ /g;
                                $v =~ s/^ //;
                                $v =~ s/ $//;
                                $rowName = $v eq '' ? '•' : $v;
                                $to->[0][ $row - $columnHeadingsRow ] =
                                  $rowName
                                  if ref $to->[0] eq 'ARRAY'
                                  and defined $columnHeadingsRow;
                            }
                        }
                    }
                    else {
                        if ( ref $to->[0] eq 'HASH' )
                        {    # old-style table comment
                            $to->[0]{'_note'} = $v if $v;
                        }
                        else {
                            $columnHeadingsRow = $row - 1
                              unless defined $columnHeadingsRow;
                            $to->[0][ $row - $columnHeadingsRow ] = $v;
                        }
                    }
                }
                elsif ( !$rowName ) {
                    $columnHeadingsRow = $row;
                    if ( ref $to->[0] eq 'ARRAY' ) {
                        $to->[$col] ||= [$v];
                    }
                    else {
                        $to->[$col]{'_column'} = $v unless $options->{minimum};
                    }
                }
                elsif (ref $cell->{Format}
                    && !$cell->{Format}{Lock}
                    && ( $v || $to->[$col] ) )
                {
                    if ( ref $to->[$col] eq 'ARRAY' ) {
                        $to->[$col][ $row - $columnHeadingsRow ] = $v;
                    }
                    else {
                        $to->[$col]{$rowName} = $v;
                    }
                    $tree->{$tableNumber} = $to
                      if $tableNumber > 0 && !$tree->{$tableNumber};
                }
            }
        }
    }
    $tree;
}

sub ymlWriter {
    my ($arg) = @_;
    my $options = {
        $arg =~ /array/i ? ( preferArrays => 1 ) : (),
        $arg =~ /min/i   ? ( minimum      => 1 ) : (),
    };
    require YAML;
    sub {
        my ( $book, $workbook ) = @_;
        die unless $book;
        my $yml = "$book.yml";
        $yml =~ s/\.xlsx?\.yml$/.yml/is;
        my $tree;
        if ( -e $yml ) {
            open my $h, '<', $yml;
            binmode $h, ':utf8';
            local undef $/;
            $tree = YAML::Load(<$h>);
        }
        open my $h, '>', $yml;
        binmode $h, ':utf8';
        print $h YAML::Dump( updateTree( $workbook, $tree, $options ) );
    };
}

sub jsonWriter {
    my ($arg) = @_;
    my $options = {
        $arg =~ /array/i ? ( preferArrays => 1 ) : (),
        $arg =~ /min/i   ? ( minimum      => 1 ) : (),
        $arg =~ /all/i   ? ( dumpAllData  => 1 ) : (),
    };
    my $jsonpp = !eval 'require JSON';
    require JSON::PP if $jsonpp;
    sub {
        my ( $book, $workbook ) = @_;
        die unless $book;
        my $json = "$book.json";
        $json =~ s/\.xlsx?\.json$/.json/is;
        my $tree;
        if ( -e $json ) {
            open my $h, '<', $json;
            binmode $h;
            local undef $/;
            $tree =
              $jsonpp ? JSON::PP::decode_json(<$h>) : JSON::decode_json(<$h>);
        }
        open my $h, '>', $json;
        binmode $h;
        print {$h}
          ( $jsonpp ? 'JSON::PP' : 'JSON' )->new->canonical(1)
          ->pretty->utf8->encode( updateTree( $workbook, $tree, $options ) );
    };
}

sub xlsWriter {
    require Spreadsheet::WriteExcel;
    sub {
        my ( $infile, $workbook ) = @_;
        die unless $infile;
        my $outfile = "$infile cleaned.xls";
        $outfile =~ s/\.xlsx? cleaned.xls$/ cleaned.xls/is;
        if ( -e $outfile ) {
            warn "$infile skipped";
            next;
        }
        my $outputBook = new Spreadsheet::WriteExcel($outfile);
        for my $worksheet ( $workbook->worksheets() ) {
            next if $sheetFilter && !$sheetFilter->( $worksheet->{Name} );
            my $outputSheet = $outputBook->add_worksheet( $worksheet->{Name} );
            my ( $row_min, $row_max ) = $worksheet->row_range();
            my ( $col_min, $col_max ) = $worksheet->col_range();
            for my $row ( $row_min .. $row_max ) {
                for my $col ( $col_min .. $col_max ) {
                    my $cell = $worksheet->get_cell( $row, $col );
                    next unless $cell;
                    next unless my $v = $cell->unformatted;
                    $v =~ /=/
                      ? $outputSheet->write_string( $row, $col, $v )
                      : $outputSheet->write( $row, $col, $v );
                }
            }
        }
    };
}

sub xlsFlattener {
    require Spreadsheet::WriteExcel;
    sub {
        my ( $infile, $workbook ) = @_;
        die unless $infile;
        my $outfile = "$infile flattened.xls";
        $outfile =~ s/\.xlsx? flattened.xls$/ flattened.xls/is;
        if ( -e $outfile ) {
            warn "$infile skipped";
            next;
        }
        my $outputBook  = new Spreadsheet::WriteExcel($outfile);
        my $outputSheet = $outputBook->add_worksheet('Flattened');
        my $outputRow   = -1;
        for my $worksheet ( $workbook->worksheets() ) {
            next if $sheetFilter && !$sheetFilter->( $worksheet->{Name} );
            my ( $row_min, $row_max ) = $worksheet->row_range();
            my ( $col_min, $col_max ) = $worksheet->col_range();
            for my $row ( $row_min .. $row_max ) {
                $outputSheet->write_string( ++$outputRow, $col_min,
                    $worksheet->{Name} );
                for my $col ( $col_min .. $col_max ) {
                    my $cell = $worksheet->get_cell( $row, $col );
                    next unless $cell;
                    next unless my $v = $cell->unformatted;
                    $v =~ /=/
                      ? $outputSheet->write_string( $outputRow, $col + 1, $v )
                      : $outputSheet->write( $outputRow, $col + 1, $v );
                }
            }
        }
    };
}

sub xlsSplitter {
    require Spreadsheet::WriteExcel;
    sub {
        my ( $infile, $workbook ) = @_;
        die unless $infile;
        for my $worksheet ( $workbook->worksheets() ) {
            my $outfile = "$infile $worksheet->{Name}.xls";
            if ( -e $outfile ) {
                warn "$infile skipped";
                next;
            }
            my $outputBook = new Spreadsheet::WriteExcel($outfile);
            next if $sheetFilter && !$sheetFilter->( $worksheet->{Name} );
            my $outputSheet = $outputBook->add_worksheet( $worksheet->{Name} );
            my ( $row_min, $row_max ) = $worksheet->row_range();
            my ( $col_min, $col_max ) = $worksheet->col_range();
            for my $row ( $row_min .. $row_max ) {
                for my $col ( $col_min .. $col_max ) {
                    my $cell = $worksheet->get_cell( $row, $col );
                    next unless $cell;
                    next unless my $v = $cell->unformatted;
                    $v =~ /=/
                      ? $outputSheet->write_string( $row, $col, $v )
                      : $outputSheet->write( $row, $col, $v );
                }
            }
        }
    };
}

sub tsvDumper {
    my ($output) = @_;
    my $joinChar = $output eq 'csv' ? ',' : "\t";
    sub {
        my ( $infile, $workbook ) = @_;
        my $fh;
        if ( ref $output ) {
            $fh = $output;
        }
        else {
            open $fh, '>', "$infile.$output";
        }
        binmode $fh, ':utf8';
        0 and print {$fh} join( $joinChar, qw(Book Sheet Row A B C) ) . "\n";
        for my $worksheet ( $workbook->worksheets() ) {
            next if $sheetFilter && !$sheetFilter->( $worksheet->{Name} );
            my ( $row_min, $row_max ) = $worksheet->row_range();
            my ( $col_min, $col_max ) = $worksheet->col_range();
            $col_min = 0;    # Use completely blank columns
            print {$fh} join(
                $joinChar,
                $infile,
                $worksheet->{Name},
                0,
                map {
                    my $aa = int( $_ / 26 );
                    ( $aa ? chr( 64 + $aa ) : '' ) . chr( 65 + ( $_ % 26 ) );
                } $col_min .. $col_max
            ) . "\n";
            for my $row ( $row_min .. $row_max ) {
                print {$fh} join(
                    $joinChar,
                    $infile,
                    $worksheet->{Name},
                    1 + $row,
                    map {
                        my $cell = $worksheet->get_cell( $row, $_ );
                        local $_ = !$cell ? '' : $cell->unformatted;
                        s/\n/\\n/gs;
                        s/\r/\\r/gs;
                        $_;
                    } $col_min .. $col_max
                ) . "\n";
            }
        }
    };
}

package NOOP2_CLASS;
our $AUTOLOAD;

sub AUTOLOAD {
    no strict 'refs';
    *{$AUTOLOAD} = sub { };
    return;
}

