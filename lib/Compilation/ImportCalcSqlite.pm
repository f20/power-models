package Compilation::ImportCalcSqlite;

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
use File::Spec::Functions qw(rel2abs abs2rel);
use Ancillary::ParallelRunning;

sub makePostProcessor {
    my ( $threads1, $writer, $settings ) = @_;
    $threads1 = $threads1 && $threads1 > 1 ? $threads1 - 1 : 0;
    require Ancillary::ParallelRunning if $threads1;
    my ( $calculator1, $calculator2 );
    ( $calculator1, $calculator2 ) = makeCalculators($settings)
      if $settings && $settings =~ /calc|convert/i;
    sub {
        my ($inFile) = @_;
        unless ( -f $inFile ) {
            $inFile = '~$models/' . $inFile;
            return unless -f $inFile;
        }
        my $calcFile = $inFile;
        $calcFile = $calculator1->($inFile) if $calculator1;
        Ancillary::ParallelRunning::waitanypid($threads1) if $threads1;
        my $pid;
        if ( $threads1 && ( $pid = fork ) ) {
            Ancillary::ParallelRunning::registerpid( $pid, $calcFile );
        }
        else {
            $0 = "perl: $calcFile";
            $calcFile = $calculator2->($inFile) if $calculator2;
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
            if ( $workbook && $writer ) {
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

sub makeCalculators {

    my ($convert) = @_;

    unless (`which osascript`) {
        die 'No spreadsheet calculator found' unless `which ssconvert`;
        return undef, sub {
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

    return sub {
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
      }
      if $convert && $convert =~ /calc/;

    $convert = $convert
      && $convert =~ /xlsx/i ? '' : ' file format Excel98to2004 file format';

    sub {
        my ($inname) = @_;
        my $inpath   = rel2abs($inname);
        my $outpath  = $inpath;
        $outpath =~ s/\.xls.?$/\.xls/i;
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

sub makeSQLiteWriter {

    my ( $settings, $sheetFilter ) = @_;

    my $db;
    my $s;
    my $bid;

    my $writer = sub {
        $s->execute( $bid, @_ );
    };

    my $commit = sub {
        sleep 1 while !$db->do('commit');
    };

    my $newBook = sub {
        require Compilation::Database;
        $db = Compilation->new(1);
        sleep 1 while !$db->do('begin immediate transaction');
        $bid = $db->addModel( $_[0] );
        sleep 1 while !$db->commit;
        sleep 1 while !$db->do('begin transaction');
        sleep 1
          while !(
            $s = $db->prepare(
                    'insert into data (bid, tab, row, col, v)'
                  . ' values (?, ?, ?, ?, ?)'
            )
          );
    };

    my $processTable = sub { };

    my $yamlCounter = -1;
    my $processYml  = sub {
        my @a;
        while ( my $b = shift ) {
            push @a, $b->[0];
        }
        $writer->( 0, 0, ++$yamlCounter, join "\n", @a, '' );
        $processTable = sub { };
    };

    sub {

        my ( $book, $workbook ) = @_;

        if ( !defined $book ) {    # pruning
            my $gbid;
            sleep 1
              while !(
                $gbid = $db->prepare(
                        'select bid, filename from books '
                      . 'where filename like ? order by filename'
                )
              );
            foreach ( split /:/, $workbook ) {
                $gbid->execute($_);
                while ( my ( $bid, $filename ) = $gbid->fetchrow_array ) {
                    warn "Deleting $filename";
                    my $a = 'y';    # could be <STDIN>
                    if ( $a && $a =~ /y/i ) {
                        warn $db->do( 'delete from books where bid=?',
                            undef, $bid ),
                          ' ',
                          $db->do( 'delete from data where bid=?', undef,
                            $bid );
                    }
                }
            }
            $commit->();
            return;
        }

        $newBook->($book);

        my $sheetNumber = 0;

        warn "Processing $book ($$)\n";
        for my $worksheet ( $workbook->worksheets() ) {
            next if $sheetFilter && !$sheetFilter->( $worksheet->{Name} );
            my ( $row_min, $row_max ) = $worksheet->row_range();
            my ( $col_min, $col_max ) = $worksheet->col_range();
            my $tableTop = 0;
            my @table;
            for my $row ( $row_min .. $row_max ) {
                for my $col ( $col_min .. $col_max ) {
                    my $cell = $worksheet->get_cell( $row, $col );
                    next unless $cell;
                    my $v = $cell->unformatted;
                    next unless defined $v;
                    if ( $col == 0 ) {
                        if ( $v eq '---' ) {
                            $processTable->(@table) if @table;
                            $tableTop     = $row;
                            @table        = ();
                            $processTable = $processYml;
                        }
                        elsif ( $v =~ /^([0-9]{2,})\. / ) {
                            $processTable->(@table) if @table;
                            $tableTop = $row;
                            @table    = ();
                            my $tableNumber = $1;
                            $processTable = sub {
                                my $offset = $#_;
                                --$offset while !defined $_[$offset][0];
                                --$offset
                                  while $offset && defined $_[$offset][0];

                                for my $row ( 0 .. $#_ ) {
                                    my $r  = $_[$row];
                                    my $rn = $row - $offset;
                                    for my $col ( 0 .. $#$r ) {
                                        $writer->(
                                            $tableNumber, $rn, $col, $r->[$col]
                                        ) if defined $r->[$col];
                                    }
                                }

                                $processTable = sub { };
                            };
                        }
                    }
                    $table[ $row - $tableTop ][$col] = $v;
                }
            }
            $processTable->(@table)
              if @table;
        }
        eval {
            warn "Committing $book ($$)\n";
            $commit->();
        };
        warn "$@ for $book ($$)\n" if $@;

    };

}

package NOOP_CLASS;
our $AUTOLOAD;

sub AUTOLOAD {
    no strict 'refs';
    *{$AUTOLOAD} = sub { };
    return;
}

1;
