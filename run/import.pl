#!/usr/bin/env perl

=head Copyright licence and disclaimer

Copyright 2008-2013 Reckon LLP and others.

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

# This script is largely capable of stand-alone operation,
# but looks for libraries in the perl5 and cpan folders too.

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
my $threads1 = 2;

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
        $writer = sqliteWriter();
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
    my $infile = $_;
    unless ( -f $infile ) {
        warn "No such file: $infile";
        next;
    }
    if ($calculate) {
        warn "Calculating $infile";
        my $book2 = $infile;
        $book2 =~ s/\.(xlsx?)$/-$$.$1/i;
        rename $infile, $book2;
        if (`which osascript`) {
            open my $fh, '| osascript';
            print $fh <<EOS;
tell application "Microsoft Excel"
	set wbf to POSIX file "$cwd/$book2"
	set wb to open workbook workbook file name wbf
	set calculate before save to true
	close wb saving yes
end tell
EOS
            close $fh;
            rename $book2, $infile;
        }
        elsif (`which ssconvert`) {
            my $book = $infile;
            $book  =~ s/'/'"'"'/g;
            $book2 =~ s/'/'"'"'/g;
            system qq%ssconvert --recalc '$book2' '$book' 2>/dev/null%;
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
        if (/\.xlsx$/is)
        {
            require Ancillary::XLSX;
            $SIG{__WARN__} = sub { };
            $workbook = Ancillary::XLSX->new($infile);
            delete $SIG{__WARN__};
        }
        else {
            require Spreadsheet::ParseExcel;
            my $parser = Spreadsheet::ParseExcel->new;
            local %SIG;
            $SIG{__WARN__} = sub { };
            $workbook = $parser->Parse($infile);
            delete $SIG{__WARN__};
        }
    };
    warn $@ if $@;
    unless ($workbook) {
        warn "Cannot parse $infile";
        next;
    }
    die "No output specified" unless $writer;
    $writer->( $infile, $workbook );
    exit 0 if defined $pid;
}

waitanypid(0);

sub updateTree {
    my ( $workbook, $tree, $preferArrays ) = @_;
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
                                $tableNumber !~ /00$/ && ( $preferArrays
                                    || $tableNumber =~ /^(?:9|17)/ )
                                ? []
                                : {}
                              ];
                            if ( ref $to->[0] eq 'ARRAY' ) {
                                $to->[0][0] = $v;
                            }
                            else {
                                $to->[0]{'_table'} = $v;
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
                        $to->[$col]{'_column'} = $v;
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
    my $preferArrays = $arg =~ /array/i;
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
        print $h YAML::Dump( updateTree( $workbook, $tree, $preferArrays ) );
    };
}

sub jsonWriter {
    my ($arg) = @_;
    my $dumpAllData  = $arg =~ /all/i;
    my $preferArrays = $arg =~ /array/i;
    my $jsonpp       = !eval 'require JSON';
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
              $jsonpp ? JSON::PP::decode_json(<$fh>) : JSON::decode_json(<$h>);
        }
        open my $h, '>', $json;
        binmode $h;
        print {$h}
          ( $jsonpp ? 'JSON::PP' : 'JSON' )->new->canonical(1)
          ->pretty->utf8->encode(
            updateTree( $workbook, $tree, $preferArrays ) );
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

sub sqliteWriter {
    require DBI;
    use constant { DB_FILE_NAME => '~$database.sqlite' };

    my ( $db, $s, $bid );

    my $writer = sub {
        $s->execute( $bid, @_ );
    };

    my $newDb = sub {
        return if $db && $db->ping;
        die $!
          unless $db = DBI->connect( 'DBI:SQLite:dbname=' . DB_FILE_NAME );
        $db->do($_) foreach grep { $_ } split /;\s*/s, <<EOSQL;
create table if not exists books (
	bid integer primary key,
	filename char
);
create table if not exists data (
	bid integer,
	tab integer,
	row integer,
	col integer,
	v double,
	primary key (bid, tab, col, row)
);
create index if not exists datatcr on data (tab, col, row);
EOSQL

        eval { $db->do('pragma journal_mode=wal') or die $!; };
        warn "Cannot set WAL journal: $@" if $@;

        eval { $db->sqlite_busy_timeout(3_600_000) or die $!; };
        warn "Cannot set timeout: $@" if $@;

        $db->{AutoCommit} = 0;
        $db->{RaiseError} = 1;
        $s =
          $db->prepare( 'insert into data (bid, tab, row, col, v)'
              . ' values (?, ?, ?, ?, ?)' );
    };

    my $cleanup = sub {
        sleep 2 while !$db->do('commit');
        $db->disconnect();
        undef $db;
    };

    my $newBook = sub {
        srand();
        die $! unless $db->do('begin immediate transaction');
        my $done;
        do {
            $bid = int( rand() * 800 ) + 100;
            $done =
              $db->do(
                'insert or ignore into books (filename, bid) values (?, ?)',
                undef, $_[0], $bid );
        } while !$done || $done < 1;
        die $! unless $db->do('commit');
        die $! unless $db->do('begin transaction');
    };

    my $processTable = sub {
        my $tableNumber = shift;
        my $offset      = $#_;
        --$offset while !defined $_[$offset][0];
        --$offset while $offset && defined $_[$offset][0];

        for my $row ( 0 .. $#_ ) {
            my $r  = $_[$row];
            my $rn = $row - $offset;
            for my $col ( 0 .. $#$r ) {
                $writer->( $tableNumber, $rn, $col, $r->[$col] )
                  if defined $r->[$col];
            }
        }
    };

    sub {
        my ( $book, $workbook ) = @_;

        $newDb->();

        if ( !defined $book ) {    # pruning
            my $gbid =
              $db->prepare(
'select bid, filename from books where filename like ? order by filename'
              ) or die $!;
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
            $cleanup->();
            return;
        }

        $newBook->($book);

        my $sheetNumber = 0;

        warn "Processing $book ($$)\n";
        for my $worksheet ( $workbook->worksheets() ) {
            next if $sheetFilter && !$sheetFilter->( $worksheet->{Name} );
            my ( $row_min, $row_max ) = $worksheet->row_range();
            my ( $col_min, $col_max ) = $worksheet->col_range();
            my $tableNumber = --$sheetNumber;
            my $tableTop    = 0;
            my @table;
            for my $row ( $row_min .. $row_max ) {
                for my $col ( $col_min .. $col_max ) {
                    my $cell = $worksheet->get_cell( $row, $col );
                    next unless $cell;
                    my $v = $cell->unformatted;
                    next unless defined $v;
                    if ( $col == 0 && $v =~ /^([0-9]{2,})\. / ) {
                        $processTable->( $tableNumber, @table )
                          if @table && $tableNumber > 0;
                        $tableNumber = $1;
                        $tableTop    = $row;
                        @table       = [$v];
                    }
                    else {
                        $table[ $row - $tableTop ][$col] = $v;
                    }
                }
            }
            $processTable->( $tableNumber, @table )
              if @table && $tableNumber > 0;
        }
        eval {
            warn "Committing $book ($$)\n";
            $cleanup->();
        };
        warn "$@ for $book ($$)\n" if $@;
    };

}

