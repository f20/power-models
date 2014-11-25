package Compilation::DataExtraction;

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
use Ancillary::ParallelRunning;

sub ymlWriter {
    my ($arg) = @_;
    my $options = {
        $arg =~ /array/i ? ( preferArrays => 1 ) : (),
        $arg =~ /min/i   ? ( minimum      => 1 ) : (),
    };
    sub {
        my ( $book, $workbook ) = @_;
        die unless $book;
        my $file = $book;
        $file =~ s/\.xl[a-z]+?$//is;
        my $tree;
        require YAML;
        if ( -e $file ) {
            open my $h, '<', "$file.yml";
            binmode $h, ':utf8';
            local undef $/;
            $tree = YAML::Load(<$h>);
        }
        my %trees = _extractInputData( $workbook, $tree, $options );
        while ( my ( $key, $value ) = each %trees ) {
            open my $h, '>', "$file$key.yml";
            binmode $h, ':utf8';
            print $h YAML::Dump($value);
        }
    };
}

sub jsonWriter {
    my ($arg) = @_;
    my $options = {
        $arg =~ /array/i ? ( preferArrays => 1 ) : (),
        $arg =~ /min/i   ? ( minimum      => 1 ) : (),
    };
    sub {
        my ( $book, $workbook ) = @_;
        die unless $book;
        my $file = $book;
        $file =~ s/\.xl[a-z]+?$//is;
        my $tree;
        my $jsonpp = !eval 'require JSON';
        require JSON::PP if $jsonpp;
        if ( -e $file ) {
            open my $h, '<', "$file.json";
            binmode $h;
            local undef $/;
            $tree =
              $jsonpp ? JSON::PP::decode_json(<$h>) : JSON::decode_json(<$h>);
        }
        my %trees = _extractInputData( $workbook, $tree, $options );
        while ( my ( $key, $value ) = each %trees ) {
            open my $h, '>', "$file$key.json";
            binmode $h;
            print {$h}
              ( $jsonpp ? 'JSON::PP' : 'JSON' )->new->canonical(1)
              ->pretty->utf8->encode($value);
        }
    };
}

sub _extractInputData {
    my ( $workbook, $tree, $options ) = @_;
    my ( %byWorksheet, %dirty, $dirtyOverall );
    for my $worksheet ( $workbook->worksheets() ) {
        my ( $row_min, $row_max ) = $worksheet->row_range();
        my ( $col_min, $col_max ) = $worksheet->col_range();
        my ( $tableNumber, $columnHeadingsRow, $to1, $to2 );
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
                            $to1 = $tree->{$tableNumber};
                            $to2 = [
                                (
                                          ref $to1->[0] ne 'HASH'
                                      and ref $to1->[0] eq 'ARRAY'
                                      || $options->{preferArrays}
                                      && $tableNumber !~ /00$/
                                )                     ? [$v]
                                : $options->{minimum} ? {}
                                :                       { '_table' => $v }
                            ];
                            $to1 = [
                                $options->{preferArrays}
                                  && $tableNumber !~ /00$/ ? [$v]
                                : $options->{minimum} ? {}
                                :                       { '_table' => $v }
                              ]
                              unless $to1->[0];
                            $dirtyOverall ||= $dirty{$tableNumber};
                            $dirty{$tableNumber} = 1;
                        }
                        else {
                            if ($v) {
                                $v =~ s/[^A-Za-z0-9-]/ /g;
                                $v =~ s/- / /g;
                                $v =~ s/ +/ /g;
                                $v =~ s/^ //;
                                $v =~ s/ $//;
                                $rowName = $v eq '' ? '•' : $v;
                                $to1->[0][ $row - $columnHeadingsRow ] =
                                  $to2->[0][ $row - $columnHeadingsRow ] =
                                  $rowName
                                  if ref $to1->[0] eq 'ARRAY'
                                  and defined $columnHeadingsRow;
                            }
                        }
                    }
                    else {
                        if ( ref $to1->[0] eq 'HASH' )
                        {    # old-style table comment
                            $to1->[0]{'_note'} = $to2->[0]{'_note'} = $v if $v;
                        }
                        else {
                            $columnHeadingsRow = $row - 1
                              unless defined $columnHeadingsRow;
                            $to1->[0][ $row - $columnHeadingsRow ] =
                              $to2->[0][ $row - $columnHeadingsRow ] = $v;
                        }
                    }
                }
                elsif ($tableNumber) {
                    if ( !$rowName ) {
                        $columnHeadingsRow = $row;
                        if ( ref $to1->[0] eq 'ARRAY' ) {
                            $to1->[$col][0] ||= $v;
                            $to2->[$col][0] ||= $v;
                        }
                        elsif ( !$options->{minimum} ) {
                            $to1->[$col]{'_column'} ||= $v;
                            $to2->[$col]{'_column'} ||= $v;
                        }
                    }
                    elsif (ref $cell->{Format}
                        && !$cell->{Format}{Lock}
                        && ( $v || $to1->[$col] ) )
                    {
                        if ( ref $to1->[$col] eq 'ARRAY' ) {
                            $to1->[$col][ $row - $columnHeadingsRow ] =
                              $to2->[$col][ $row - $columnHeadingsRow ] = $v;
                        }
                        else {
                            $to1->[$col]{$rowName} = $to2->[$col]{$rowName} =
                              $v;
                        }
                        $tree->{$tableNumber} ||= $to1;
                        $byWorksheet{" $worksheet->{Name}"}{$tableNumber} =
                          $to2;
                    }
                }
            }
        }
    }
    '', $tree, $dirtyOverall ? %byWorksheet : ();
}

sub databaseWriter {

    my ($settings) = @_;

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
            require Compilation::Database;
            $db ||= Compilation->new(1);
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

        warn "Processing $book ($$)\n";
        for my $worksheet ( $workbook->worksheets() ) {
            next
              if $settings->{sheetFilter}
              && !$settings->{sheetFilter}->($worksheet);
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

sub checksumWriter {
    sub {
        my ( $book, $workbook ) = @_;
        $book =~ s#.*/##;
        for my $worksheet ( $workbook->worksheets() ) {
            my ( $row_min, $row_max ) = $worksheet->row_range();
            my ( $col_min, $col_max ) = $worksheet->col_range();
            my $tableNumber;
          ROW: for my $row ( $row_min .. $row_max ) {
                my $rowName;
              COL: for my $col ( $col_min .. $col_max ) {
                    my $cell = $worksheet->get_cell( $row, $col );
                    my $v;
                    $v = $cell->unformatted if $cell;
                    next unless defined $v;
                    if ( $col == 0 ) {
                        if ( !ref $cell->{Format} || $cell->{Format}{Lock} ) {
                            if ( $v && $v =~ /^([0-9]{2,})\. / ) {
                                $tableNumber = $1;
                                next ROW;
                            }
                        }
                        else {
                            next ROW if $v;
                            next COL;
                        }
                    }
                    if ( $tableNumber && $v =~ /model checksum/i ) {
                        my $check =
                          $worksheet->get_cell( $row + 1, $col )->unformatted;
                        if ( $v =~ /7/ ) {
                            $check = 5.5e-8 + 1e-7 * $check;
                            $check =~ s/.*\.(...)(....)5.*/$1 $2/;
                        }
                        warn "$book\t$tableNumber\t$check\n";
                    }
                }
            }
        }
    };
}

sub _setSha1 {
    my ( $scalar, $key, $value ) = @_;
    if ( $key =~ s#^([^/]*)/## ) {
        $scalar->{$1} = _setSha1( $scalar->{$1}, $key, $value );
    }
    else { $scalar->{$key} = $value; }
    $scalar;
}

sub jbzWriter {
    sub {
        my ( $book, $workbook ) = @_;
        my %scalars;
        for my $worksheet ( $workbook->worksheets() ) {
            my $scalar;
            my ( $row_min, $row_max ) = $worksheet->row_range();
            for my $row ( $row_min .. $row_max ) {
                my $cell = $worksheet->get_cell( $row, 0 );
                my $v;
                $v = $cell->unformatted if $cell;
                $scalar = _setSha1( $scalar, $1, $2 )
                  if $v && $v =~ /(\S+): ([0-9a-fA-F]{40})/;
            }
            $scalars{ $worksheet->{Name} } = $scalar if $scalar;
        }
        return unless %scalars;
        use JSON;
        $book =~ s/\.xl\S+//i;
        $book .= '.jbz';
        $book =~ s/'/'"'"'/g;
        open my $fh, qq%|bzip2>'$book'% or goto FAIL;
        binmode $fh or goto FAIL;
        print {$fh}
          JSON->new->canonical(1)
          ->utf8->pretty->encode(
            keys %scalars > 1 ? \%scalars : values %scalars )
          or goto FAIL;
        return;
      FAIL: warn $!;
        return;
    };
}

1;
