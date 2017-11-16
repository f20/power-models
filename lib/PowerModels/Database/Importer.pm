package PowerModels::Database::Importer;

=head Copyright licence and disclaimer

Copyright 2009-2016 Franck Latrémolière, Reckon LLP and others.

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
        require PowerModels::Database::Database;
        $db = PowerModels::Database->new(1);
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
            require PowerModels::Database::Database;
            $db ||= PowerModels::Database->new(1);
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
                          $db->do( 'delete from data where bid=?',
                            undef, $bid );
                    }
                }
            }
            $commit->();
            return;
        }

        $newBook->($book);

        warn "process $book ($$)\n";
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
                    eval { $v = Encode::decode( 'UTF-16BE', $v ); }
                      if $v =~ m/\x{0}/;
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
                                --$offset
                                  while !$_[$offset]
                                  || @{ $_[$offset] } <
                                  2;    # slightly risky/broken heuristics here?
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
            $processTable->(@table) if @table;
        }
        eval {
            warn "commit $book ($$)\n";
            $commit->();
        };
        warn "$@ for $book ($$)\n" if $@;

    };

}

1;
