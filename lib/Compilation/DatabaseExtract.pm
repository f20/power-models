package Compilation;

=head Copyright licence and disclaimer

Copyright 2013 Franck Latrémolière and others. All rights reserved.

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
use DBI;

sub makeDatabaseReader {
    my $db = DBI->connect('dbi:SQLite:dbname=~$database.sqlite')
      or die "Cannot open sqlite database: $!";
    my $q = $db->prepare(
        'select v from data where bid=? and tab=? and col=? and row=?');
    my $s = $db->prepare(
        'select sum(v) from data where bid=? and tab=? and col=? and row>0');
    my $dataReader = sub {
        my ( $bid, $data ) = @_;
        ( $data->{$_} ) =
          @{ $data->{$_} } > 2
          ? $db->selectrow_array( $q, undef, $bid, @{ $data->{$_} } )
          : $db->selectrow_array( $s, undef, $bid, @{ $data->{$_} } )
          foreach grep { ref( $data->{$_} ) eq 'ARRAY' } keys %$data;
        $data;
    };
    my $lt = $db->prepare('select tab from data where bid=? group by tab');
    my $bq = $db->prepare('select bid, filename from books');
    $bq->execute;
    my %files;
    while ( my ( $bid, $filename ) = $bq->fetchrow_array ) {
        $filename =~ s/\.xlsx?//i;
        $filename .= ".$bid" if exists $files{$filename};
        my %exists = ( bid => $bid );
        $lt->execute($bid);
        while ( my ($tab) = $lt->fetchrow_array ) {
            undef $exists{$tab};
        }
        $files{$filename} = \%exists;
    }
    $dataReader, \%files;
}

sub extract1076from1001 {
    my ( $dataReader, $bookTableIndexHash ) = makeDatabaseReader();
    foreach ( sort keys %$bookTableIndexHash ) {
        my $d = $dataReader->(
            $bookTableIndexHash->{$_}{bid},
            {
                base        => [ 1001, 5, 4 ],
                cdcm        => [ 1001, 5, 39 ],
                noncdcmded  => [ 1001, 5, 34 ],
                k           => [ 1001, 5, 24 ],
                passthrough => [ 1001, 5, 10 ],
            }
        );
        require YAML;
        YAML::DumpFile(
            ( (/^(.+?)-20/)[0] || $_ ) . "-1076.yml",
            {
                1076 => [
                    [],
                    [
                        '"Allowed revenue" (£/year)',
                        $d->{cdcm} -
                          $d->{passthrough} -
                          $d->{k} -
                          $d->{noncdcmded}
                    ],
                    [ '"Pass-through charges" (£/year)', $d->{passthrough} ],
                    [
                        "Adjustment for previous year's"
                          . ' under (over) recovery (£/year)',
                        $d->{k}
                    ],
                    [
                        'Revenue raised outside this model (£/year)',
                        -$d->{noncdcmded}
                    ]
                ]
            }
        );
    }
}

1;
