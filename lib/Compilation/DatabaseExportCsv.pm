package Compilation;

=head Copyright licence and disclaimer

Copyright 2009-2012 Reckon LLP and others.

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

sub _writeCsvLine {
    print {shift} join( ',',
        map { local $_ = defined $_ ? $_ : ''; s/"/""/g; qq%"$_"%; } @_ ),
      "\n";
}

sub csvCreate {
    my ( $self, $small ) = @_;
    my $numCo = 0;
    $$self->do(
'create temporary table companies ( cid integer primary key, bid int, company char)'
    );
    my $findCo =
      $$self->prepare('select bid, filename from books order by filename');
    my $addCo =
      $$self->prepare('insert into companies (bid, company) values (?, ?)');
    $findCo->execute;
    while ( my ( $bid, $co ) = $findCo->fetchrow_array ) {
        next unless $co =~ s/\.xlsx?$//is;
        $addCo->execute( $bid, $co );
        ++$numCo;
    }
    warn "$numCo datasets";
    $$self->do(
'create temporary table columns ( colid integer primary key, tab int, col int)'
    );
    $$self->do('create unique index columnstabcol on columns (tab, col)');
    my $tabq = $$self->prepare(
        $small
        ? 'select tab from data where tab like "45__" or tab like "9__" group by tab'
        : 'select tab from data where tab>0 group by tab'
    );
    $tabq->execute();
    while ( my ($tab) = $tabq->fetchrow_array ) {
        warn $tab;
        open my $fh, '>', $tab . '.csv';
        $$self->do('delete from columns');
        $$self->do(
'insert into columns (tab, col) select tab, col from data where tab=? and col>0 group by tab, col order by tab, col',
            undef, $tab
        );
        _writeCsvLine(
            $fh,
            'company',
            'line',
            map { $_->[0] } @{
                $$self->selectall_arrayref(
'select "t" || tab || "c" || col from columns order by colid'
                )
            }
        );
        my $q =
          $$self->prepare(
'select bid, company, row from companies inner join data using (bid) where tab=? and col=1 and row>0 order by company, row'
          );
        $q->execute($tab);
        while ( my ( $bid, $co, $row ) = $q->fetchrow_array ) {
            _writeCsvLine(
                $fh, $co, $row,
                map { $_ && defined $_->[0] ? $_->[0] : undef } @{
                    $$self->selectall_arrayref(
'select v from columns left join data on (data.tab=columns.tab and data.col=columns.col and bid=? and row=?) order by colid',
                        undef, $bid, $row
                    )
                }
            );
        }
        $$self->do('delete from columns');
    }
    if (
        0 < $$self->do(
'insert into columns (tab, col) select tab, col from data where tab>1099 and tab<1181 and col>0 and row=1 and exists (select * from data as d2 where d2.tab=data.tab and d2.col=data.col+1 and d2.bid=data.bid) group by tab, col order by tab, col'
        )
      )
    {
        warn 11;
        open my $fh, '>', '11.csv';
        _writeCsvLine(
            $fh,
            'company',
            map { $_->[0] } @{
                $$self->selectall_arrayref(
'select "t" || tab || "c" || col from columns order by colid'
                )
            }
        );
        my $q =
          $$self->prepare(
            'select bid, company from companies order by company');
        $q->execute;
        while ( my ( $bid, $co ) = $q->fetchrow_array ) {
            _writeCsvLine(
                $fh, $co,
                map { $_->[0] } @{
                    $$self->selectall_arrayref(
'select v from columns left join data on (data.tab=columns.tab and data.col=columns.col and data.row=1 and bid=?) order by colid',
                        undef, $bid
                    )
                }
            );
        }
    }
}

1;
