package SpreadsheetModel::Data::Database;

=head Copyright licence and disclaimer

Copyright 2009-2014 Reckon LLP and others.

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
    print {shift} join(
        ',',
        map {
            local $_ = defined $_ ? $_ : '';
            s/"/""/g;
            s/\r//sg;
            s/\n/\\n /sg;
            qq%"$_"%;
        } @_
      ),
      "\n";
}

sub _normalisedRowName {
    ( local $_ ) = @_;
    s/[^A-Za-z0-9-]/ /g;
    s/- / /g;
    s/ +/ /g;
    s/^ //;
    s/ $//;
    $_;
}

sub dumpTallCsv {
    my ( $self, $inputOnlyFlag ) = @_;
    open my $fh, '>', '~$' . $$ . '.csv';
    binmode $fh, ':utf8';
    _writeCsvLine(
        $fh,
        'Model number',
        'Area',
        'Period',
        'Options',
        'Table number',
        'Table name',
        'Column number',
        'Column name',
        'Row number',
        'Full row name',
        'Normalised row label',
        'Value'
    );
    $self->do(
        'create temporary table tabminrow (bid int, tab int, minrow int)');
    $self->do(
        'insert into tabminrow
            select bid, tab, min(row) as minrow
                from data '
          . ( $inputOnlyFlag ? 'where tab>999 and tab<2000 ' : '' )
          . 'group by bid, tab'
    );
    $self->do('create temporary table tabnames (bid int, tab int, name char)');
    $self->do(
        'insert into tabnames
            select tabminrow.bid, tabminrow.tab, data.v as name
                from tabminrow left join data on (
                    data.bid=tabminrow.bid and data.tab=tabminrow.tab and row=minrow and col=0
                )'
    );
    my $fetch = $self->prepare(
        'select
            data.bid,
            books.company,
            books.period,
            books.option,
            data.tab,
            tabnames.name,
            data.col,
            dcol.v,
            data.row,
            drow.v,
            data.v
        from data
            inner join books using (bid)
            inner join tabnames on (data.bid=tabnames.bid and data.tab=tabnames.tab)
            left join data as drow on (data.bid=drow.bid and data.tab=drow.tab and data.row=drow.row and drow.col=0)
            left join data as dcol on (data.bid=dcol.bid and data.tab=dcol.tab and data.col=dcol.col and dcol.row=0)
            where data.row>0 and data.col>0'
    );
    $fetch->execute;

    while ( my (@row) = $fetch->fetchrow_array ) {
        splice @row, 10, 0, _normalisedRowName( $row[9] );
        _writeCsvLine( $fh, @row );
    }
    close $fh;
    rename '~$' . $$ . '.csv',
      '~$' . ( $inputOnlyFlag ? 'input-data' : 'all-data' ) . '.csv';
}

sub dumpEdcmCsv {

    my ( $self, $allTables ) = @_;
    my $numCo = 0;
    $self->do( 'create temporary table companies'
          . ' ( cid integer primary key, bid int, company char, settings char )'
    );
    my $findCo =
      $self->prepare( 'select books.bid, filename, v from books'
          . ' left join data on ('
          . 'books.bid=data.bid and tab=0 and row=0 and col=0'
          . ') order by filename' );
    my $addCo =
      $self->prepare(
        'insert into companies (bid, company, settings) values (?, ?, ?)');
    $findCo->execute;
    my %exclude =
      ( identification => undef, method => undef, wantTables => undef, );
    while ( my ( $bid, $co, $option ) = $findCo->fetchrow_array ) {
        next unless $co =~ s/\.xlsx?$//is;
        $co =~ s/.*\///s;
        $addCo->execute(
            $bid, $co,
            join "\n",
            map {
                local $_ = $_;
                s/-(?:FCP|LRIC)//;
                $_;
              } grep { /^([a-zA-Z]\S+): \S/ && !exists $exclude{$1}; }
              split /\n/,
            $option
        );
        ++$numCo;
    }
    warn "$numCo datasets";
    return unless $numCo;

    $self->do( 'create temporary table columns'
          . ' ( colid integer primary key, tab int, col int )' );
    $self->do('create unique index columnstabcol on columns (tab, col)');

    my $fetchSets =
      $self->prepare('select settings from companies group by settings');
    my $numSets = $fetchSets->execute;

    while ( my ($set) = $fetchSets->fetchrow_array ) {
        my $setName = '_';
        $setName = "CSV-$1" if $set =~ /^template: (\S.*)/m;
        $setName =~ s/['"]//g;
        $setName .= '_' while !mkdir($setName) && !( -d $setName && -w _ );
        my $tempFile = '~$tmp-' . $$ . '.csv';
        {
            my %zero = (
                checksums               => 0,
                dcp183                  => 0,
                dcp185                  => 0,
                dcp189                  => 0,
                dcp206                  => 0,
                legacy201               => 0,
                lowerIntermittentCredit => 0,
            );
            foreach ( split /\n/, $set ) {
                next unless /^(\S+): '?([^']*)'?$/;
                $zero{$1} = $2 unless $2 eq '' || $2 eq '~';
            }
            open my $fh, '>', $tempFile;
            my @k = sort keys %zero;
            print {$fh} join( ',', @k ) . "\n";
            print {$fh} join( ',', map { /,/ ? qq%"$_"% : $_; } @zero{@k} )
              . "\n";
            close $fh;
            rename $tempFile, $setName . '/0.csv';
        }

        my $tabq = $self->prepare(
            $allTables
            ? 'select tab from data inner join companies using (bid) where settings=?'
              . ' and tab>0 group by tab'
            : 'select tab from data inner join companies using (bid) where settings=?'
              . ' and tab like "9__" or tab=4501 or tab=4601 group by tab'
        );
        $tabq->execute($set);
        while ( my ($tab) = $tabq->fetchrow_array ) {
            warn $tab;
            open my $fh, '>', $tempFile;
            $self->do('delete from columns');
            $self->do(
                'insert into columns (tab, col) select tab, col from'
                  . ' companies inner join data using (bid) '
                  . 'where settings=? and tab=?'
                  . ' group by tab, col order by tab, col',
                undef, $set, $tab
            );
            _writeCsvLine(
                $fh,
                'company',
                'line',
                map { $_->[0] } @{
                    $self->selectall_arrayref(
                            'select "t" || tab || "c" || col'
                          . ' from columns where col>0 order by colid'
                    )
                }
            );
            my $q =
              $self->prepare( 'select bid, company, row from '
                  . 'companies inner join data using (bid)'
                  . ' where settings=? and tab=? and col=1 and row>0'
                  . ' order by company, row' );
            $q->execute( $set, $tab );
            while ( my ( $bid, $co, $row ) = $q->fetchrow_array ) {
                _writeCsvLine(
                    $fh, $co,
                    map { $_ && defined $_->[0] ? $_->[0] : undef } @{
                        $self->selectall_arrayref(
                            'select v from columns left join data on '
                              . '(data.tab=columns.tab and data.col=columns.col and bid=? and row=?)'
                              . ' order by colid',
                            undef, $bid, $row
                        )
                    }
                );
            }
            close $fh;
            rename $tempFile, $setName . '/' . $tab . '.csv';
            $self->do('delete from columns');
        }

        foreach my $group ( [ 11, 1099, 1181 ], [ 47, 4789, 4800 ],
            [ 48, 4799, 4900 ] )
        {
            if (
                0 < $self->do(
                    'insert into columns (tab, col) select tab, col from data'
                      . ' inner join companies using (bid) '
                      . 'where settings=? and tab>? and tab<? and col>0 and row=1 and '
                      . 'exists (select * from data as d2 where d2.tab=data.tab'
                      . ' and d2.col=data.col+1 and d2.bid=data.bid) '
                      . 'group by tab, col order by tab, col',
                    undef,
                    $set,
                    @{$group}[ 1, 2 ]
                )
              )
            {
                warn $group = $group->[0];
                open my $fh, '>', $tempFile;
                _writeCsvLine(
                    $fh,
                    'company',
                    map { $_->[0] } @{
                        $self->selectall_arrayref(
                                'select "t" || tab || "c" || col '
                              . 'from columns order by colid'
                        )
                    }
                );
                my $q =
                  $self->prepare( 'select bid, company from companies'
                      . ' where settings=? order by company' );
                $q->execute($set);
                while ( my ( $bid, $co ) = $q->fetchrow_array ) {
                    _writeCsvLine(
                        $fh, $co,
                        map { $_->[0] } @{
                            $self->selectall_arrayref(
                                'select v from columns left join data on'
                                  . ' (data.tab=columns.tab and data.col=columns.col'
                                  . ' and data.row=1 and bid=?)'
                                  . ' order by colid',
                                undef, $bid
                            )
                        }
                    );
                }
                close $fh;
                rename $tempFile, "$setName/$group.csv";
            }
            $self->do('delete from columns');
        }
    }

}

1;
