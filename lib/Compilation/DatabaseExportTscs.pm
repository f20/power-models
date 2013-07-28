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

sub tscsCreateIntermediateTables {
    my ($self) = @_;
    $$self->do('begin immediate transaction') or die $!;
    $$self->do(
'create temporary table models ( bid int, model char, company char, period char, com integer, per integer)'
    );
    my $addCo =
      $$self->prepare(
        'insert into models (bid, model, company, period) values (?, ?, ?, ?)'
      );
    $addCo->execute( $_->[0], $_->[1], $_->[3], join ' ',
        grep { $_ } @{$_}[ 4, 5 ] )
      foreach $self->listModels;

    $$self->do($_) foreach grep { $_ } split /;\n*/, <<EOSQL;
drop table if exists companies;
create table companies (com integer primary key, c char);
create index companiesc on companies (c);
insert into companies (c) select company from models group by company order by company;
update models set com = (select com from companies where c=company); 
drop table if exists periods;
create table periods (per integer primary key, p char);
create index periodsp on periods (p);
insert into periods (p) select period from models group by period order by period;
update models set per = (select per from periods where p=period); 
drop table if exists mytables;
create table mytables (com integer, per integer, tab integer, mytable char, b1 integer, mr integer, primary key (com, per, tab));
insert into mytables (com, per, b1, tab, mr) select com, per, bid, tab, min(row) from models inner join data using (bid) where tab > 0 group by com, per, bid, tab;
update mytables set mytable=(select v from data where bid=b1 and tab=mytables.tab and row=mr and col=0);
drop table if exists columns;
create table columns (com integer, per integer, tab integer, col integer, mycolumn char, primary key (com, per, tab, col));
insert into columns select com, per, tab, col, v from models inner join data using (bid) where row=0 group by com, per, tab, col;
drop table if exists rows;
create table rows (com integer, per integer, tab integer, row integer, myrow char, primary key (com, per, tab, row));
insert into rows select com, per, tab, row, v from models inner join data using (bid) where col=0 group by com, per, tab, row;
drop table if exists tscs;
create table tscs ( com integer, per integer, tab integer, col integer, row integer, v double, primary key (tab, col, row, com, per));
insert into tscs (com, per, tab, col, row, v) select com, per, tab, col, row, v from data inner join models using (bid) where not (row<0 or row=0 and col=0);
drop table models;
EOSQL

=head Not done

insert into tscs (com, per, tab, col, row, v) select com, 0, 0, 0, 0, c from companies;
drop table companies;
insert into tscs (com, per, tab, col, row, v) select 0, per, 0, 0, 0, p from periods;
drop table periods;
insert into tscs (com, per, tab, col, row, v) select com, per, tab, 0, 0, v from data inner join mytables using (bid, tab, row) where col=0;
drop table mytables;

=cut

    sleep 2 while !$$self->commit;
}

sub tscsCreateOutputFiles {
    my ( $self, $workbookModule, $fileExtension, $options ) = @_;
    my ( $file, $wb, $smallNumberFormat, $bigNumberFormat, $thFormat,
        $thcFormat, $captionFormat, $titleFormat );
    $file = '';
    my @topLine =
      map { $_->[0]; }
      @{ $$self->selectall_arrayref('select p from periods order by per') };
    my @topLineCsv =
      ( qw(dno col row), map { local $_ = $_; s/[^0-9]//g; "v$_"; } @topLine );
    @topLine = ( qw(DNO Column Row), @topLine );

    my $tabList =
      $$self->prepare(
'select tab, mytable from mytables group by tab, mytable order by tab, mytable'
      );
    $tabList->execute;
    my $fetch =
      $$self->prepare(
'select c, tscs.per, mycolumn, myrow, tscs.v, abs(tscs.v)<1000 from tscs, companies, columns, rows where tscs.tab=? and tscs.com=companies.com and tscs.com=columns.com and tscs.per=columns.per and tscs.tab=columns.tab and tscs.col=columns.col and tscs.com=rows.com and tscs.per=rows.per and tscs.tab=rows.tab and tscs.row=rows.row order by tscs.com, tscs.tab, tscs.col, tscs.row, tscs.per'
      );
    my $tab1 = 0;

    while ( my ( $tab, $table ) = $tabList->fetchrow_array ) {
        my ( $ws, $row, $csv );
        if ( $options->{wb} ) {
            if ( substr( $tab, 0, 2 ) ne $file ) {
                $wb =
                  $workbookModule->new( 'TSCS-'
                      . ( $file = substr( $tab, 0, 2 ) )
                      . $fileExtension );
                $wb->setFormats($options);
                $smallNumberFormat = $wb->getFormat('0.000copynz');
                $bigNumberFormat   = $wb->getFormat('0copynz');
                $thFormat          = $wb->getFormat('th');
                $thcFormat         = $wb->getFormat('thc');
                $captionFormat     = $wb->getFormat('caption');
                $titleFormat       = $wb->getFormat('notes');
            }
            $ws = $wb->add_worksheet( $tab == $tab1 ? $tab + rand() : $tab );
            $tab1 = $tab;
            $ws->set_column( 0, 2,   36 );
            $ws->set_column( 3, 250, 18 );
            $ws->hide_gridlines(2);
            $ws->freeze_panes( 1, 0 );
            $ws->write_string( 0, 0, $table, $titleFormat );
            $row = 2;
            $ws->write_string( $row, $_, $topLine[$_], $thcFormat )
              for 0 .. $#topLine;
        }

        if ( $options->{csv} ) {
            open $csv, '>', "ts$tab.csv";
        }

        $fetch->execute($tab);
        my $prev = '';
        my @csv  = @topLineCsv;

        while ( my ( $company, $per, $column, $myrow, $value, $small ) =
            $fetch->fetchrow_array )
        {
            my $cur = join '|', $company, $column, $myrow;
            unless ( $prev eq $cur ) {
                $prev = $cur;
                if ($ws) {
                    ++$row;
                    $ws->write_string( $row, 0, $company, $thFormat );
                    $ws->write_string( $row, 1, $column,  $thFormat );
                    $ws->write_string( $row, 2, $myrow,   $thFormat );
                }
                if ($csv) {
                    print {$csv} join( ',', @csv ) . "\n";
                    @csv = (
                        (
                            map { local $_ = $_; s/(["\\])/\\$1/g; qq%"$_"%; }
                              $company,
                            $column,
                            $myrow
                        ),
                        ( map { '' } 3 .. $#topLineCsv )
                    );
                }
            }
            $ws->write( $row, 2 + $per, $value, $small
                ? $smallNumberFormat
                : $bigNumberFormat )
              if $ws;
            $csv[ 2 + $per ] = $value;
        }
        print {$csv} join( ',', @csv ) . "\n" if $csv;
        $ws->autofilter( 2, 0, $row, $#topLine ) if $ws;
    }
}

1;
