package Ancillary::DatabaseExport;

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
use DBI;
use Encode 'decode_utf8';

sub new {
    my $databaseHandle = DBI->connect( 'dbi:SQLite:dbname=~$database.sqlite',
        '', '', { sqlite_unicode => 1, AutoCommit => 0, } )
      or die "Cannot open sqlite database: $!";
    bless \$databaseHandle, shift;
}

sub DESTROY { ${ $_[0] }->disconnect; }

sub summariesByCompany {
    my ( $self, $workbookModule, $fileExtension, $name, @sheets ) = @_;
    my %bidMap;
    foreach (
        @{ $$self->selectall_arrayref('select bid, filename from books') } )
    {
        ( my $bid, local $_ ) = @$_;
        s/\.xlsx?$//is;
        s/-r[0-9]+$//is;
        s/-(?:FCP|LRIC[a-z]*)//is;
        next unless s/-([^-]+)$//s;
        $bidMap{$_}{$1} = $bid;
    }

    my $getTitle =
      $$self->prepare(
        'select v from data where bid=? and tab=? and col=? and row=0');

    my $getData =
      $$self->prepare(
        'select row, v from data where bid=? and tab=? and col=? and row>0');

    while ( my ( $company, $optionhr ) = each %bidMap ) {
        warn "Making $company-$name$fileExtension";
        my $wb = $workbookModule->new("$company-$name$fileExtension");
        $wb->setFormats;
        my $thcFormat   = $wb->getFormat('thc');
        my $titleFormat = $wb->getFormat('notes');
        my $thtarFormat = $wb->getFormat('thtar');

        foreach (@sheets) {
            my ( $sheetName, $columnsar ) = %$_;
            my @columns = @$columnsar or die $sheetName;
            my $ws = $wb->add_worksheet($sheetName);
            $ws->write_string( 0, 0, ( shift @columns ), $titleFormat );
            $ws->set_column( 0, 0, 18 );
            $ws->hide_gridlines(2);
            $ws->freeze_panes( 1, 1 );

            for ( my $c = 0 ; $c < @columns ; ++$c ) {
                local $_ = $columns[$c];
                my $title;
                ( $_, $title ) = %$_ if ref $_;
                my ( $tab, $col, $opt ) = /t([0-9]+)c([0-9]+)-(.*)/;
                next unless $tab;
                unless ($title) {
                    $getTitle->execute( $optionhr->{$opt}, $tab, $col );
                    ($title) = $getTitle->fetchrow_array;
                    eval { $title = decode_utf8 $title; $title =~ s/&amp;/&/g; };
                }
                $title ||= 'No title';
                my $format = $wb->getFormat(
                      $title =~ /kVArh|kWh/ ? '0.000copynz'
                    : $title =~ /\bp\//     ? '0.00copynz'
                    : $title =~ /%/         ? '%softpm'
                    : $title =~ /change/i   ? '0softpm'
                    : '0softnz'
                );
                $ws->set_column( 1 + $c, 1 + $c, $title =~ /name/i ? 54 : 18 );
                unless ($c) {
                    $getData->execute( $optionhr->{$opt}, $tab, 0 );
                    while ( my ( $r, $v ) = $getData->fetchrow_array ) {
                        $ws->write( $r + 2, 0, $v, $thtarFormat );
                    }
                }
                $ws->write_string( 2, $c + 1, $title, $thcFormat );
                $getData->execute( $optionhr->{$opt}, $tab, $col );
                while ( my ( $r, $v ) = $getData->fetchrow_array ) {
                    eval { $v = decode_utf8 $v; $v =~ s/&amp;/&/g; };
                    $ws->write( $r + 2, $c + 1, $v, $format );
                }

            }
        }
    }
}

sub _preventOverwriting {
    my $ws = shift;
    my $k  = "@_";
    die join " ", $k, caller if exists $ws->{$k};
    undef $ws->{$k};
}

sub tableCompilations {
    my ( $self, $workbookModule, $fileExtension, $options, $optionName,
        $fileSearch, $tabSearch )
      = @_;
    my $spacing;
    my $numCo = 0;
    {
        $$self->do(
'create temporary table models ( cid integer primary key, bid int, model char)'
        );
        my $findCo =
          $$self->prepare(
'select bid, filename from books where filename regexp ? order by filename'
          );
        my $addCo =
          $$self->prepare('insert into models (bid, model) values (?, ?)');
        $findCo->execute($fileSearch);
        my @models;

        while ( my ( $bid, $co ) = $findCo->fetchrow_array ) {
            next unless $co =~ s/\.xlsx?$//is;
            my $sort = $co;
            $sort =~ s/CE-NEDL/NPG-Northeast/;
            $sort =~ s/CE-YEDL/NPG-Yorkshire/;
            $sort =~ s/^NP-/NPG-/;
            $sort =~ s/CN-East/WPD-EastM/;
            $sort =~ s/CN-West/WPD-WestM/;
            $sort =~ s/EDFEN/UKPN/;
            $sort =~ s/WPD-Wales/WPD-SWales/;
            $sort =~ s/WPD-West\b/WPD-SWest/;
            push @models, [ $bid, $co, $sort ];
            ++$numCo;
        }
        $addCo->execute( @{$_}[ 0, 1 ] )
          foreach sort { $a->[2] cmp $b->[2]; } @models;
    }

    warn "$numCo datasets for $optionName ($fileSearch $tabSearch)";
    $spacing = $numCo + 3;
    my $leadBid = $$self->selectrow_array('select bid from models limit 1');
    my $tabList =
      $$self->prepare('select tab from data where bid=? group by tab');
    $tabList->execute($leadBid);
    my ( $file, $wb, $smallNumberFormat, $bigNumberFormat, $thFormat,
        $thcFormat, $captionFormat, $titleFormat )
      = ('XXX');

    while ( my $tabNumber = $tabList->fetchrow_array ) {

        warn $tabNumber;
        next
          unless $tabNumber > 0
          and $tabNumber =~ /^$tabSearch/;

        my $lastRow =
          $$self->selectrow_array(
            'select max(row) from data where bid=? and tab=?',
            undef, $leadBid, $tabNumber );

        my $lastCol =
          $$self->selectrow_array(
            'select max(col) from data where bid=? and tab=?',
            undef, $leadBid, $tabNumber );

        my $topRow = 0;

        0 and warn "Table $tabNumber $topRow..$lastRow x $lastCol";

        my @textCols = 0 ? ( 1 .. $lastCol ) : ();
        my @textRows = 0 ? ( 1 .. ( $lastRow - $topRow ) ) : ();
        my @valueCols = 0 ? () : ( 1 .. $lastCol );
        my @valueRows = 0 ? () : ( 1 .. ( $lastRow - $topRow ) );

        unless ( $tabNumber =~ /^$file/ ) {
            $file = substr $tabNumber, 0, 2;
            $wb = $workbookModule->new("$optionName-$file$fileExtension");
            $wb->setFormats($options);
            $smallNumberFormat = $wb->getFormat('0.000copynz');
            $bigNumberFormat   = $wb->getFormat('0copynz');
            $thFormat          = $wb->getFormat('th');
            $thcFormat         = $wb->getFormat('thc');
            $captionFormat     = $wb->getFormat('caption');
            $titleFormat       = $wb->getFormat('notes');
        }

        my $wsc = $wb->add_worksheet( $tabNumber . 'c' );
        $wsc->set_column( 0, 0,   38 );
        $wsc->set_column( 1, 250, 18 );
        $wsc->hide_gridlines(2);
        $wsc->freeze_panes( 1, 1 );

        my $wsr = $wb->add_worksheet( $tabNumber . 'r' );
        $wsr->set_column( 0, 0,   38 );
        $wsr->set_column( 1, 250, 18 );
        $wsr->hide_gridlines(2);
        $wsr->freeze_panes( 1, 1 );

        if (undef) {
            my $wst = $wb->add_worksheet( $tabNumber . 't' );
            $wst->set_column( 0, 0,   38 );
            $wst->set_column( 1, 250, 18 );
            $wst->hide_gridlines(2);
            $wst->freeze_panes( 1, 1 );
        }

        {
            my $tableName = $$self->selectrow_array(
'select v from data where bid=? and tab=? and col=0 order by row limit 1',
                undef, $leadBid, $tabNumber
            );
            1 and _preventOverwriting $wsc, 0, 0;
            $wsc->write_string( 0, 0, "$tableName — by column", $titleFormat );
            1 and _preventOverwriting $wsr, 0, 0;
            $wsr->write_string( 0, 0, "$tableName — by row", $titleFormat );
        }

        {
            my $q = $$self->prepare('select cid, model from models');
            $q->execute;
            while ( my ( $cid, $co ) = $q->fetchrow_array ) {
                $co =~ s#.*/##;
                $co =~ tr/-/ /;
                1
                  and
                  _preventOverwriting( $wsc, 3 + $spacing * ( $_ - 1 ) + $cid,
                    0 )
                  foreach @valueCols, @textCols;
                $wsc->write_string( 3 + $spacing * ( $_ - 1 ) + $cid,
                    0, $co, $thFormat )
                  foreach @valueCols, @textCols;
                1
                  and _preventOverwriting $wsr,
                  3 + $spacing * ( $_ - 1 ) + $cid, 0
                  foreach @valueRows, @textRows;
                $wsr->write_string( 3 + $spacing * ( $_ - 1 ) + $cid,
                    0, $co, $thFormat )
                  foreach @valueRows, @textRows;
            }
        }

        {
            my $q =
              $$self->prepare( 'select col, v from data where bid='
                  . $leadBid
                  . ' and tab='
                  . $tabNumber
                  . ' and col>0 and row='
                  . $topRow );
            $q->execute;
            while ( my ( $col, $b ) = $q->fetchrow_array ) {
                1
                  and _preventOverwriting $wsc, 2 + $col * $spacing - $spacing,
                  0;
                $wsc->write_string( 2 + $col * $spacing - $spacing,
                    0, $b, $captionFormat );
                1 and _preventOverwriting $wsr, 3 + $spacing * ( $_ - 1 ), $col
                  foreach @valueRows, @textRows;
                $wsr->write_string( 3 + $spacing * ( $_ - 1 ),
                    $col, $b, $thcFormat )
                  foreach @valueRows, @textRows;
            }
        }

        {
            my $q =
              $$self->prepare( 'select row, v from data where bid='
                  . $leadBid
                  . ' and tab='
                  . $tabNumber
                  . ' and col=0 and row>'
                  . $topRow );
            $q->execute;
            while ( my ( $row, $b ) = $q->fetchrow_array ) {
                1
                  and _preventOverwriting $wsc, 3 + $spacing * ( $_ - 1 ),
                  $row - $topRow
                  foreach @valueCols, @textCols;
                $wsc->write_string(
                    3 + $spacing * ( $_ - 1 ),
                    $row - $topRow,
                    $b, $thcFormat
                ) foreach @valueCols, @textCols;
                1
                  and _preventOverwriting $wsr,
                  2 + ( $row - $topRow - 1 ) * $spacing, 0;
                $wsr->write_string( 2 + ( $row - $topRow - 1 ) * $spacing,
                    0, $b, $captionFormat );
            }
        }

        my @format;

        foreach my $col (@valueCols) {
            $format[$col] =
              $$self->selectrow_array( 'select 1 from models inner join'
                  . ' data using (bid) where abs(v) > 9999 and tab='
                  . $tabNumber
                  . ' and col='
                  . $col
                  . ' and row>'
                  . $topRow ) ? $bigNumberFormat : $smallNumberFormat;
            my $q =
              $$self->prepare( 'select cid, row, v from models inner join'
                  . ' data using (bid) where tab='
                  . $tabNumber
                  . ' and col='
                  . $col
                  . ' and row>'
                  . $topRow );
            $q->execute;
            while ( my ( $cid, $row, $v ) = $q->fetchrow_array ) {
                1
                  and _preventOverwriting $wsc,
                  3 + $spacing * $col - $spacing + $cid, $row - $topRow;
                $wsc->write(
                    3 + $spacing * $col - $spacing + $cid,
                    $row - $topRow,
                    $v, $format[$col]
                );
            }
        }

        foreach my $row (@valueRows) {
            my $q =
              $$self->prepare( 'select cid, col, v from models inner join'
                  . ' data using (bid) where tab='
                  . $tabNumber
                  . ' and col>0 and row='
                  . ( $row + $topRow ) );
            $q->execute;
            while ( my ( $cid, $col, $v ) = $q->fetchrow_array ) {
                1
                  and _preventOverwriting $wsr,
                  3 + $spacing * ( $row - 1 ) + $cid, $col;
                $wsr->write( 3 + $spacing * ( $row - 1 ) + $cid,
                    $col, $v, $format[$col] );
            }
        }

        foreach my $col (@textCols) {
            my $q =
              $$self->prepare( 'select cid, row, v from models inner join'
                  . ' data using (bid) where tab='
                  . $tabNumber
                  . ' and col='
                  . $col
                  . ' and row>'
                  . $topRow );
            $q->execute;
            while ( my ( $cid, $row, $v ) = $q->fetchrow_array ) {
                1
                  and _preventOverwriting $wsc,
                  3 + $spacing * ( $col - 1 ) + $cid, $row - $topRow;
                $wsc->write_string(
                    3 + $spacing * ( $col - 1 ) + $cid,
                    $row - $topRow,
                    $v, $bigNumberFormat
                );
            }
        }

        foreach my $row (@textRows) {
            my $q =
              $$self->prepare( 'select cid, col, v from models inner join'
                  . ' data using (bid) where tab='
                  . $tabNumber
                  . ' and col>0 and row='
                  . ( $row + $topRow ) );
            $q->execute;
            while ( my ( $cid, $col, $v ) = $q->fetchrow_array ) {
                1
                  and _preventOverwriting $wsr,
                  3 + $spacing * ( $row - 1 ) + $cid, $col;
                $wsr->write_string( 3 + $spacing * ( $row - 1 ) + $cid,
                    $col, $v, $bigNumberFormat );
            }
        }

    }

}

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

sub tscsCreateIntermediateTables {
    my ($self) = @_;
    $$self->do('begin immediate transaction') or die $!;
    $$self->do(
'create temporary table models ( bid int, model char, company char, period char, com integer, per integer)'
    );
    my $findCo = $$self->prepare('select bid, filename from books');
    my $addCo =
      $$self->prepare(
        'insert into models (bid, model, company, period) values (?, ?, ?, ?)'
      );
    $findCo->execute;
    my @models;

    while ( my ( $bid, $co ) = $findCo->fetchrow_array ) {
        next unless $co =~ s/\.xlsx?$//is;
        local $_ = $co;
        s/^M-//;
        s/CE-NEDL/NP-Northeast/;
        s/CE-YEDL/NP-Yorkshire/;
        s/CN-East/WPD-EastM/;
        s/CN-West/WPD-WestM/;
        s/EDFEN/UKPN/;
        s/NPG-/NP-/;
        s/SP-/SPEN-/;
        s/SSE-/SSEPD-/;
        s/WPD-Wales/WPD-SWales/;
        s/WPD-West\b/WPD-SWest/;
        push @models, [ $bid, $co, $_ ];
    }
    $addCo->execute( $_->[0], $_->[1],
        map { local $_ = $_; tr/-/ /; $_; } ( $_->[2] =~ m#^(.*?)-([0-9-]+)# ) )
      foreach sort { $a->[2] cmp $b->[2]; } @models;
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
