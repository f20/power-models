package PowerModels::Database;

=head Copyright licence and disclaimer

Copyright 2009-2018 Reckon LLP and others.

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
    $self->do('begin immediate transaction') or die $!;
    $self->do( 'create temporary table models'
          . ' (bid int, model char, company char, period char, com integer, per integer)'
    );
    my $addCo =
      $self->prepare(
        'insert into models (bid, model, company, period) values (?, ?, ?, ?)'
      );
    $addCo->execute( $_->[0], $_->[1], $_->[2], join ' ',
        grep { $_ } @{$_}[ 3, 4 ] )
      foreach $self->listModels;

    $self->do($_) foreach grep { $_ } split /;\n*/, <<EOSQL;
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
create table mytables (tab integer, per integer, bidone integer, minrow integer, primary key (tab, per));
insert into mytables (per, bidone, tab, minrow)
    select per, bid, tab, min(row) from models inner join data using (bid) where tab > 0 group by per, tab;
drop table if exists columns;
create table columns (com integer, per integer, tab integer, col integer, mycolumn char, primary key (com, per, tab, col));
insert into columns
    select com, per, tab, col, v from models inner join data using (bid) where row=0 group by com, per, tab, col;
drop table if exists rows;
create table rows (com integer, per integer, tab integer, row integer, myrow char, primary key (com, per, tab, row));
insert into rows
    select com, per, tab, row, v from models inner join data using (bid) where col=0 group by com, per, tab, row;
drop table if exists rownumbers;
create table rownumbers (rowtab integer, rowname char, rownumber integer, primary key (rowtab, rowname));
insert into rownumbers select tab, myrow, min(row) from rows group by myrow, tab;
drop table if exists tscs;
create table tscs ( com integer, per integer, tab integer, col integer, row integer, v double, primary key (tab, col, row, com, per));
insert into tscs (com, per, tab, col, row, v)
    select com, per, tab, col, row, v from data inner join models using (bid) where not (row<0 or row=0 and col=0);
drop table models;
EOSQL

    sleep 1 while !$self->commit;

}

sub tscsCompilation {

    my ( $self, $workbookModule, $options ) = @_;

    require Spreadsheet::WriteExcel::Utility;

    my $tabList =
      $self->prepare('select tab from mytables group by tab order by tab');
    $tabList->execute;
    my $fetch =
      $self->prepare(
            'select c, tscs.per, mycolumn, myrow, tscs.v, abs(tscs.v)<1000'
          . ' from tscs, companies, columns, rows, rownumbers'
          . ' where rowname=myrow and tscs.tab=rowtab and tscs.tab=? and tscs.com=companies.com'
          . ' and tscs.com=columns.com and tscs.per=columns.per and tscs.tab=columns.tab'
          . ' and tscs.col=columns.col and tscs.com=rows.com and tscs.per=rows.per and tscs.tab=rows.tab'
          . ' and tscs.row=rows.row'
          . ' order by tscs.com, tscs.tab, tscs.col, rownumber, myrow, tscs.per'
      );
    my $qTableTitles =
      $self->prepare( 'select v from periods left join mytables'
          . ' on (periods.per=mytables.per and mytables.tab=?)'
          . ' left join data'
          . ' on (bid=bidone and row=minrow and col=0 and data.tab=mytables.tab)'
          . ' order by periods.per' );

    my (
        $wb,          $smallNumberFormat, $bigNumberFormat,
        $thFormat,    $thcFormat,         $captionFormat,
        $titleFormat, $ws,                $row,
        $col,
    );
    my $newWorkbook = sub {
        my $wb = $workbookModule->new( $_[0] )
          or die "Cannot create workbook $_[0]";
        $wb->setFormats($options);
        $smallNumberFormat = $wb->getFormat('0.000copynz');
        $bigNumberFormat   = $wb->getFormat('0copynz');
        $thFormat          = $wb->getFormat('th');
        $thcFormat         = $wb->getFormat('thc');
        $captionFormat     = $wb->getFormat('caption');
        $titleFormat       = $wb->getFormat('notes');
        $wb;
    };

    my $file = '';
    my @topLine =
      map { $_->[0]; }
      @{ $self->selectall_arrayref('select p from periods order by per') };
    push @topLine, 'Changes' if @topLine > 1;

    if ( $options->{singleSheet} ) {
        $wb = $newWorkbook->( 'TSCS' . $workbookModule->fileExtension );
        $ws = $wb->add_worksheet('TSCS');
        $ws->set_column( 0, 0,   18 );
        $ws->set_column( 1, 1,   9 );
        $ws->set_column( 2, 3,   36 );
        $ws->set_column( 4, 250, 18 );
        $ws->hide_gridlines(2);
        $ws->freeze_panes( 1, 0 );
        $ws->write_string( 0, 0, 'Data compilation', $titleFormat );
        @topLine = ( qw(DNO Table Column Row), @topLine );
        $row = 2;
        $ws->write_string( $row, $_, $topLine[$_], $thcFormat )
          for 0 .. $#topLine;
    }
    else {
        @topLine = ( qw(DNO Column Row), @topLine );
    }

    while ( my ($tab) = $tabList->fetchrow_array ) {
        next
          if $options->{tablesMatching} && !grep { $tab =~ /$_/ }
          @{ $options->{tablesMatching} };

        if ( !$options->{singleSheet} && substr( $tab, 0, 2 ) ne $file ) {
            $wb =
              $newWorkbook->( 'TSCS-'
                  . ( $file = substr( $tab, 0, 2 ) )
                  . $workbookModule->fileExtension );
        }
        if ( $options->{singleSheet} ) {
            $col = 4;
            ++$row;
        }
        else {
            $ws = $wb->add_worksheet($tab);
            $ws->set_column( 0, 0,   18 );
            $ws->set_column( 1, 2,   36 );
            $ws->set_column( 3, 250, 18 );
            $ws->hide_gridlines(2);
            $ws->freeze_panes( 1, 0 );
            $ws->write_string( 0, 0, "Table $tab", $titleFormat );
            $row = 2;
            $ws->write_string( $row, $_, $topLine[$_], $thcFormat )
              for 0 .. $#topLine;
            $col = 3;
        }
        $qTableTitles->execute($tab);
        my @columns;

        while ( my ($tableName) = $qTableTitles->fetchrow_array ) {
            push @columns, $col;
            if ($tableName) {
                if ( $options->{singleSheet} ) {
                    $ws->write_string( $row, $col, $tableName, $thFormat );
                }
                else {
                    $ws->write_string( 0, $col, $tableName, $titleFormat );
                }
            }
            ++$col;
        }

        $fetch->execute($tab);
        my $prev = '';

        while ( my ( $company, $per, $column, $myrow, $value, $small ) =
            $fetch->fetchrow_array )
        {
            my $cur = join '|', $company, $column, $myrow;
            $column =~ s/[\r\n]+/ /gs;
            unless ( $prev eq $cur ) {
                $prev = $cur;
                ++$row;
                if ( $options->{singleSheet} ) {
                    $ws->write_string( $row, 0, $company, $thFormat );
                    $ws->write_string( $row, 1, $tab,     $thFormat );
                    $ws->write_string( $row, 2, $column,  $thFormat );
                    $ws->write_string( $row, 3, $myrow,   $thFormat );
                }
                else {
                    $ws->write_string( $row, 0, $company, $thFormat );
                    $ws->write_string( $row, 1, $column,  $thFormat );
                    $ws->write_string( $row, 2, $myrow,   $thFormat );
                }
                $ws->write(
                    $row,
                    $#topLine,
                    '=' . join(
                        '+', 0,
                        map {
                            '('
                              . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                                $row, $columns[ $_ - 1 ] )
                              . '<>'
                              . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                                $row, $columns[$_] )
                              . ')';
                        } 1 .. $#columns
                    ),
                    $bigNumberFormat
                ) if @columns > 1;
            }
            $ws->write(
                $row, 2 + $per + ( $options->{singleSheet} ? 1 : 0 ),
                $value, $small
                ? $smallNumberFormat
                : $bigNumberFormat
            );
        }
        $ws->autofilter( 2, 0, $row, $#topLine );
    }

}

1;
