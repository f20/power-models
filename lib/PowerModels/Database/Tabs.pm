package PowerModels::Database;

=head Copyright licence and disclaimer

Copyright 2009-2015 Reckon LLP and others.

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

sub _preventOverwriting {
    my $ws = shift;
    my $k  = "@_";
    die join ' ', 'Overwrite prevention:', $ws->get_name, $k, caller
      if exists $ws->{$k};
    undef $ws->{$k};
}

sub tableCompilations {
    my ( $self, $workbookModule, $options, $optionName,
        $fileSearch, $tabSearch ) = @_;
    my $spacing;
    my $numCo = 0;
    {
        $self->do( 'create temporary table models'
              . ' (cid integer primary key, bid int, model char)' );
        my $findCo =
          $self->prepare(
                'select bid, filename from books where filename regexp ?'
              . ' order by filename' );
        my $addCo =
          $self->prepare('insert into models (bid, model) values (?, ?)');
        $findCo->execute($fileSearch);
        my @models;

        while ( my ( $bid, $co ) = $findCo->fetchrow_array ) {
            next unless $co =~ s/\.xlsx?$//is;
            require SpreadsheetModel::Data::DnoAreas;
            push @models,
              [
                $bid, $co,
                SpreadsheetModel::Data::DnoAreas::normaliseDnoName($co)
              ];
            ++$numCo;
        }
        $addCo->execute( @{$_}[ 0, 1 ] )
          foreach sort { $a->[2] cmp $b->[2]; } @models;
    }

    warn "$numCo datasets for $optionName ($fileSearch $tabSearch)";
    $spacing = $numCo + 4;
    my ($leadBid) = $self->selectrow_array('select bid from models limit 1');
    my $tabList =
      $self->prepare('select tab from data where bid=? group by tab');
    $tabList->execute($leadBid);
    my ( $file, $wb, $smallNumberFormat, $bigNumberFormat, $thFormat,
        $thcFormat, $captionFormat, $titleFormat )
      = ('XXX');

    while ( my ($tabNumber) = $tabList->fetchrow_array ) {

        next
          unless $tabNumber > 0
          and $tabNumber =~ /^$tabSearch/;

        my $lastRow =
          $self->selectrow_array(
            'select max(row) from data where bid=? and tab=?',
            undef, $leadBid, $tabNumber );

        my $lastCol =
          $self->selectrow_array(
            'select max(col) from data where bid=? and tab=?',
            undef, $leadBid, $tabNumber );

        my $topRow = 0;

        1 and warn "Table $tabNumber $topRow..$lastRow x $lastCol";

        my @textCols = 0 ? ( 1 .. $lastCol ) : ();
        my @textRows = 0 ? ( 1 .. ( $lastRow - $topRow ) ) : ();
        my @valueCols = 0 ? () : ( 1 .. $lastCol );
        my @valueRows = 0 ? () : ( 1 .. ( $lastRow - $topRow ) );

        unless ( $tabNumber =~ /^$file/ ) {
            $file = substr $tabNumber, 0, 2;
            $wb =
              $workbookModule->new(
                $optionName . '-' . $file . $workbookModule->fileExtension );
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

        {
            my $tableName = $self->selectrow_array(
                'select v from data where bid=? and tab=? and col=0'
                  . ' order by row limit 1',
                undef, $leadBid, $tabNumber
            );
            1 and _preventOverwriting $wsc, 0, 0;
            $wsc->write_string( 0, 0, "$tableName — by column", $titleFormat );
            1 and _preventOverwriting $wsr, 0, 0;
            $wsr->write_string( 0, 0, "$tableName — by row", $titleFormat );
        }

        {
            my $q = $self->prepare('select cid, model from models');
            $q->execute;
            while ( my ( $cid, $co ) = $q->fetchrow_array ) {
                $co =~ s#.*/##;
                $co =~ tr/-/ /;
                1
                  and
                  _preventOverwriting( $wsc, 4 + $spacing * ( $_ - 1 ) + $cid,
                    0 )
                  foreach @valueCols, @textCols;
                $wsc->write_string( 4 + $spacing * ( $_ - 1 ) + $cid,
                    0, $co, $thFormat )
                  foreach @valueCols, @textCols;
                1
                  and _preventOverwriting $wsr,
                  4 + $spacing * ( $_ - 1 ) + $cid, 0
                  foreach @valueRows, @textRows;
                $wsr->write_string( 4 + $spacing * ( $_ - 1 ) + $cid,
                    0, $co, $thFormat )
                  foreach @valueRows, @textRows;
            }
        }

        {
            my $q =
              $self->prepare( 'select col, v from data where bid='
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
                1 and _preventOverwriting $wsr, 4 + $spacing * ( $_ - 1 ), $col
                  foreach @valueRows, @textRows;
                $wsr->write_string( 4 + $spacing * ( $_ - 1 ),
                    $col, $b, $thcFormat )
                  foreach @valueRows, @textRows;
            }
        }

        {
            my $q =
              $self->prepare( 'select row, v from data where bid='
                  . $leadBid
                  . ' and tab='
                  . $tabNumber
                  . ' and col=0 and row>'
                  . $topRow );
            $q->execute;
            while ( my ( $row, $b ) = $q->fetchrow_array ) {
                1
                  and _preventOverwriting $wsc, 4 + $spacing * ( $_ - 1 ),
                  $row - $topRow
                  foreach @valueCols, @textCols;
                $wsc->write_string(
                    4 + $spacing * ( $_ - 1 ),
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
              $self->selectrow_array( 'select 1 from models inner join'
                  . ' data using (bid) where abs(v) > 9999 and tab='
                  . $tabNumber
                  . ' and col='
                  . $col
                  . ' and row>'
                  . $topRow ) ? $bigNumberFormat : $smallNumberFormat;
            my $q =
              $self->prepare( 'select cid, row, v from models inner join'
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
                  4 + $spacing * $col - $spacing + $cid, $row - $topRow;
                $wsc->write(
                    4 + $spacing * $col - $spacing + $cid,
                    $row - $topRow,
                    $v, $format[$col]
                );
            }
        }

        foreach my $row (@valueRows) {
            my $q =
              $self->prepare( 'select cid, col, v from models inner join'
                  . ' data using (bid) where tab='
                  . $tabNumber
                  . ' and col>0 and row='
                  . ( $row + $topRow ) );
            $q->execute;
            while ( my ( $cid, $col, $v ) = $q->fetchrow_array ) {
                1
                  and _preventOverwriting $wsr,
                  4 + $spacing * ( $row - 1 ) + $cid, $col;
                $wsr->write( 4 + $spacing * ( $row - 1 ) + $cid,
                    $col, $v, $format[$col] );
            }
        }

        foreach my $col (@textCols) {
            my $q =
              $self->prepare( 'select cid, row, v from models inner join'
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
                  4 + $spacing * ( $col - 1 ) + $cid, $row - $topRow;
                $wsc->write_string(
                    4 + $spacing * ( $col - 1 ) + $cid,
                    $row - $topRow,
                    $v, $bigNumberFormat
                );
            }
        }

        foreach my $row (@textRows) {
            my $q =
              $self->prepare( 'select cid, col, v from models inner join'
                  . ' data using (bid) where tab='
                  . $tabNumber
                  . ' and col>0 and row='
                  . ( $row + $topRow ) );
            $q->execute;
            while ( my ( $cid, $col, $v ) = $q->fetchrow_array ) {
                1
                  and _preventOverwriting $wsr,
                  4 + $spacing * ( $row - 1 ) + $cid, $col;
                $wsr->write_string( 4 + $spacing * ( $row - 1 ) + $cid,
                    $col, $v, $bigNumberFormat );
            }
        }

    }

}

1;
