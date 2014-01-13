package SpreadsheetModel::Logger;

=head Copyright licence and disclaimer

Copyright 2008-2013 Franck Latrémolière, Reckon LLP and others.

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

require SpreadsheetModel::Object;
our @ISA = qw(SpreadsheetModel::Object);

use Spreadsheet::WriteExcel::Utility;

sub log {
    my $self = shift;
    push @{ $self->{objects} }, grep {
             $_->{location}
          || !$_->{sources}
          || $#{ $_->{sources} }
          || $_->{cols} != $_->{sources}[0]{cols}
          || $_->{rows} != $_->{sources}[0]{rows}

          # this is probably a heuristic for some more precise conditions
          # improved in June 2009

    } @_;
}

sub check {
    my ($self) = @_;
    $self->{lines} = [ SpreadsheetModel::Object::splitLines( $self->{lines} ) ]
      if $self->{lines};
    $self->{objects} = [];
    return;
}

sub lastCol {
    3;
}

sub lastRow {
    $_[0]->{realRows} ? $#{ $_[0]->{realRows} } : $#{ $_[0]->{objects} };
}

sub wsWrite {
    my ( $self, $wb, $ws, $row, $col ) = @_;
    ( $row, $col ) = ( ( $ws->{nextFree} ||= -1 ) + 1, 0 )
      unless defined $row && defined $col;
    $ws->write( $row++, $col, "$self->{name}", $wb->getFormat('caption') );
    my $numFormat0 = $wb->getFormat('0softnz');
    my $numFormat1 = $wb->getFormat('0.000soft');
    my $textFormat = $wb->getFormat('text');
    my $linkFormat = $wb->getFormat('link');
    if ( $self->{lines} ) {
        $ws->write( $row++, $col, "$_", $textFormat )
          foreach @{ $self->{lines} };
    }

    my @h = ( 'Worksheet', 'Data table', 'Type of table' );
    push @h, 'Dimensions', 'Count', 'Average' if $wb->{logAll};

    $ws->write( $row, $col + $_, "$h[$_]", $wb->getFormat('th') ) for 0 .. $#h;
    $row++;

    my @objectList = sort {
        ( $a->{$wb}{worksheet}{sheetNumber} || 666 )
          <=> ( $b->{$wb}{worksheet}{sheetNumber} || 666 )
    } grep { $_->{$wb}{worksheet} } @{ $self->{objects} };

    my $r = 0;
    my %columnsetDone;
    foreach my $obj (@objectList) {

        my $cset;

        unless ( $wb->{logAll} ) {
            $cset = $obj->{location};
            undef $cset unless ref $cset eq 'SpreadsheetModel::Columnset';
            if ($cset) {
                next if exists $columnsetDone{$cset};
                undef $columnsetDone{$cset};
            }
        }

        my ( $wo, $ro, $co ) = @{ $obj->{$wb} }{qw(worksheet row col)};
        my $ty = $cset ? $cset->objectType : $obj->objectType;
        my $ce = xl_rowcol_to_cell( $ro - 1, $co );
        my $wn = $wo ? $wo->get_name : 'BROKEN LINK';
        $wn =~ s/\000//g;    # squash strange rare bug
        my $na = $cset ? "$cset->{name}" : "$obj->{name}";
        0 and $ws->set_row( $row + $r, undef, undef, 1 ) unless $na;
        $self->{realRows}[$r] = $na;
        $ws->write_url( $row + $r, $col + 1, "internal:'$wn'!$ce", $na,
            $linkFormat );
        $ws->write_string( $row + $r, $col + 2, $ty, $textFormat );
        $ws->write_string( $row + $r, $col,     $wn, $textFormat );

        if ( $wb->{logAll} && $obj->isa('SpreadsheetModel::Dataset') ) {
            my ( $wss, $rows, $cols ) = $obj->wsWrite( $wb, $ws );
            my $wsn = $wss ? $wss->get_name : 'BROKEN LINK';
            my $c1 =
              Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $rows,
                $cols, 0, 0 );
            my $c2 = Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                $rows + $obj->lastRow,
                $cols + $obj->lastCol,
                0, 0
            );
            my $range = "'$wsn'!$c1:$c2";
            $ws->write( $row + $r, $col + 3,
                ( 1 + $obj->lastRow ) . ' × ' . ( 1 + $obj->lastCol ),
                $textFormat );
            $ws->write( $row + $r, $col + 4, "=COUNT($range)",   $numFormat0 );
            $ws->write( $row + $r, $col + 5, "=AVERAGE($range)", $numFormat1 );
        }
        ++$r;
    }

    $ws->autofilter( $row - 1, $col, $row + $r - 1, $col + 2 );
    0 and $ws->filter_column( $col, 'x <> ""' );

    $ws->{nextFree} = $row + $r
      unless $ws->{nextFree} > $row + $r;
}

1;
