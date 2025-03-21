﻿package SpreadsheetModel::Logger;

# Copyright 2008-2021 Franck Latrémolière and others.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice,
# this list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
# this list of conditions and the following disclaimer in the documentation
# and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
# EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
# THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

use warnings;
use strict;
use utf8;

require SpreadsheetModel::Object;
our @ISA = qw(SpreadsheetModel::Object);

use Spreadsheet::WriteExcel::Utility;

sub log {
    my $logger = shift;
    push @{ $logger->{objects} }, grep {
             $_->{location}
          || !$_->{sources}
          || $#{ $_->{sources} }
          || $_->{cols} != $_->{sources}[0]{cols}
          || $_->{rows} != $_->{sources}[0]{rows}
    } @_;
}

sub check {
    my ($logger) = @_;
    $logger->{lines} =
      [ SpreadsheetModel::Object::splitLines( $logger->{lines} ) ]
      if $logger->{lines};
    $logger->{objects} = [];
    return;
}

sub lastCol {
    3;
}

sub lastRow {
    $_[0]->{realRows} ? $#{ $_[0]->{realRows} } : $#{ $_[0]->{objects} };
}

sub loggableObjects {
    my ($logger) = @_;
    my %containerSeen;
    my @list;
    foreach my $obj ( grep { defined $_ } @{ $logger->{objects} } ) {
        my @displayList = $obj;
        if ( my $container = $obj->{location} ) {
            if ( UNIVERSAL::isa( $container, 'SpreadsheetModel::Columnset' ) ) {
                @displayList = ()
                  unless $logger->{showColumns} || $container->{logColumns}
                  and grep {
                    ref $_ ne 'SpreadsheetModel::Stack'
                      || @{ $_->{sources} } > 1;
                  } @{ $container->{columns} };
                unless ( exists $containerSeen{$container} ) {
                    unshift @displayList, $container;
                    undef $containerSeen{$container};
                }
            }
            elsif (
                UNIVERSAL::isa( $container, 'SpreadsheetModel::CalcBlock' ) )
            {
                @displayList = ()
                  unless $logger->{showRows} || $container->{logRows};
                unless ( exists $containerSeen{$container} ) {
                    unshift @displayList, $container;
                    undef $containerSeen{$container};
                }
            }
        }
        push @list, @displayList;
    }
    @list;
}

sub _normaliseForSorting {
    local $_ = "$_[0]";
    $_ = "0$_" while /^[0-9]{1,3}\./s;
    $_;
}

sub wsWrite {

    my ( $logger, $wb, $ws, $row, $col ) = @_;
    ( $row, $col ) = ( ( $ws->{nextFree} ||= -1 ) + 1, 0 )
      unless defined $row && defined $col;
    if ( $logger->{name} ) {
        $ws->set_row( $row, $wb->{captionRowHeight} );
        $ws->write( $row++, $col, "$logger->{name}",
            $wb->getFormat('caption') );
        ++$row;
    }

    my $headerFormat = $wb->getFormat( [ base => 'th',        locked => 0 ] );
    my $numFormat0   = $wb->getFormat( [ base => '0softnz',   locked => 0 ] );
    my $numFormat1   = $wb->getFormat( [ base => '0.000soft', locked => 0 ] );
    my $textFormat   = $wb->getFormat( [ base => 'text',      locked => 0 ] );
    my $linkFormat   = $wb->getFormat( [ base => 'link',      locked => 0 ] );

    if ( $logger->{lines} ) {
        $ws->write( $row++, $col, "$_", $textFormat )
          foreach @{ $logger->{lines} };
    }

    my @h = ( 'Worksheet', 'Table', 'Type of table' );
    push @h, 'Dimensions', 'Count', 'Average' if $logger->{showDetails};

    $ws->write( $row, $col + $_, "$h[$_]", $headerFormat ) for 0 .. $#h;
    $row++;

    my @objectList =
      sort {
        _normaliseForSorting( $a->{name} )
          cmp _normaliseForSorting( $b->{name} );
      }
      grep { $_->{$wb}{worksheet} && $_->{name} } @{ $logger->{objects} };

    my $r = 0;
    my %containerSeen;

    foreach ( $logger->loggableObjects ) {
        my ( $wo, $ro, $co ) =
          (   $_->isa('SpreadsheetModel::Columnset') ? $_->{columns}[0]
            : $_->isa('SpreadsheetModel::CalcBlock') ? $_->{items}[0]
            :   $_ )->wsWrite( $wb, $ws );
        my $url = $_->wsUrl($wb);
        my $wn =
            $wo
          ? $wo->get_name
          : die "Broken link to $_->{name} $_->{debug}";
        $wn =~ s/\000//g;    # squash strange rare bug
        my $na = "$_->{name}";
        $logger->{realRows}[$r] = $na;

        if ( $_->{logColumns} ) {
            $ws->write_string( $row + $r, $col + 1, $na, $textFormat );
            $ws->write_string( $row + $r, $col + 2, '(not used further)',
                $textFormat )
              if $logger->{showFinalTables}
              && !$_->{forwardLinks}
              && !UNIVERSAL::isa( $_->{location},
                'SpreadsheetModel::Objectset' );
        }
        else {
            $ws->write_url( $row + $r, $col + 1, $url, $na, $linkFormat );
            $ws->write_string(
                $row + $r,
                $col + 2,
                $_->objectType
                  . (
                    $logger->{showFinalTables}
                      && !$_->{forwardLinks}
                      && !UNIVERSAL::isa( $_->{location},
                        'SpreadsheetModel::Objectset' )
                    ? ' (not used further)'
                    : ''
                  ),
                $textFormat
            );
        }
        $ws->write_string( $row + $r, $col, $wn, $textFormat );

        if (   $logger->{showDetails}
            && $_->isa('SpreadsheetModel::Dataset') )
        {
            my $c1 = xl_rowcol_to_cell( $ro, $co );
            my $c2 = xl_rowcol_to_cell( $ro + $_->lastRow, $co + $_->lastCol );
            my $range = "'$wn'!$c1:$c2";
            $ws->write( $row + $r, $col + 3,
                ( 1 + $_->lastRow ) . ' × ' . ( 1 + $_->lastCol ),
                $textFormat );
            $ws->write( $row + $r, $col + 4, "=COUNT($range)",   $numFormat0 );
            $ws->write( $row + $r, $col + 5, "=AVERAGE($range)", $numFormat1 );
        }

        ++$r;

    }

    $ws->autofilter(
        $row - 1, $col,
        $row + $r - 1,
        $col + ( $logger->{showDetails} ? 5 : 2 )
    );

    $ws->{protectionOptions} ||= {
        autofilter => 1,
        sort       => 1,
    };

    $ws->{nextFree} = $row + $r
      unless $ws->{nextFree} > $row + $r;

}

1;
