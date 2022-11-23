package SpreadsheetModel::Book::DerivativeDatasetMaker;

# Copyright 2014-2022 Franck Latrémolière and others.
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
use Spreadsheet::WriteExcel::Utility;

sub applySourceModelsToDataset {

    my ( $self, $model, $sourceModelsMap, $customActionMap ) = @_;

    my $copyClosure = sub {
        my ($cell) = @_;
        defined $cell ? "=$cell" : undef;
    };

    $customActionMap->{$_} ||= $copyClosure
      foreach 'defaultClosure',
      map { /^([0-9]+)/ ? $1 : (); } keys %$sourceModelsMap;

    while ( my ( $theTableNo, $formulaMaker ) = each %$customActionMap ) {

        $model->{dataset}{$theTableNo} = sub {

            my ( $tableNo, $wb, $ws ) = @_;

            my ( @rows, @columns );

            my $applySourceTable = sub {
                my ( $d, @columnsToApply ) = @_;
                my ( $s, $r, $c ) =
                  ( $d->{columns} ? $d->{columns}[0] : $d )
                  ->wsWrite( $wb, $ws );
                $s = $s->get_name;
                unless (@columnsToApply) {
                    my $width = 0;
                    if ( $d->{columns} ) {
                        $width += $_->lastCol + 1 foreach @{ $d->{columns} };
                    }
                    else {
                        $width = 1 + $d->lastCol;
                    }
                    @columnsToApply = 1 .. $width;
                }

                foreach my $row (
                    $d->{rows} && $d->{rows}{fakeExtraList}
                    ? @{ $d->{rows}{fakeExtraList} }
                    : ()
                  )
                {
                    local $_ = $row;
                    s/.*\n//s;
                    s/[^A-Za-z0-9 -]/ /g;
                    s/- / /g;
                    s/ +/ /g;
                    s/^ //;
                    s/ $//;
                    foreach my $col (@columnsToApply) {
                        $columns[$col]{$_} =
                          $formulaMaker->( undef, $row, $col, $wb, $ws, $_ );
                    }
                }

                @rows =
                  $d->{rows}
                  ? @{ $d->{rows}{list} }
                  : 'MAGICAL SINGLE ROW NAME';

                for ( my $i = 0 ; $i < @rows ; ++$i ) {
                    local $_ = $rows[$i];
                    s/.*\n//s;
                    s/[^A-Za-z0-9 -]/ /g;
                    s/- / /g;
                    s/ +/ /g;
                    s/^ //;
                    s/ $//;
                    my $row = $_;

                    foreach my $col (@columnsToApply) {
                        $columns[$col]{$row} = $formulaMaker->(
                            "'$s'!"
                              . xl_rowcol_to_cell( $r + $i, $c + $col - 1 ),
                            $rows[$i], $col, $wb, $ws, $row
                        );
                    }
                }

            };

            my $theSourceModel =
                 $sourceModelsMap->{$tableNo}
              || $sourceModelsMap->{baseline}
              || $sourceModelsMap->{previous};
            my $sourceTable =
              { map { $_->{number} => $_; }
                  @{ $theSourceModel->{inputTables} } }->{$tableNo};
            $applySourceTable->($sourceTable) if $sourceTable;

            if (
                my @columnSpecificRules =
                grep { /^${tableNo}c[0-9]+_?[0-9]*$/ } keys %$sourceModelsMap
              )
            {
                foreach (@columnSpecificRules) {
                    /^${tableNo}c([0-9]+)_?([0-9]*)$/;
                    my ($sourceTable) =
                      grep { $_->{number} == $tableNo; }
                      @{ $sourceModelsMap->{$_}{inputTables} };
                    $applySourceTable->( $sourceTable, $1 .. ( $2 || $1 ) )
                      if $sourceTable;
                }
            }

            if ( $model->{dataset}
                && ref( my $hardData = $model->{dataset}{$tableNo} ) eq
                'ARRAY' )
            {    # override with our own hard data
                for ( my $icolumn = 1 ; $icolumn < @$hardData ; ++$icolumn ) {
                    foreach my $irow ( keys %{ $hardData->[$icolumn] } ) {
                        $columns[$icolumn]{
                              @rows == 1
                            ? $rows[0]
                            : $irow
                        } = $_->[$icolumn]{$irow};
                    }
                }
            }

            if ( $theSourceModel->{dataset}
                && ref( my $tabData = $theSourceModel->{dataset}{$theTableNo} )
                eq 'ARRAY' )
            {    # fill in any gaps with the source model's hard data
                for ( my $icolumn = 1 ; $icolumn < @$tabData ; ++$icolumn ) {
                    foreach my $irow ( keys %{ $tabData->[$icolumn] } ) {
                        $columns[$icolumn]{$irow} = $tabData->[$icolumn]{$irow}
                          unless
                          exists $columns[$icolumn]{'MAGICAL SINGLE ROW NAME'}
                          || exists $columns[$icolumn]{$irow};
                    }
                }
            }

            \@columns;

        }

    }

}

1;
