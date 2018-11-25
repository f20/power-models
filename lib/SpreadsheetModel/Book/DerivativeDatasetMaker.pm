package SpreadsheetModel::Book::DerivativeDatasetMaker;

# Copyright 2014-2018 Franck Latrémolière, Reckon LLP and others.
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

    my %lastResortDatasets;
    foreach my $sourceModel ( values %$sourceModelsMap ) {
        next if $lastResortDatasets{ 0 + $sourceModel };
        my $m = $sourceModel;
        my @lastResortDatasets;
        while ( my $d = $m->{dataset} ) {
            push @lastResortDatasets, $d;
            $m = $m->{sourceModel};
        }
        $lastResortDatasets{ 0 + $sourceModel } = \@lastResortDatasets;
    }

    my $copyClosure = sub {
        my ($cell) = @_;
        defined $cell ? "=$cell" : undef;
    };

    $customActionMap->{$_} ||= $copyClosure
      foreach 'defaultClosure',
      map { /^([0-9]+)/ ? $1 : (); } keys %$sourceModelsMap;

    while ( my ( $theTable, $formulaMaker ) = each %$customActionMap ) {

        my $theHardData = $model->{dataset}{$theTable};
        $model->{dataset}{$theTable} = sub {

            my ( $table, $wb, $ws ) = @_;
            my $hardData =
              $table eq $theTable ? $theHardData : $model->{dataset}{$table};
            return if ref $hardData eq 'CODE';

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

            my $sourceModel =
                 $sourceModelsMap->{$table}
              || $sourceModelsMap->{baseline}
              || $sourceModelsMap->{previous};
            my $sourceTable =
              { map { $_->{number} => $_; } @{ $sourceModel->{inputTables} } }
              ->{$table};
            $applySourceTable->($sourceTable) if $sourceTable;

            if (
                my @columnSpecificRules =
                grep { /^${table}c[0-9]+_?[0-9]*$/ } keys %$sourceModelsMap
              )
            {
                foreach (@columnSpecificRules) {
                    /^${table}c([0-9]+)_?([0-9]*)$/;
                    my ($sourceTable) =
                      grep { $_->{number} == $table; }
                      @{ $sourceModelsMap->{$_}{inputTables} };
                    $applySourceTable->( $sourceTable, $1 .. ( $2 || $1 ) )
                      if $sourceTable;
                }
            }

            map {
                for ( my $icolumn = 1 ; $icolumn < @$_ ; ++$icolumn ) {
                    foreach my $irow ( keys %{ $_->[$icolumn] } ) {
                        $columns[$icolumn]{ @rows == 1 ? $rows[0] : $irow } =
                          $_->[$icolumn]{$irow};
                    }
                }
              } grep { $_ } $hardData,
              map    { $_->{$table} }
              ref $model->{dataOverride} eq 'ARRAY'
              ? @{ $model->{dataOverride} }
              : $model->{dataOverride};

            map {
                for ( my $icolumn = 1 ; $icolumn < @$_ ; ++$icolumn ) {
                    foreach my $irow ( keys %{ $_->[$icolumn] } ) {
                        $columns[$icolumn]{$irow} = $_->[$icolumn]{$irow}
                          unless
                          exists $columns[$icolumn]{'MAGICAL SINGLE ROW NAME'}
                          || exists $columns[$icolumn]{$irow};
                    }
                }
              }
              grep { ref $_ eq 'ARRAY' }
              map  { $_->{$theTable} }
              @{ $lastResortDatasets{ 0 + $sourceModel } };

            \@columns;

          }

    }

}

1;
