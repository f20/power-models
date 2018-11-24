package SpreadsheetModel::ChartSeries;

# Copyright 2017 Franck Latrémolière, Reckon LLP and others.
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

use Spreadsheet::WriteExcel::Utility;

use constant {
    CS_OBJ => 0,
    CS_ROW => 1,
    CS_COL => 2,
};

sub new {
    my ( $class, $obj, $row, $col ) = @_;
    bless [ $obj, $row, $col ], $class;
}

sub valuesNameCategories {
    my ( $self, $wb, $ws ) = @_;
    my $ar = $self->[CS_OBJ]{columns};
    die 'This currently only supports extracting a row from a Columnset'
      unless $ar && defined $self->[CS_ROW] && !defined $self->[CS_COL];
    my ( $dataWorksheet, $rFirstColumn, $cFirstColumn ) =
      $ar->[0]->wsWrite( $wb, $ws );
    $dataWorksheet = $dataWorksheet->get_name;
    my ( $wLastColumn, $rLastColumn, $cLastColumn ) =
      $ar->[$#$ar]->wsWrite( $wb, $ws );
    (
        values => "='$dataWorksheet'!"
          . xl_rowcol_to_cell( $rFirstColumn + $self->[CS_ROW], $cFirstColumn )
          . ':'
          . xl_rowcol_to_cell( $rLastColumn + $self->[CS_ROW], $cLastColumn ),
        name       => $ar->[0]->objectShortName,
        categories => "='$dataWorksheet'!"
          . xl_rowcol_to_cell( $rFirstColumn - 1, $cFirstColumn ) . ':'
          . xl_rowcol_to_cell( $rLastColumn - 1,  $cLastColumn ),
    );
}

sub lastCol {
    my ($self) = @_;
    my $ar = $self->[CS_OBJ]{columns};
    die 'This currently only supports extracting a row from a Columnset'
      unless $ar && defined $self->[CS_ROW] && !defined $self->[CS_COL];
    0 + @$ar;
}

sub lastRow {
    0;
}

1;
