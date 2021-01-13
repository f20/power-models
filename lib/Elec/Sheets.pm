package Elec;

# Copyright 2012-2019 Franck Latrémolière and others.
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
use SpreadsheetModel::Shortcuts ':all';
use SpreadsheetModel::Book::FrontSheet;

sub finishModel {
    my ( $model, $wbook ) = @_;
    my $append = ( $model->{idAppend}{$wbook} || '' )
      . ( $model->{checksumAppend}{$wbook} || '' );
    foreach ( @{ $model->{titleWrites}{$wbook} } ) {
        my ( $ws, $row, $col, $n, $fmt ) = @$_;
        $ws->write( $row, $col, qq%="$n"$append%, $fmt,
                'Not calculated: '
              . 'open in spreadsheet app and allow calculations' );
    }
}

sub worksheetsAndClosures {

    my ( $model, $wbook ) = @_;

    $wbook->{lastSheetNumber} = 49;

    my @detailedTables;
    push @detailedTables, @{ $model->{detailedTables} }
      if $model->{detailedTables};
    push @detailedTables, @{ $model->{detailedTablesBottom} }
      if $model->{detailedTablesBottom};

    'Input' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber} = 15;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   20 );
        $wsheet->set_column( 1, 250, 20 );
        $model->{titleWrites}{$wbook} = [];
        $wbook->{titleWriter} =
          sub { push @{ $model->{titleWrites}{$wbook} }, [@_]; };
        $model->{inputTables} ||= [];
        my $idTable = Dataset(
            number  => 1500,
            dataset => $model->{dataset},
            name    => 'Company, charging year, data version',
            cols    => Labelset(
                list => [
                    'Company', $model->{interpolator} ? 'Not used' : 'Year',
                    'Version',
                ]
            ),
            defaultFormat => 'puretexthard',
            data          => [
                'no company',
                $model->{interpolator} ? undef : 'no year',
                'no data version',
            ],
            usePlaceholderData => 1,
            forwardLinks       => {},
        );
        $_->wsWrite( $wbook, $wsheet ) foreach Notes(
            name  => 'Input data',
            lines => 'This sheet contains the input data.'
          ),
          $idTable,
          $model->{table1653}
          ? Notes( lines => 'Individual user data', location => 'Customers', )
          : (),
          sort { ( $a->{number} || 9909 ) <=> ( $b->{number} || 9909 ) }
          @{ $model->{inputTables} };
        require Spreadsheet::WriteExcel::Utility;
        my ( $sh, $ro, $co ) = $idTable->wsWrite( $wbook, $wsheet );
        $sh = $sh->get_name;
        $model->{idAppend}{$wbook} =
            qq%&" for "&'$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co )
          . q%&" in "&%
          . (
              $model->{interpolator}
            ? $model->{interpolator}->chargingPeriodLabel( $wbook, $wsheet )
            : "'$sh'!"
              . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                $ro, $co + 1
              )
          )
          . qq%&" ("&'$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co + 2 )
          . '&")"';
      }

      ,

      $model->{table1653}
      ? (
        'Customers' => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 2 );
            $wsheet->set_column( 0, 0, $model->{ulist}          ? 50 : 20 );
            $wsheet->set_column( 1, 1, $model->{table1653Names} ? 50 : 20 );
            $wsheet->set_column( 2, 250, 20 );
            $model->{table1653}->wsWrite( $wbook, $wsheet );
        }
      )
      : ()

      ,

      'Volumes' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   20 );
        $wsheet->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Volumes' ), @{ $model->{volumeTables} };
      }

      ,

      $model->{checkTables} && @{ $model->{checkTables} }
      ? (
        'Checks' => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->set_column( 0, 0,   20 );
            $wsheet->set_column( 1, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( name => 'Network usage checks' ),
              @{ $model->{checkTables} };
        }
      )
      : ()

      ,

      $model->{bandTables} && @{ $model->{bandTables} }
      ? (
        'Bands' => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->set_column( 0, 0,   20 );
            $wsheet->set_column( 1, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( name => 'Time band analysis' ),
              @{ $model->{bandTables} };
        }
      )
      : ()

      ,

      'Costs' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   20 );
        $wsheet->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Relevant costs and charges' ),
          @{ $model->{costTables} };
      }

      ,

      'Buildup' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   20 );
        $wsheet->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Tariff build-up' ),
          @{ $model->{buildupTables} };
      }

      ,

      'Tariffs' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   20 );
        $wsheet->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Tariffs' ), @{ $model->{tariffTables} };
        if ( $model->{tariffChecksum} ) {
            require Spreadsheet::WriteExcel::Utility;
            my ( $sh, $ro, $co ) =
              $model->{tariffChecksum}->wsWrite( $wbook, $wsheet );
            $sh = $sh->get_name;
            my $cell = qq%'$sh'!%
              . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co );
            my $checksumText = qq%" [checksum "&TEXT($cell,"000 0000")&"]"%;
            $model->{checksumAppend}{$wbook} =
              qq%&IF(ISNUMBER($cell),$checksumText,"")%;
        }
      }

      ,

      'Revenues' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   20 );
        $wsheet->set_column( 1, 250, 20 );
        my $noLinks = $wbook->{noLinks};
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Revenues' ), @{ $model->{revenueTables} };
        $wbook->{noLinks} = $noLinks;
      }

      ,

      @detailedTables
      ? (
        'Details' => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->set_column( 0, 0,
                $model->{ulist} || $model->{table1653Names} ? 50 : 20 );
            $wsheet->set_column( 1, 1,
                $model->{detailedTablesNames} ? 50 : 20 );
            $wsheet->set_column( 2, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( name => 'Detailed tables' ),
              @detailedTables;
        }
      )
      : ()

      ,

      'Index' => SpreadsheetModel::Book::FrontSheet->new(
        model     => $model,
        copyright => 'Copyright 2012-2021 Franck Latrémolière and others.'
      )->closure($wbook);

}

1;
