package Financial;

=head Copyright licence and disclaimer

Copyright 2015-2017 Franck Latrémolière, Reckon LLP and others.

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
use SpreadsheetModel::Shortcuts ':all';
use SpreadsheetModel::Book::FrontSheet;
require Spreadsheet::WriteExcel::Utility;

sub sheetPriority {
    ( my $model, local $_ ) = @_;
    return 7 if /^Index/;
    return 6 if /^Input/;
    0;
}

sub worksheetsAndClosures {

    my ( $model, $wbook ) = @_;

    my ( $workingsSheet, $inputSheet );

    'Input' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber}    = 14;
        $wbook->{lastSheetNumber} = 19;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   13 );
        $wsheet->set_column( 1, 1,   42 );
        $wsheet->set_column( 2, 250, 13 );
        $wsheet->{nextFree} = 2;
        my ( $sh, $ro, $co ) = Dataset(
            number             => 1400,
            dataset            => $model->{dataset},
            name               => 'Model identification',
            rows               => Labelset( list => [qw(Title Subtitle)] ),
            defaultFormat      => 'puretexthard',
            data               => [ 'no title', 'no subtitle' ],
            usePlaceholderData => 1,
            forwardLinks       => {},
        )->wsWrite( $wbook, $wsheet );
        $sh = $sh->get_name;
        {
            $wbook->{titleAppend} =
                qq%" for "&'$sh'!%
              . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co )
              . qq%&" ("&'$sh'!%
              . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro + 1,
                $co )
              . '&")"';
        }
        $_->wsWrite( $wbook, $wsheet )
          foreach sort { ( $a->{number} || 9909 ) <=> ( $b->{number} || 9909 ) }
          @{ $model->{inputTables} };
        my $nextFree = delete $wsheet->{nextFree};
        Notes( lines =>
              [ 'Input data', '', 'This sheet contains the input data.' ] )
          ->wsWrite( $wbook, $wsheet );
        $wsheet->{nextFree} = $nextFree;
        $inputSheet = $wsheet;
      }

      ,

      $model->{preprocessingSheets}
      ? map {
        my ( $name, $closure ) = @$_;
        $name => sub {
            my ($wsheet) = @_;
            $closure->( $wbook, $wsheet );
        };
      } @{ $model->{preprocessingSheets} }
      : ()

      ,

      'Acct1' => sub {
        my ($wsheet) = @_;
        $workingsSheet = $wsheet;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   42 );
        $wsheet->set_column( 1, 250, 15 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Workings for annual income statement' );
      }

      ,

      'Income' => sub {
        my ($wsheet) = @_;
        $wsheet->{workingsSheet} = $workingsSheet;
        $wsheet->set_landscape;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   42 );
        $wsheet->set_column( 1, 250, 15 );
        $_->wsWrite( $wbook, $wsheet ) foreach Notes(
            name => $model->{oldTerminology}
            ? 'Profit and loss account'
            : 'Income statement (profit and loss)'
          ),
          @{ $model->{incomeTables} };
        delete $wsheet->{workingsSheet};
      }

      ,

      ( $model->{quarterly} ? 'Qly' : 'Mly' ) => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   42 );
        $wsheet->set_column( 1, 250, 15 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Cash calculations on a '
              . ( $model->{quarterly} ? 'quarterly' : 'montly' )
              . ' basis' ), @{ $model->{monthlyTables} };
      }

      ,

      'Acct2' => sub {
        my ($wsheet) = @_;
        $workingsSheet = $wsheet;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   42 );
        $wsheet->set_column( 1, 250, 15 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes(
            name => 'Workings for annual statement of financial position' );
      }

      ,

      'Balance' => sub {
        my ($wsheet) = @_;
        $wsheet->{workingsSheet} = $workingsSheet;
        $workingsSheet = $wsheet;
        $wsheet->set_landscape;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   42 );
        $wsheet->set_column( 1, 250, 15 );
        $_->wsWrite( $wbook, $wsheet ) foreach Notes(
            name => $model->{oldTerminology} ? 'Balance sheet'
            : 'Statement of financial position (balance sheet)'
          ),
          @{ $model->{balanceTables} },
          Notes(
            name => 'Calculations for the '
              . (
                $model->{oldTerminology} ? 'cashflow statement'
                : 'statement of cash flows'
              )
          );
        delete $wsheet->{workingsSheet};
      }

      ,

      'Cashflow' => sub {
        my ($wsheet) = @_;
        $wsheet->{workingsSheet} = $workingsSheet;
        $workingsSheet = $wsheet;
        $wsheet->set_landscape;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   42 );
        $wsheet->set_column( 1, 250, 15 );
        my @tables = @{ $model->{cashflowTables} };
        $_->wsWrite( $wbook, $wsheet ) foreach Notes(
            name => $model->{oldTerminology}
            ? 'Cashflow statement'
            : 'Statement of cash flows'
          ),
          shift @tables;
        delete $wsheet->{workingsSheet};
        $_->wsWrite( $wbook, $wsheet ) foreach @tables;
      }

      ,

      'Ratios' => sub {
        my ($wsheet) = @_;
        $wsheet->{workingsSheet} = $workingsSheet;
        $wsheet->set_landscape;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   45 );
        $wsheet->set_column( 1, 250, 15 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Financial ratios' ),
          @{ $model->{ratioTables} };
        $_->wsWrite( $wbook, $inputSheet ) foreach @{ $model->{inputCharts} };
        $_->wsWrite( $wbook, $wsheet ) foreach @{ $model->{standaloneCharts} };
      }

      ,

      'Index' => SpreadsheetModel::Book::FrontSheet->new(
        model => $model,
        copyright =>
          'Copyright 2015-2017 Franck Latrémolière, Reckon LLP and others.'
      )->closure($wbook);

}

1;
