package Elec;

=head Copyright licence and disclaimer

Copyright 2012-2016 Franck Latrémolière, Reckon LLP and others.

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
use Spreadsheet::WriteExcel::Utility;

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
        $wsheet->set_column( 0, 0,   36 );
        $wsheet->set_column( 1, 250, 20 );
        $wsheet->{nextFree} = 2;
        $model->{inputTables} ||= [];
        my ( $sh, $ro, $co ) = Dataset(
            number        => 1500,
            dataset       => $model->{dataset},
            name          => 'Company, charging year, data version',
            cols          => Labelset( list => [qw(Company Year Version)] ),
            defaultFormat => 'puretexthard',
            data          => [ 'no company', 'no year', 'no data version' ],
            usePlaceholderData => 1,
            forwardLinks       => {},
        )->wsWrite( $wbook, $wsheet );
        $sh = $sh->get_name;
        {
            $wbook->{titleAppend} =
                qq%" for "&'$sh'!%
              . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co )
              . qq%&" in "&'$sh'!%
              . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro,
                $co + 1 )
              . qq%&" ("&'$sh'!%
              . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro,
                $co + 2 )
              . '&")"';
        }
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->{table1653}
          ? Notes( lines => 'Individual user data', location => 'Customers', )
          : (),
          sort { ( $a->{number} || 9909 ) <=> ( $b->{number} || 9909 ) }
          @{ $model->{inputTables} };
        my $nextFree = delete $wsheet->{nextFree};
        Notes( lines =>
              [ 'Input data', '', 'This sheet contains the input data.' ] )
          ->wsWrite( $wbook, $wsheet );
        $wsheet->{nextFree} = $nextFree;
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
        $wsheet->set_column( 0, 0,   36 );
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
            $wsheet->set_column( 0, 0,   36 );
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
            $wsheet->set_column( 0, 0,   36 );
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
        $wsheet->set_column( 0, 0,   36 );
        $wsheet->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Relevant costs and charges' ),
          @{ $model->{costTables} };
      }

      ,

      'Buildup' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   36 );
        $wsheet->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Tariff build-up' ),
          @{ $model->{buildupTables} };
      }

      ,

      'Tariffs' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   36 );
        $wsheet->set_column( 1, 250, 20 );
        my $noLinks = $wbook->{noLinks};
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Tariffs' ), @{ $model->{tariffTables} };
        $wbook->{noLinks} = $noLinks;
      }

      ,

      'Revenues' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   36 );
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
        model => $model,
        copyright =>
          'Copyright 2012-2016 Franck Latrémolière, Reckon LLP and others.'
      )->closure($wbook);

}

1;
