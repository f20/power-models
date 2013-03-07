package EUoS;

=head Copyright licence and disclaimer

Copyright 2012-2013 Franck Latrémolière, Reckon LLP and others.

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
require Spreadsheet::WriteExcel::Utility;
use SpreadsheetModel::Shortcuts ':all';

use POSIX ();

sub worksheetsAndClosures {

    my ( $model, $wbook ) = @_;

    $wbook->{lastSheetNumber} = 49;

    'Input' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber} = 15;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   48 );
        $wsheet->set_column( 1, 250, 20 );
        return if $model->{oneSheet};
        $wsheet->{nextFree} = 2;
        $model->{inputTables} ||= [];
        $model->{dataset}{1500}[3]{'Company charging year data version'} =
          $model->{version}
          if $model->{version} && $model->{dataset};
        my $te = Dataset(
            number        => 1500,
            dataset       => $model->{dataset},
            name          => 'Company, charging year, data version',
            cols          => Labelset( list => [qw(Company Year Version)] ),
            defaultFormat => 'texthard',
            data          => [
                'Illustrative company',
                'Illustrative year',
                'Illustrative dataset'
            ]
        );
        my ( $sh, $ro, $co ) = $te->wsWrite( $wbook, $wsheet );

        # require Spreadsheet::WriteExcel::Utility;
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
          foreach sort { ( $a->{number} || 9909 ) <=> ( $b->{number} || 9909 ) }
          @{ $model->{inputTables} };
        my $nextFree = delete $wsheet->{nextFree};
        Notes( lines =>
              [ 'Input data', '', 'This sheet contains the input data.' ] )
          ->wsWrite( $wbook, $wsheet );
        $wsheet->{nextFree} = $nextFree;
      }

      ,

      'Volumes' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   42 );
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
            $wsheet->set_column( 0, 0,   42 );
            $wsheet->set_column( 1, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( name => 'Network usage checks' ),
              @{ $model->{checkTables} };
        }
      )
      : ()

      ,

      'Costs' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   42 );
        $wsheet->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Relevant costs and charges' ),
          @{ $model->{costTables} };
      }

      ,

      'Buildup' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   42 );
        $wsheet->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Tariff build-up' ),
          @{ $model->{buildupTables} };
      }

      ,

      'Tariffs' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   42 );
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
        $wsheet->set_column( 0, 0,   42 );
        $wsheet->set_column( 1, 250, 20 );
        my $noLinks = $wbook->{noLinks};
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Revenues' ), @{ $model->{revenueTables} };
        $wbook->{noLinks} = $noLinks;
      }

      ,

      $model->{detailedTables} && @{ $model->{detailedTables} }
      ? (
        'Details' => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->set_column( 0, 0,   48 );
            $wsheet->set_column( 1, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( name => 'Detailed tables' ),
              @{ $model->{detailedTables} };
        }
      )
      : ()

      ,

      $model->{oneSheet} ? ( 'All' => sub { } ) : (
        'Index' => sub {

            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->fit_to_pages( 1, 2 );
            $wsheet->set_column( 0, 0,   30 );
            $wsheet->set_column( 1, 1,   90 );
            $wsheet->set_column( 2, 250, 30 );
            $model->frontPageNotices->wsWrite( $wbook, $wsheet );

            $wsheet->write_string(
                2, 2,
                'Colour coding',
                $wbook->getFormat('thc')
            );
            $wsheet->write_string( 3, 2, 'Data input',
                $wbook->getFormat('0.000hard') );
            $wsheet->write_string(
                4, 2,
                'Unused cell in input data table',
                $wbook->getFormat('unused')
            );
            $wsheet->write_string( 5, 2, 'Calculation',
                $wbook->getFormat('0.000soft') );
            $wsheet->write_string( 6, 2, 'Copy data',
                $wbook->getFormat('0.000copy') );
            $wsheet->write_string(
                7, 2,
                'Unused cell in calculation table',
                $wbook->getFormat('unavailable')
            );
            $wsheet->write_string(
                8, 2,
                'Constant value',
                $wbook->getFormat('0.000con')
            );
            $wsheet->write_string(
                9, 2,
                'Unlocked cell for notes',
                $wbook->getFormat('scribbles')
            );

            $wbook->{logger}->wsWrite( $wbook, $wsheet );

            local $/ = "\n";

            Notes(
                name       => '',
                rowFormats => ['caption'],
                lines      => [
                    'Model identification and configuration',
                    $model->{yaml},
                    '',
                    'Generated on '
                      . POSIX::strftime( '%a %e %b %Y %H:%M:%S',
                        @{ $model->{localTime} } )
                      . ( $ENV{SERVER_NAME} ? " by $ENV{SERVER_NAME}" : '' ),
                ]
            )->wsWrite( $wbook, $wsheet );

        }
      );

}

sub frontPageNotices {
    my ($model) = @_;
    Notes(
        name  => 'Index',
        lines => [
            <<'EOL',

Copyright 2012-2013 Franck Latrémolière, Reckon LLP and others.

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
EOL
            $model->{noLinks} ? () : <<EOL,

This workbook is structured as a series of named and numbered tables. There
is a list of tables below, with hyperlinks.  Above each calculation table,
there is a description of the calculations made, and a hyperlinked list of
the tables or parts of tables from which data are used in the calculation.

Hyperlinks point to the first column heading of the relevant table, or to
the first column heading of the relevant part of the table in the case of
references to a particular set of columns within a composite data table.
Scrolling up or down is usually required after clicking a hyperlink in order
to bring the relevant data and/or headings into view.

Some versions of Microsoft Excel can display a "Back" button, which can be
useful when using hyperlinks to navigate around the workbook.
EOL
            <<EOL,

UNLESS STATED OTHERWISE, THIS WORKBOOK IS ONLY A PROTOTYPE FOR TESTING
PURPOSES AND ALL THE DATA IN THIS MODEL ARE FOR ILLUSTRATION ONLY.
EOL
        ]
    );
}

1;
