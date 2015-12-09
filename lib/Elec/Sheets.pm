package Elec;

=head Copyright licence and disclaimer

Copyright 2012-2015 Franck Latrémolière, Reckon LLP and others.

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
require Spreadsheet::WriteExcel::Utility;
require SpreadsheetModel::FormatLegend;

sub worksheetsAndClosures {

    my ( $model, $wbook ) = @_;

    my @detailedTables;
    push @detailedTables, @{ $model->{detailedTables} }
      if $model->{detailedTables};
    push @detailedTables, @{ $model->{detailedTables2} }
      if $model->{detailedTables2};

    'Input' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber} = 15;
        $wbook->{lastSheetNumber} = $model->{table1653} ? 48 : 49;
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
            $wsheet->set_column( 0, 0, $model->{ulist} ? 50 : 20 );
            $wsheet->set_column( 1, 1, $model->{ulist} ? 20 : 50 );
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
            $wsheet->set_column( 0, 0,   $model->{ulist} ? 50 : 20 );
            $wsheet->set_column( 1, 250, $model->{ulist} ? 20 : 50 );
            $wsheet->set_column( 2, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( name => 'Detailed tables' ),
              @detailedTables;
        }
      )
      : ()

      ,

      'Index' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_print_scale(50);
        $wsheet->set_column( 0, 0,   16 );
        $wsheet->set_column( 1, 1,   112 );
        $wsheet->set_column( 2, 250, 32 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->topNotes, $model->licenceNotes,
          SpreadsheetModel::FormatLegend->new,
          $wbook->{logger}, $model->technicalNotes;
      }

      ;

}

sub technicalNotes {
    my ($model) = @_;
    require POSIX;
    Notes(
        name       => '',
        rowFormats => ['caption'],
        lines      => [
            'Technical model rules and version control',
            $model->{yaml},
            '',
            'Generated on '
              . POSIX::strftime( '%a %e %b %Y %H:%M:%S',
                $model->{localTime} ? @{ $model->{localTime} } : localtime )
              . ( $ENV{SERVER_NAME} ? " by $ENV{SERVER_NAME}" : '' ),
        ]
    );
}

sub topNotes {
    my ($model) = @_;
    Notes(
        name  => 'Index',
        lines => [
            <<EOL,

UNLESS STATED OTHERWISE, THIS WORKBOOK IS ONLY A PROTOTYPE FOR TESTING PURPOSES AND ALL THE DATA IN THIS MODEL ARE FOR ILLUSTRATION ONLY.
EOL
            $model->{noLinks} ? () : <<EOL,

This workbook is structured as a series of named and numbered tables. There is a list of tables below, with hyperlinks.  Above
each calculation table, there is a description of the calculations and hyperlinks to tables from which data are used. Hyperlinks
point to the relevant table column heading of the relevant table. Scrolling up or down is usually required after clicking a
hyperlink in order to bring the relevant data and/or headings into view. Some versions of Microsoft Excel can display a "Back"
button, which can be useful when using hyperlinks to navigate around the workbook.
EOL
        ]
    );
}

sub licenceNotes {
    Notes(
        name  => '',
        lines => <<'EOL',
Copyright 2012-2014 Franck Latrémolière, Reckon LLP and others.
The code used to generate this spreadsheet includes open-source software published at https://github.com/f20/power-models.
Use and distribution of the source code is subject to the conditions stated therein. 
Any redistribution of this software must retain the following disclaimer:
THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED
TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL AUTHORS OR
CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED
AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN
ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
EOL
    );
}

1;
