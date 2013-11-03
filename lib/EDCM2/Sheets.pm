package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2013 Energy Networks Association Limited and others.

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

=head Table numbers used in this file

1100

=cut

use warnings;
use strict;
use utf8;
require Spreadsheet::WriteExcel::Utility;
use SpreadsheetModel::Shortcuts ':all';

sub generalNotes {
    my ($model) = @_;
    Notes(
        name => $model->{legacy201} ? 'Overview' : 'Index',
        lines => [
            $model->{colour} && $model->{colour} =~ /orange/ ? <<EOL : (),

This document, model or dataset has been prepared by Reckon LLP on the instructions of the DCUSA Panel or
one of its working groups.  Only the DCUSA Panel and its working groups have authority to approve this
material as meeting their requirements.  Reckon LLP makes no representation about the suitability of this
material for the purposes of complying with any licence conditions or furthering any relevant objective.
EOL
            <<EOL,

UNLESS STATED OTHERWISE, THIS WORKBOOK IS ONLY A PROTOTYPE FOR TESTING
PURPOSES AND ALL THE DATA IN THIS MODEL ARE FOR ILLUSTRATION ONLY.
EOL
            $model->{noLinks} ? () : <<EOL,

This workbook is structured as a series of named and numbered tables. There is a list of tables below, with
hyperlinks.  Above each calculation table, there is a description of the calculations made, and a hyperlinked
list of the tables or parts of tables from which data are used in the calculation. Hyperlinks point to the
relevant table column heading of the relevant table. Scrolling up or down is usually required after clicking
a hyperlink in order to bring the relevant data and/or headings into view. Some versions of Microsoft Excel
can display a "Back" button, which can be useful when using hyperlinks to navigate around the workbook.
EOL
            <<'EOL',

Copyright 2009-2012 Energy Networks Association Limited and others. Copyright 2013 Franck Latrémolière,
Reckon LLP and others. Redistribution and use in source and binary forms, with or without modification, are
permitted provided that the following conditions are met:
1. Redistributions of source code must retain the above copyright notice, this list of conditions and the
following disclaimer.
2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the
following disclaimer in the documentation and/or other materials provided with the distribution.
THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES,
INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
ARE DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS
OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY
WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
EOL
        ]
    );
}

sub worksheetsAndClosures {

    my ( $model, $wbook ) = @_;

    $wbook->{lastSheetNumber} = 40;

    '11' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber} = 11;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   35 );
        $wsheet->set_column( 1, 250, 20 );
        $wsheet->{nextFree} = 2;
        $model->{inputTables} ||= [];
        my $te = Dataset(
            number        => 1100,
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
        $sh = $sh->get_name;
        require Spreadsheet::WriteExcel::Utility;
        $wbook->{titleAppend} =
            qq%" for "&'$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co )
          . qq%&" in "&'$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co + 1 )
          . qq%&" ("&'$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co + 2 )
          . '&")"';
        $_->wsWrite( $wbook, $wsheet )
          foreach !$model->{ldnoRev} || $model->{ldnoRev} !~ /only/i
          ? (
            $model->{method} eq 'none' ? () : Notes(
                lines    => 'Power flow input data',
                location => $model->{method} =~ /LRIC/i
                ? 913
                : 911,
            ),
            Notes( lines => 'Tariff input data', location => 935, )
          )
          : (),
          sort { ( $a->{number} || 9909 ) <=> ( $b->{number} || 9909 ) }
          @{ $model->{inputTables} };
        my $nextFree = delete $wsheet->{nextFree};
        Notes( lines => 'General input data' )->wsWrite( $wbook, $wsheet );
        $wsheet->{nextFree} = $nextFree;
      }

      ,

      $model->{impactInputTables}
      ? (
        Impact => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->set_column( 0, 0,   16 );
            $wsheet->set_column( 1, 10,  40 );
            $wsheet->set_column( 2, 250, 16 );
            my $noLinks = delete $wbook->{noLinks};
            $wbook->{noLinks} = 1;
            $_->wsWrite( $wbook, $wsheet )
              foreach @{ $model->{impactInputTables} };
            $wbook->{noLinks} = $noLinks;
        }
      )
      : (),

      ,

      !$model->{ldnoRev} || $model->{ldnoRev} !~ /only/i
      ? (

        $model->{method} eq 'none' ? () : (
            (
                $model->{method} =~ /LRIC/i
                ? 913
                : 911
            ) => sub {
                my ($wsheet) = @_;
                $wsheet->{sheetNumber} = 9;
                $wsheet->freeze_panes( 1, 2 );
                $wsheet->set_column( 0, 0,   20 );
                $wsheet->set_column( 1, 1,   35 );
                $wsheet->set_column( 2, 2,   20 );
                $wsheet->set_column( 3, 3,   35 );
                $wsheet->set_column( 4, 250, 20 );
                $_->wsWrite( $wbook, $wsheet ) foreach $model->{table911};
            }
          )

        ,

        '935' => sub {
            my ($wsheet) = @_;
            $wsheet->{sheetNumber} = 9;
            $wsheet->freeze_panes( 1, 2 );
            $wsheet->set_column( 0, 0,   16 );
            $wsheet->set_column( 1, 1,   50 );
            $wsheet->set_column( 2, 7,   20 );
            $wsheet->set_column( 8, 8,   50 );
            $wsheet->set_column( 9, 250, 20 );
            $_->wsWrite( $wbook, $wsheet ) foreach $model->{table935};
          }

        ,

        'Calc1' => sub {
            my ($wsheet) = @_;
            $wsheet->{lastTableNumber} =
              $model->{method} && $model->{method} =~ /LRIC/i ? 0 : -1;
            $wsheet->{tableNumberIncrement} = 2;
            $wsheet->freeze_panes( 1, 1 );
            $wsheet->set_column( 0, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( lines => 'Calculations part 1' ),
              @{ $model->{calc1Tables} };
          }

        ,

        'Calc2' => sub {
            my ($wsheet) = @_;
            $wsheet->{lastTableNumber} =
              $model->{method} && $model->{method} =~ /LRIC/i ? 0 : -1;
            $wsheet->{tableNumberIncrement} = 2;
            $wsheet->freeze_panes( 1, 1 );
            $wsheet->set_column( 0, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( lines => 'Calculations part 2' ),
              @{ $model->{calc2Tables} };
          }

        ,

        'Calc3' => sub {
            my ($wsheet) = @_;
            $wsheet->{lastTableNumber} =
              $model->{method} && $model->{method} =~ /LRIC/i ? 0 : -1;
            $wsheet->{tableNumberIncrement} = 2;
            $wsheet->freeze_panes( 1, 1 );
            $wsheet->set_column( 0, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( lines => 'Calculations part 3' ),
              @{ $model->{calc3Tables} };
          }

        ,

        'Calc4' => sub {
            my ($wsheet) = @_;
            $wsheet->{lastTableNumber} =
              $model->{method} && $model->{method} =~ /LRIC/i ? 0 : -1;
            $wsheet->{tableNumberIncrement} = 2;
            $wsheet->freeze_panes( 1, 1 );
            $wsheet->set_column( 0, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( lines => 'Calculations part 4' ),
              @{ $model->{calc4Tables} };
          }

        ,

        'Results' => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 2 );
            $wsheet->set_column( 0, 0,   20 );
            $wsheet->set_column( 1, 1,   50 );
            $wsheet->set_column( 2, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( lines => 'Results' ), @{ $model->{tariffTables} };
          }

        ,
        'HSummary' => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 2 );
            $wsheet->set_column( 0, 0,   20 );
            $wsheet->set_column( 1, 1,   50 );
            $wsheet->set_column( 2, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( lines => 'Results' ), @{ $model->{revenueTables} };
          }

        ,
      )

      : (),

      $model->{transparency}
      ? (
        'OneLiners' => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->fit_to_pages( 1, 1 );
            $wsheet->set_column( 0, 250, 30 );
            my %olo;
            while ( my ( $num, $obj ) = each %{ $model->{transparency}{olo} } )
            {
                my $number = int( $num / 100 );
                $olo{$number}[ $num - $number * 100 - 1 ] = $obj;
            }
            $_->wsWrite( $wbook, $wsheet ) foreach Notes(
                lines => 'Updated baseline data

This sheet contains data to populate tables 1191 to 1194 in a slave model.'
              ),
              (
                map {
                    Columnset(
                        name   => "Data for table $_",
                        number => 3600 + $_,
                        columns =>
                          [ map { Stack( sources => [$_] ) } @{ $olo{$_} } ]
                      )
                } sort keys %olo
              ),
              (
                map {
                    my $obj  = $model->{transparency}{oli}{$_};
                    my $name = 'Copy of ' . $obj->{name};
                    $obj->isa('SpreadsheetModel::Columnset')
                      ? Columnset(
                        name    => $name,
                        number  => 3600 + $_,
                        columns => [
                            map { Stack( sources => [$_] ) }
                              @{ $obj->{columns} }
                        ]
                      )
                      : Stack(
                        name    => $name,
                        number  => 3600 + $_,
                        sources => [$obj]
                      );
                  } sort { $a <=> $b }
                  keys %{ $model->{transparency}{oli} }
              );
            $wbook->{lastSheetNumber} = $wsheet->{sheetNumber} = 48;
        }
      )
      : $model->{noOneLiners} ? ()
      : (
        'OneLiners' => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->fit_to_pages( 1, 1 );
            $wsheet->set_column( 0, 250, 30 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes(
                lines => 'Copy of all single-line tables in the model' ), map {
                $_->isa('SpreadsheetModel::Columnset')
                  ? Columnset(
                    name => "Copy of $_->{name}",
                    columns =>
                      [ map { Stack( sources => [$_] ) } @{ $_->{columns} } ]
                  )
                  : Stack(
                    name    => "Copy of $_->{name}",
                    sources => [$_]
                  );
                } grep { $_->lastRow == 0 } @{ $wbook->{logger}{objects} };
        }
      )

      ,

      $model->{ldnoRev}
      ?

      (

        'LDNORev' => sub {
            my ($wsheet) = @_;
            $wbook->{lastSheetNumber} = $wsheet->{sheetNumber} = 60;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->set_column( 0, 0,   50 );
            $wsheet->set_column( 1, 250, 20 );
            $_->wsWrite( $wbook, $wsheet ) foreach @{ $model->{ldnoRevTables} };
          }

      )

      : ()

      ,

      $model->{TotalsTables}
      ? (
        'Total' => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 2 );
            $wsheet->set_column( 0, 0,   20 );
            $wsheet->set_column( 1, 1,   50 );
            $wsheet->set_column( 2, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( lines => 'Total' ), @{ $model->{TotalsTables} };
        }
      )
      : ()

      ,

      $model->{legacy201} ? 'Overview' : 'Index' => sub {

        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->fit_to_pages( 1, 2 );
        $wsheet->set_column( 0, 0,   30 );
        $wsheet->set_column( 1, 1,   90 );
        $wsheet->set_column( 2, 250, 30 );

        $model->generalNotes->wsWrite( $wbook, $wsheet );

        $wsheet->write_string( 2, 2, 'Colour coding',
            $wbook->getFormat('thc') );
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

      };

}

1;
