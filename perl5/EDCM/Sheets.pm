package EDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and contributors. All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted
provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of
conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of
conditions and the following disclaimer in the documentation and/or other materials provided
with the distribution.

THIS SOFTWARE IS PROVIDED BY ENERGY NETWORKS ASSOCIATION LIMITED AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL ENERGY
NETWORKS ASSOCIATION LIMITED OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF
THE POSSIBILITY OF SUCH DAMAGE.

=cut

=head Table numbers used in this file

1100

=cut

use warnings;
use strict;
use utf8;
require Spreadsheet::WriteExcel::Utility;
use SpreadsheetModel::Shortcuts ':all';

sub worksheetsAndClosures {

    my ( $model, $wbook ) = @_;

    $wbook->{lastSheetNumber} = 40;

    '11' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber} = 11;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   35 );
        $wsheet->set_column( 1, 250, 20 );
        $wsheet->{nextFree} = 3;
        $model->{inputTables} ||= [];
        my $te = Dataset(
            number        => 1100,
            appendTo      => $model->{inputTables},
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
              . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro,
                $co + 1 )
              . qq%&" ("&'$sh'!%
              . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro,
                $co + 2 )
              . '&")"';
        $_->wsWrite( $wbook, $wsheet )
          foreach sort { ( $a->{number} || 9909 ) <=> ( $b->{number} || 9909 ) }
          @{ $model->{inputTables} };
        delete $wsheet->{nextFree};
        $model->notes11->wsWrite( $wbook, $wsheet );
      }

      ,

      !$model->{ldnoRev} || $model->{ldnoRev} !~ /only/i
      ? (

        $model->{method} eq 'none' ? () : (
            (
                $model->{method} =~ /LRIC/i
                ? ( $model->{method} =~ /split/i ? '913' : '912' )
                : '911'
            ) => sub {
                my ($wsheet) = @_;
                $wsheet->{sheetNumber} = 9;
                $wsheet->freeze_panes( 1, 2 );
                $wsheet->set_column( 0, 0,   20 );
                $wsheet->set_column( 1, 1,   35 );
                $wsheet->set_column( 2, 2,   20 );
                $wsheet->set_column( 3, 3,   35 );
                $wsheet->set_column( 4, 250, 20 );
                $_->wsWrite( $wbook, $wsheet )
                  foreach $model->notes911, $model->{table911};
            }
          )

        ,

        $model->{table935}
        ?

          (
            '935' => sub {
                my ($wsheet) = @_;
                $wsheet->{sheetNumber} = 9;
                $wsheet->freeze_panes( 1, 3 );
                $wsheet->set_column( 0, 0,   16 );
                $wsheet->set_column( 1, 1,   50 );
                $wsheet->set_column( 2, 3,   16 );
                $wsheet->set_column( 4, 250, 20 );
                $_->wsWrite( $wbook, $wsheet )
                  foreach $model->notes935, $model->{table935};
            }
          )

        :

          (
            '953' => sub {
                my ($wsheet) = @_;
                $wsheet->{sheetNumber} = 9;
                $wsheet->freeze_panes( 1, 3 );
                $wsheet->set_column( 0, 1, 16 );
                $wsheet->set_column( 2, 2, 50 );
                0 and $wsheet->set_column( 3, 5, 20 );
                0 and $wsheet->set_column( 6, 7, 35 );
                $wsheet->set_column( 3, 250, 20 );
                $_->wsWrite( $wbook, $wsheet )
                  foreach $model->notes953, $model->{table953};
            }
          )

        ,

        $model->{dbSheet}
        ? (
            'db' => sub {
                my ($wsheet) = @_;
                $wsheet->set_column( 0, 250, 20 );
                my $nextCol = 0;
                $wbook->{ 0 + $model->{tariffSet} } = sub {
                    my @here = ( $wsheet, 0, $nextCol );
                    $nextCol += $_[0] + 1;
                    @here;
                };
            }
          )
        : ()

        ,

        1 ? () : (
            'Com' => sub {
                my ($wsheet) = @_;
                $wsheet->freeze_panes( 1, 0 );
                $wsheet->set_column( 0, 250, 20 );
                $_->wsWrite( $wbook, $wsheet )
                  foreach Notes( lines => 'Com' ), @{ $model->{tablesA} };
            }
          )

        ,

        1 ? () : (
            'Loc' => sub {
                my ($wsheet) = @_;
                $wsheet->freeze_panes( 1, 3 );
                $wsheet->set_column( 0, 250, 20 );
                $_->wsWrite( $wbook, $wsheet )
                  foreach Notes( lines => 'Locations' ),
                  Columnset(
                    name    => 'Location pre-processing',
                    columns => $model->{locationData}
                  );
            }
          )

        ,

        $model->{noGen} ? () : (
            'Gen' => sub {
                my ($wsheet) = @_;
                $wsheet->freeze_panes( 1, 0 );
                $wsheet->set_column( 0, 250, 20 );
                $_->wsWrite( $wbook, $wsheet )
                  foreach Notes( lines => 'Generation' ),
                  @{ $model->{tablesG} };
            }
          )

        ,

        'Dem' => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->set_column( 0, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( lines => 'Demand' ), @{ $model->{tablesD} };
          }

        ,

        'Results' => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 3 );
            $wsheet->set_column( 0, 1,   20 );
            $wsheet->set_column( 2, 2,   50 );
            $wsheet->set_column( 3, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( lines => 'Results' ), @{ $model->{tariffTables} },
              @{ $model->{revenueTables} };
          }

        ,

        $model->{summaries} && !$model->{noGen}
        ? (

            'Y' => sub {
                my ($wsheet) = @_;

                $wsheet->freeze_panes( 1, 3 );
                $wsheet->set_column( 0, 1,   20 );
                $wsheet->set_column( 2, 2,   50 );
                $wsheet->set_column( 3, 250, 20 );
                $_->wsWrite( $wbook, $wsheet ) foreach Notes(
                    lines => [
                        'This sheet is not part of the calculation of charges.'
                    ]
                  ),
                  @{ $model->{summaryTables} }, @{ $model->{volatilityTables} };
              }

            ,
          )
        : (),

        $model->{summaries} && $model->{summaries} =~ /wrong/
        ? (

            'Mess' => sub {
                my ($wsheet) = @_;

                #            $wsheet->{sheetNumber} = 49;
                $wsheet->freeze_panes( 1, 3 );
                $wsheet->set_column( 0, 1,   20 );
                $wsheet->set_column( 2, 2,   35 );
                $wsheet->set_column( 3, 250, 20 );
                $_->wsWrite( $wbook, $wsheet )
                  foreach Notes( name => 'Internal workings' ),
                  @{ $model->{matrixTables} };
              }

            ,

            'Split' => sub {
                my ($wsheet) = @_;

                #            $wsheet->{sheetNumber} = 49;
                $wsheet->freeze_panes( 1, 3 );
                $wsheet->set_column( 0, 1,   20 );
                $wsheet->set_column( 2, 2,   35 );
                $wsheet->set_column( 3, 250, 20 );
                $_->wsWrite( $wbook, $wsheet )
                  foreach Notes( name => 'Splits' ), @{ $model->{splitTables} };
              }

            ,

            1 ? () : (
                'Explain' => sub {
                    my ($wsheet) = @_;
                    $wsheet->freeze_panes( 1, 0 );
                    $wsheet->set_column( 0, 0,   32 );
                    $wsheet->set_column( 1, 250, 24 );
                    my $logger = $wbook->{logger};
                    delete $wbook->{logger};
                    my $modulo = $wbook->{modulo};
                    delete $wbook->{modulo};
                    my $noLinks = $wbook->{noLinks};
                    $wbook->{noLinks} = 1;
                    $_->wsWrite( $wbook, $wsheet )
                      foreach Notes( lines => 'Individual explanation' ),
                      @{ $model->{explainTables} };
                    delete $wbook->{noLinks};
                    $wbook->{logger} = $logger if $logger;
                    $wbook->{modulo} = $modulo if $modulo;
                    $wbook->{noLinks} = $noLinks;
                }
            ),

          )

        :

          (),

      )

      : (),

      $model->{ldnoRev}
      ?

      (

        'LDNORev' => sub {
            my ($wsheet) = @_;
            $wsheet->{sheetNumber} = 60;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->set_column( 0, 0,   50 );
            $wsheet->set_column( 1, 250, 20 );
            $_->wsWrite( $wbook, $wsheet ) foreach @{ $model->{ldnoRevTables} };
          }

      )

      : ()

      ,

      'Overview' => sub {

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

sub notes11 {
    my ($model) = @_;
    Notes( lines => 'General input data' );
}

sub notes911 {
    my ($model) = @_;
    Notes(
        lines => [
            'Power flow input data',
            '', "This model has space for $model->{numLocations} locations.",
        ]
    );
}

sub notes953 {
    my ($model) = @_;
    Notes(
        lines => [
            'Tariff input data',
            '', "This model has space for $model->{numTariffs} tariffs.",
        ]
    );
}

sub notes935 {
    my ($model) = @_;
    Notes(
        lines => [
            'Tariff input data',
            '', "This model has space for $model->{numTariffs} tariff pairs.",
        ]
    );
}

sub notesCharge12 {
    my ($model) = @_;
    Notes( lines => 'Charge 1/2/exit' );
}

sub notesPreprocessing {
    my ($model) = @_;
    Notes( lines => 'Pre-processing of power flow data' );
}

sub notesScaling {
    my ($model) = @_;
    Notes(
        lines => $model->{adder} eq '3'
        ? 'Scaling (type A)'
        : 'Sole use assets'
    );
}

sub notesSoleUse {
    Notes( lines => 'Sole use assets' );
}

sub notesAssets {
    Notes( lines => 'Notional assets' );
}

sub notesAlloc {
    Notes( lines => 'Allocations' );
}

sub notesUse {
    Notes( lines => 'Network use' );
}

sub notesTariffs {
    my ($model) = @_;
    Notes(
        lines => $model->{adder} eq '3'
        ? 'Tariffs'
        : 'Tariffs before matching'
    );
}

sub notesTariffFinal {
    my ($model) = @_;
    Notes( lines => 'Tariffs' );
}

sub notesRevenue {
    my ($model) = @_;
    Notes(
        lines => $model->{adder} eq '3'
        ? 'Revenue'
        : 'Revenue before matching'
    );
}

sub notesMatching {
    my ($model) = @_;
    Notes( lines => 'Scaling (type B)' );
}

sub generalNotes {
    my ($model) = @_;
    Notes(
        name  => 'Overview',
        lines => [
            $model->{noLinks} ? () : <<'EOL',

Copyright 2009-2011 Energy Networks Association Limited and contributors. All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted
provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of
conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of
conditions and the following disclaimer in the documentation and/or other materials provided
with the distribution.

THIS SOFTWARE IS PROVIDED BY ENERGY NETWORKS ASSOCIATION LIMITED AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL ENERGY
NETWORKS ASSOCIATION LIMITED OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF
THE POSSIBILITY OF SUCH DAMAGE.
EOL
        ]
    );
}

sub notesSumV {
    my ($model) = @_;
    Notes( lines => 'Summary for volatility' );
}

sub notesMatrix {
    my ($model) = @_;
    Notes( lines => 'Matrix' );
}

1;
