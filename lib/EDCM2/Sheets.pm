package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2012 Energy Networks Association Limited and others.
Copyright 2013-2014 Franck Latrémolière, Reckon LLP and others.

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

sub worksheetsAndClosures {

    my ( $model, $wbook ) = @_;

    '11' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber}    = 11;
        $wbook->{lastSheetNumber} = 40;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
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
            $wsheet->set_column( 0,  0,   16 );
            $wsheet->set_column( 1,  1,   40 );
            $wsheet->set_column( 2,  8,   16 );
            $wsheet->set_column( 9,  9,   40 );
            $wsheet->set_column( 10, 250, 16 );
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
            $wsheet->set_landscape;
            my $locationColumn = $model->{dcp189} ? 9 : 8;
            $wsheet->set_column( 0,                   0,                   16 );
            $wsheet->set_column( 1,                   1,                   50 );
            $wsheet->set_column( 2,                   $locationColumn - 1, 20 );
            $wsheet->set_column( $locationColumn,     $locationColumn,     50 );
            $wsheet->set_column( $locationColumn + 1, 250,                 20 );
            $_->wsWrite( $wbook, $wsheet ) foreach $model->{table935};
          }

        ,

        $model->{legacy201} || !$model->{locationTables} ? () : (
            Loc => sub {
                my ($wsheet) = @_;
                $wsheet->{sheetNumber} = 39;
                $wsheet->freeze_panes( 1, 2 );
                $wsheet->set_column( 0, 0,   16 );
                $wsheet->set_column( 1, 1,   50 );
                $wsheet->set_column( 2, 250, 20 );
                $_->wsWrite( $wbook, $wsheet )
                  foreach Notes( lines => 'Preprocessing of location data' ),
                  @{ $model->{locationTables} };
            }
        ),

        ,

        $model->{newOrder}
        ? (
            Calc => sub {
                my ($wsheet) = @_;
                $wsheet->{sheetNumber} = 40;
                $wsheet->{lastTableNumber} =
                  $model->{method} && $model->{method} =~ /LRIC/i ? 0 : -1;
                $wsheet->{tableNumberIncrement} = 2;
                $wsheet->freeze_panes( 1, 1 );
                $wsheet->set_column( 0, 250, 20 );
                $_->wsWrite( $wbook, $wsheet )
                  foreach Notes( lines => 'Calculations' ),
                  @{ $model->{newOrder} };
            },
          )

        :

          (
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

          )

        ,

        'Results' => sub {
            my ($wsheet) = @_;
            $wsheet->{sheetNumber} = 45;
            $wsheet->freeze_panes( 1, 2 );
            $wsheet->set_column( 0, 0,   20 );
            $wsheet->set_column( 1, 1,   50 );
            $wsheet->set_column( 2, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( lines => 'Results' ), @{ $model->{tariffTables} };
          }

        ,

        $model->{summaries} && $model->{summaries} =~ /matri/i
        ?

          (
            'Mat' => sub {
                my ($wsheet) = @_;
                $wsheet->freeze_panes( 1, 2 );
                $wsheet->set_column( 0, 0,   20 );
                $wsheet->set_column( 1, 1,   50 );
                $wsheet->set_column( 2, 250, 20 );
                my ( @matrices, @total, @diff );
                foreach my $col ( 0, 1 ) {
                    my $name =
                      $col
                      ? 'capacity charge p/kVA/day'
                      : 'super-red rate p/kWh';
                    push @{ $model->{matricesData}[$col] },
                      $total[$col] = Arithmetic(
                        name          => 'Total notional ' . $name,
                        defaultFormat => $col ? '0.00soft' : '0.000soft',
                        arithmetic    => '='
                          . join( '+',
                            map { "IV$_" }
                              1 .. @{ $model->{matricesData}[$col] } ),
                        arguments => {
                            map {
                                ( "IV$_" =>
                                      $model->{matricesData}[$col][ $_ - 1 ] );
                            } 1 .. @{ $model->{matricesData}[$col] }
                        },
                      );
                    push @{ $model->{matricesData}[$col] },
                      $diff[$col] = Arithmetic(
                        name          => "Difference $name",
                        arithmetic    => '=IV1-IV2',
                        defaultFormat => $total[$col]{defaultFormat},
                        arguments     => {
                            IV1 => $total[$col],
                            IV2 => $model->{tariffTables}[0]{columns}
                              [ 1 + 2 * $col ],
                        }
                      );
                    push @{ $model->{matricesData}[$col] },
                      Stack( sources => [ $model->{matricesData}[2] ] )
                      unless $col;
                    push @{ $model->{matricesData}[$col] },
                      Arithmetic(
                        name          => 'Consistency check (p/kVA/day)',
                        defaultFormat => '0.00soft',
                        arithmetic    => '=IV1*IV3*IV4/IV5+IV6',
                        arguments     => {
                            IV1 => $diff[0],
                            IV3 => $model->{matricesData}[2],
                            IV4 => $model->{matricesData}[3],
                            IV5 => $model->{matricesData}[4],
                            IV6 => $diff[1],
                        }
                      ) if $col;
                    unshift @{ $model->{matricesData}[$col] },
                      Stack(
                        sources => [ $model->{tariffTables}[0]{columns}[0] ] );
                    push @matrices,
                      Columnset(
                        name    => "Total $name",
                        columns => $model->{matricesData}[$col]
                      );
                }

                $_->wsWrite( $wbook, $wsheet )
                  foreach Notes( lines => 'Matrices and revenue summary' ),
                  @matrices, @{ $model->{revenueTables} };
            },
          )

        :

          (
            'HSummary' => sub {
                my ($wsheet) = @_;
                $wsheet->{sheetNumber} = 46;
                $wsheet->freeze_panes( 1, 2 );
                $wsheet->set_landscape;
                $wsheet->set_column( 0, 0,   20 );
                $wsheet->set_column( 1, 1,   50 );
                $wsheet->set_column( 2, 250, 20 );
                $_->wsWrite( $wbook, $wsheet )
                  foreach Notes( lines => 'Revenue summary' ),
                  @{ $model->{revenueTables} };
            },
          )

        ,

      )

      : (),

      $model->{transparency}
      ? (
        'Aggregates' => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->set_column( 0, 250, 30 );
            my %olTabCol;
            while ( my ( $num, $obj ) =
                each %{ $model->{transparency}{olTabCol} } )
            {
                my $number = int( $num / 100 );
                $olTabCol{$number}[ $num - $number * 100 - 1 ] = $obj;
            }
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( name => 'Aggregates' ), (
                map {
                    Columnset(
                        name => 'Summary aggregate data part ' . ( $_ - 1190 ),
                        number  => 3600 + $_,
                        columns => [
                            map { Stack( sources => [$_] ) } @{ $olTabCol{$_} }
                        ]
                      )
                } sort keys %olTabCol
              ),
              (
                map {
                    my $obj  = $model->{transparency}{olFYI}{$_};
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
                  keys %{ $model->{transparency}{olFYI} }
              );
            $wsheet->{sheetNumber} = 48;
        }
      )
      : $model->{noOneLiners} ? ()
      : (
        'OneLiners' => sub {
            my ($wsheet) = @_;
            $wsheet->{sheetNumber} = 47;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->fit_to_pages( 1, 1 );
            $wsheet->set_column( 0, 250, 30 );
            return unless $wbook->{logger};
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

      $model->{customerTemplates}
      ? (
        ImpT => sub {
            my ($wsheet) = @_;
            $wsheet->fit_to_pages( 1, 1 );
            $wsheet->set_column( 0, 0,   60 );
            $wsheet->set_column( 1, 250, 30 );
            my $logger  = delete $wbook->{logger};
            my $noLinks = $wbook->{noLinks};
            $wbook->{noLinks} = 1;
            splice @{ $model->{tablesTemplateImport} }, 1, 0,
              $wbook->colourCode(1);
            $_->wsWrite( $wbook, $wsheet )
              foreach @{ $model->{tablesTemplateImport} };
            delete $wbook->{noLinks};
            $wbook->{logger} = $logger if $logger;
            $wbook->{noLinks} = $noLinks;
        },
        ExpT => sub {
            my ($wsheet) = @_;
            $wsheet->fit_to_pages( 1, 1 );
            $wsheet->set_column( 0, 0,   60 );
            $wsheet->set_column( 1, 250, 30 );
            my $logger  = delete $wbook->{logger};
            my $noLinks = $wbook->{noLinks};
            $wbook->{noLinks} = 1;
            splice @{ $model->{tablesTemplateExport} }, 1, 0,
              $wbook->colourCode(1);
            $_->wsWrite( $wbook, $wsheet )
              foreach @{ $model->{tablesTemplateExport} };
            delete $wbook->{noLinks};
            $wbook->{logger} = $logger if $logger;
            $wbook->{noLinks} = $noLinks;
        },
        VBACode => sub {
            my ($wsheet) = @_;
            $model->vbaWrite( $wbook, $wsheet );
        },
      )
      : ()

      ,

      $model->{ldnoRev}
      ?

      (

        'LDNORev' => sub {
            my ($wsheet) = @_;
            $wsheet->{sheetNumber} = 60;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->set_column( 0, 0,   50 );
            $wsheet->set_column( 1, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach grep { $_; } @{ $model->{ldnoRevTables} };
          }

      )

      : ()

      ,

      $model->{TotalsTables}
      ? (
        'Total' => sub {
            my ($wsheet) = @_;
            $wsheet->{sheetNumber} = 61;
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
        $wsheet->set_column( 1, 1,   105 );
        $wsheet->set_column( 2, 250, 30 );
        $wbook->{logger}{showColumns} = 1 if $model->{newOrder};
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->topNotes, $model->licenceNotes, $wbook->colourCode,
          $wbook->{logger}, $model->technicalNotes;
      };

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
        name => $model->{legacy201} ? 'Overview' : 'Index',
        lines => [
            $model->{colour} && $model->{colour} =~ /orange|gold/ ? <<EOL : (),

This document, model or dataset has been prepared by Reckon LLP on the instructions of the DCUSA Panel or one of its working
groups.  Only the DCUSA Panel and its working groups have authority to approve this material as meeting their requirements. 
Reckon LLP makes no representation about the suitability of this material for the purposes of complying with any licence
conditions or furthering any relevant objective.
EOL
            $model->{colour} && $model->{colour} =~ /gold/ ? <<EOL :

UNLESS STATED OTHERWISE, ALL THE DATA IN THIS MODEL ARE FOR ILLUSTRATION ONLY.
EOL
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
Copyright 2009-2012 Energy Networks Association Limited and others.  Copyright 2013-2014 Franck Latrémolière, Reckon LLP and others.
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
