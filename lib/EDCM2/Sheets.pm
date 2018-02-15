package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2012 Energy Networks Association Limited and others.
Copyright 2013-2017 Franck Latrémolière, Reckon LLP and others.

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
    $_->finishModel($wbook)
      foreach grep { $_; } @{$model}{qw(embeddedCdcm embeddedModelM)};
}

sub notesTransparency {
    my ($model) = @_;
    Notes(
        name  => 'DNO totals data',
        lines => [
            'If table 1190 is set to "TRUE",'
              . ' then tables 1191-1193 are not used. '
              . 'In this configuration, the model only works when populated'
              . ' with data for all customers, including'
              . ' data that are commercially confidential.',
            'If table 1190 is set to "FALSE", so that the model can be used for'
              . ' validation and forecasting of DNO charges, then'
              . ' the DNO totals in tables 1191-1193 need to be mirror'
              . ' thee totals in the DNO\'s charging model.',
            'There is a growing trend among DNOs of refusing to disclose'
              . ' these DNO totals, giving some version of a'
              . ' manifestly invalid excuse which sounds like it might'
              . ' have been cooked up in a DNO smoked-filled room. ',
            1 ? ()
            : 'Previous attempts at seeking assistance from Ofgem have failed. ',
            $model->{mitigateUndueSecrecy}
            ? $model->{mitigateUndueSecrecy}->additionalLines
            : (     'There are tools on dcmf.co.uk/models which might'
                  . ' in some cases help mitigate undue DNO secrecy.' ),
        ],
    );
}

sub sheetPriority {
    my ( $model, $sheet ) = @_;
    my $score = 0;
    $score = 15 if $sheet =~ /Volumes\$/;
    $score = 12 if $sheet =~ /\$$/;
    $score ||= 6 if $sheet =~ /(?:Overview|Index)$/is;
    $score;
}

sub worksheetsAndClosures {

    my ( $model, $wbook ) = @_;

    my @wsheetsAndClosures;

    push @wsheetsAndClosures, '11' => sub {

        my ($wsheet) = @_;
        $wsheet->{sheetNumber} = 11;
        $wbook->{lastSheetNumber} =
             $model->{layout} && $model->{layout} =~ /matrix/
          || $model->{impactInputTables}
          || $model->{embeddedCdcm} ? 19 : 40;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );
        $model->{titleWrites}{$wbook} = [];
        $wbook->{titleWriter} =
          sub { push @{ $model->{titleWrites}{$wbook} }, [@_]; };
        $model->{inputTables} ||= [];

        my $idTable = Dataset(
            number        => 1100,
            dataset       => $model->{dataset},
            name          => 'Company, charging year, data version',
            cols          => Labelset( list => [qw(Company Year Version)] ),
            defaultFormat => 'puretexthard',
            data          => [ 'no company', 'no year', 'no data version' ],
            usePlaceholderData => 1,
            forwardLinks       => {},
        );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'General input data' ), $idTable;

        foreach (
            sort { ( $a->{number} || 999_999 ) <=> ( $b->{number} || 999_999 ) }
            @{ $model->{inputTables} }
          )
        {
            map { $_->wsWrite( $wbook, $wsheet ); } $model->notesTransparency
              if $_->{number} && $_->{number} == 1190;
            $_->wsWrite( $wbook, $wsheet );
        }

        require Spreadsheet::WriteExcel::Utility;
        my ( $sh, $ro, $co ) = $idTable->wsWrite( $wbook, $wsheet );
        $sh = $sh->get_name;
        $model->{idAppend}{$wbook} =
            qq%&" for "&'$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co )
          . qq%&" in "&'$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co + 1 )
          . qq%&" ("&'$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co + 2 )
          . '&")"';

    };

    push @wsheetsAndClosures, 'Volumes$' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );
        $wsheet->{sheetNumber} = 14;
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Volumes' ), @{ $model->{volumeTables} };
      }
      if $model->{volumeTables};

    push @wsheetsAndClosures,
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
        $_->wsWrite( $wbook, $wsheet ) foreach @{ $model->{impactInputTables} };
        $wbook->{noLinks} = $noLinks;
      }
      if $model->{impactInputTables};

    push @wsheetsAndClosures,

      $model->{method} eq 'none' ? () : (
        (
            $model->{method} =~ /LRIC/i
            ? 913
            : 911
        ) => sub {
            my ($wsheet) = @_;
            $wsheet->{sheetNumber} = 9;
            $wsheet->freeze_panes( 5, 2 );
            $wsheet->set_column( 0, 0,   20 );
            $wsheet->set_column( 1, 1,   35 );
            $wsheet->set_column( 2, 2,   20 );
            $wsheet->set_column( 3, 3,   35 );
            $wsheet->set_column( 4, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( name => 'Power flow input data' ),
              $model->{table911913};
        }
      )

      ,

      '935' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber} = 9;
        $wsheet->freeze_panes(
            $model->{table935}{sourceLines}
            ? 6 + @{ $model->{table935}{sourceLines} }
            : 6,
            2
        );
        $wsheet->set_landscape;
        my $locationColumn = $model->{dcp189} ? 9 : 8;
        $wsheet->set_column( 0,                   0,                   16 );
        $wsheet->set_column( 1,                   1,                   50 );
        $wsheet->set_column( 2,                   $locationColumn - 1, 20 );
        $wsheet->set_column( $locationColumn,     $locationColumn,     50 );
        $wsheet->set_column( $locationColumn + 1, 250,                 20 );

        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Tariff input data' ), $model->{table935};

        if ( $model->{tariff1Row} ) {
            $wsheet->set_row( $model->{tariff1Row} - 2, 22 );
        }

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
      )

      if !$model->{ldnoRev} || $model->{ldnoRev} !~ /only/i;

    if ( $model->{embeddedModelM}
        && UNIVERSAL::can( $model->{embeddedModelM}, 'worksheetsAndClosures' ) )
    {
        my @mwac = $model->{embeddedModelM}->worksheetsAndClosures($wbook);
        while (@mwac) {
            my $sheet   = shift @mwac;
            my $closure = shift @mwac;
            next if $sheet =~ /^(?:Index|Result)/;
            next
              if $sheet =~ /^Input/
              && !$model->{embeddedModelM}{noSingleInputSheet};
            push @wsheetsAndClosures, "M($sheet)", $closure;
        }
    }

    if ( $model->{embeddedCdcm}
        && UNIVERSAL::can( $model->{embeddedCdcm}, 'worksheetsAndClosures' ) )
    {
        my @cwac = $model->{embeddedCdcm}->worksheetsAndClosures($wbook);
        while (@cwac) {
            my $sheet   = shift @cwac;
            my $closure = shift @cwac;
            next if $sheet =~ /^(?:Index|Result)/;
            next
              if $sheet =~ /^Input/
              && !$model->{embeddedCdcm}{noSingleInputSheet};
            push @wsheetsAndClosures, "C($sheet)", $closure;
        }
    }

    push @wsheetsAndClosures,

      $model->{layout} && $model->{layout} =~ /matrix/i
      ? (
        'Parameters' => sub {
            my ($wsheet) = @_;
            $wsheet->{firstTableNumber} =
              $model->{method} && $model->{method} =~ /LRIC/i ? 2 : 1;
            $wsheet->{tableNumberIncrement} = 2;
            $wsheet->freeze_panes( 1, 1 );
            $wsheet->set_column( 0, 250, 28 );
            $wsheet->set_landscape
              if $model->{layout} && $model->{layout} =~ /wide/i;
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( lines => 'Constant parameters' ),
              @{ $model->{generalTables} };
        },
        'DNO totals' => sub {
            my ($wsheet) = @_;
            $wsheet->{firstTableNumber} =
              $model->{method} && $model->{method} =~ /LRIC/i ? 2 : 1;
            $wsheet->{tableNumberIncrement} = 2;
            $wsheet->freeze_panes( 1, 1 );
            $wsheet->set_column( 0, 250, 28 );
            $wsheet->set_landscape
              if $model->{layout} && $model->{layout} =~ /wide/i;
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( lines => 'DNO-wide aggregates' );
        },
        'Charging rates' => sub {
            my ($wsheet) = @_;
            $wsheet->{firstTableNumber} =
              $model->{method} && $model->{method} =~ /LRIC/i ? 2 : 1;
            $wsheet->freeze_panes( 1, 1 );
            $wsheet->set_column( 0, 250, 28 );
            $wsheet->set_landscape
              if $model->{layout} && $model->{layout} =~ /wide/i;
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( lines => 'Charging rates' );
        },
        'Matrix' => sub {
            my ($wsheet) = @_;
            $wsheet->{firstTableNumber} =
              $model->{method} && $model->{method} =~ /LRIC/i ? 2 : 1;
            $wsheet->{tableNumberIncrement} = 2;
            $_->{tableNumberIncrement}      = 2
              foreach grep { $_ } @{$wbook}{ 'DNO totals', 'Charging rates' };
            $wsheet->freeze_panes( $model->{tariff1Row} || 1, 2 );
            $wsheet->set_landscape;
            $wsheet->set_column( 0, 0,   16 );
            $wsheet->set_column( 1, 1,   50 );
            $wsheet->set_column( 2, 250, 20 );

            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( lines => 'Matrix' ),
              @{ $model->{matrixTables} };

            if ( my $ws = $wbook->{'DNO totals'} ) {
                $_->wsWrite( $wbook, $ws ) foreach @{ $model->{tableList} };
            }
        },
      )

      : ref $model->{tableList} eq 'ARRAY' ? (
        Calc => sub {
            my ($wsheet) = @_;
            $wsheet->{firstTableNumber} =
              $model->{method} && $model->{method} =~ /LRIC/i ? 2 : 1;
            $wsheet->{tableNumberIncrement} = 2;
            $wsheet->freeze_panes( 1, 1 );
            $wsheet->set_column( 0, 250, 20 );
            $wsheet->set_landscape
              if $model->{layout} && $model->{layout} =~ /wide/i;
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( lines => 'Calculations' ),
              @{ $model->{generalTables} }, @{ $model->{tableList} };
        },
      )

      : (
        'Calc1' => sub {
            my ($wsheet) = @_;
            $wsheet->{firstTableNumber} =
              $model->{method} && $model->{method} =~ /LRIC/i ? 2 : 1;
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
            $wsheet->{firstTableNumber} =
              $model->{method} && $model->{method} =~ /LRIC/i ? 2 : 1;
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
            $wsheet->{firstTableNumber} =
              $model->{method} && $model->{method} =~ /LRIC/i ? 2 : 1;
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
            $wsheet->{firstTableNumber} =
              $model->{method} && $model->{method} =~ /LRIC/i ? 2 : 1;
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
        $wsheet->freeze_panes( $model->{tariff1Row} || 1, 2 );
        $wsheet->set_column( 0, 0,   16 );
        $wsheet->set_column( 1, 1,   50 );
        $wsheet->set_column( 2, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( lines => 'Results' ),
          @{ $model->{tariffTables} };
        if ( $model->{checksum_1_7} ) {
            require Spreadsheet::WriteExcel::Utility;
            my ( $sh, $ro, $co ) =
              $model->{checksum_1_7}->wsWrite( $wbook, $wsheet );
            $sh = $sh->get_name;
            my $cell =
              qq%'$sh'!%
              . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co );
            my $checksumText = qq%" [checksum "&TEXT($cell,"000 0000")&"]"%;
            $model->{checksumAppend}{$wbook} =
              qq%&IF(ISNUMBER($cell),$checksumText,"")%;
        }
        if ( $model->{tariff1Row} ) {
            $wsheet->set_row( $model->{tariff1Row} - 2, 42 );
        }
      }

      ,

      $model->{summaries} && $model->{summaries} =~ /matri/i
      ?

      (
        'Mat' => sub {
            my ($wsheet) = @_;
            $wsheet->{sheetNumber} = 46;
            $wsheet->freeze_panes( 1, 2 );
            $wsheet->set_column( 0, 0,   20 );
            $wsheet->set_column( 1, 1,   50 );
            $wsheet->set_column( 2, 250, 20 );
            my ( @matrices, @total, @diff );
            foreach my $col ( 0, 1 ) {
                my $name =
                  $col
                  ? 'capacity charge p/kVA/day'
                  : "$model->{timebandName} rate p/kWh";
                push @{ $model->{matricesData}[$col] },
                  $total[$col] = Arithmetic(
                    name          => 'Total notional ' . $name,
                    defaultFormat => $col ? '0.00soft' : '0.000soft',
                    arithmetic    => '='
                      . join( '+',
                        map { "A$_" } 1 .. @{ $model->{matricesData}[$col] } ),
                    arguments => {
                        map {
                            ( "A$_" => $model->{matricesData}[$col][ $_ - 1 ] );
                        } 1 .. @{ $model->{matricesData}[$col] }
                    },
                  );

                # push @{ $model->{matricesData}[$col] },
                $diff[$col] = Arithmetic(
                    name          => "Difference $name",
                    arithmetic    => '=A1-A2',
                    defaultFormat => $total[$col]{defaultFormat},
                    arguments     => {
                        A1 => $total[$col],
                        A2 =>
                          $model->{tariffTables}[0]{columns}[ 1 + 2 * $col ],
                    }
                );
                unshift @{ $model->{matricesData}[$col] },
                  Stack( sources => [ $model->{tariffTables}[0]{columns}[0] ] );
                push @matrices,
                  Columnset(
                    name    => "Total $name",
                    columns => $model->{matricesData}[$col]
                  );
            }

            my $purpleUse =
              Stack( sources => [ $model->{matricesData}[2] ] );

            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( lines => 'Matrices and revenue summary' ),
              @matrices,
              Columnset(
                name    => 'Consistency check',
                columns => [
                    Stack(
                        sources => [ $model->{tariffTables}[0]{columns}[0] ]
                    ),
                    $diff[0],
                    $purpleUse,
                    $diff[1],
                    Arithmetic(
                        name          => 'This should be zero (p/kVA/day)',
                        defaultFormat => '0.00soft',
                        arithmetic    => '=A1*A3*A4/A5+A6',
                        arguments     => {
                            A1 => $diff[0],
                            A3 => $purpleUse,
                            A4 => $model->{matricesData}[3],
                            A5 => $model->{matricesData}[4],
                            A6 => $diff[1],
                        }
                    ),
                ]
              ),
              @{ $model->{revenueTables} };
        },
      )

      :

      (
        $model->{vertical} ? 'Summary' : 'HSummary' => sub {
            my ($wsheet) = @_;
            $wsheet->{sheetNumber} = 46;
            $wsheet->freeze_panes( $model->{tariff1Row} || 1, 2 );
            $wsheet->set_landscape unless $model->{vertical};
            $wsheet->set_column( 0, 0,   20 );
            $wsheet->set_column( 1, 1,   50 );
            $wsheet->set_column( 2, 250, 20 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( lines => 'Revenue summary' ),
              @{ $model->{revenueTables} };
        },
      )

      if !$model->{ldnoRev} || $model->{ldnoRev} !~ /only/i;

    push @wsheetsAndClosures,

      $model->{omitDnoTotals}
      ? ()
      : $model->{transparency} ? (
        'DNO totals' => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->set_column( 0, 250, 30 );
            $wsheet->{sheetNumber} = 47;
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( name => 'DNO totals' ),
              @{ $model->{aggregateTables} };
            $wsheet->{sheetNumber} = 48;
        }
      )
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
      ),

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
              SpreadsheetModel::Book::FormatLegend->new(1);
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
              SpreadsheetModel::Book::FormatLegend->new(1);
            $_->wsWrite( $wbook, $wsheet )
              foreach @{ $model->{tablesTemplateExport} };
            delete $wbook->{noLinks};
            $wbook->{logger} = $logger if $logger;
            $wbook->{noLinks} = $noLinks;
        },
        VBACode => sub {
            my ($wsheet) = @_;
            $model->vbaWrite( $wbook, $wsheet );
        }
      )
      : (),

      $model->{ldnoRev}
      ?

      (

        ( $model->{ldnoRev} =~ /qno/i ? 'QNORev' : 'LDNORev' ) => sub {
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

      $model->{ldnoMarginColumns}
      ? (
        ( ( $model->{scenario} || 'Baseline' ) . '$' ) => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 1 );
            $wsheet->set_column( 0, 0,   50 );
            $wsheet->set_column( 1, 250, 20 );
            $wsheet->{sheetNumber} = 62;
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( name => 'Margins' ),
              Columnset(
                name    => 'Revenues and margins',
                columns => $model->{ldnoMarginColumns},
              );
        }
      )
      : (),

      $model->{TotalsTables}
      ? (
        'Total revenue' => sub {
            my ($wsheet) = @_;
            $wsheet->{sheetNumber} = 61;
            $wsheet->freeze_panes( 1, 1 );
            $wsheet->set_column( 0, 250, 28 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( lines => 'Total revenue' ),
              @{ $model->{TotalsTables} };
        }
      )
      : ()

      ,

      ( $model->{legacy201} ? 'Overview' : 'Index' ) =>
      SpreadsheetModel::Book::FrontSheet->new(
        model => $model,
        $model->{legacy201} ? ( name => 'Overview' ) : (),
        copyright =>
          'Copyright 2009-2012 Energy Networks Association Limited and others. '
          . 'Copyright 2013-2017 Franck Latrémolière, Reckon LLP and others.'
      )->closure($wbook);

    @wsheetsAndClosures;

}

1;
