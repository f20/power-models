package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others.
Copyright 2011-2017 Franck Latrémolière, Reckon LLP and others.

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

sub finishWorkbook {
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

sub sheetPriority {
    my ( $model, $sheet ) = @_;
    my $score = 0;
    $score = 5 if !$score && $sheet =~ /Tariffs\$$/;
    $score ||= 2
      if $sheet =~ /(?:Overview|Index)$/is || $sheet =~ /\//;
    $score ||= 1
      if $model->{frontSheets} && grep { $sheet eq $_ }
      @{ $model->{frontSheets} };
    $score;
}

sub worksheetsAndClosures {

    my ( $model, $wbook ) = @_;

    my @wsheetsAndClosures;

    push @wsheetsAndClosures,

      'Input' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber}    = 11;
        $wbook->{lastSheetNumber} = 19;
        $wsheet->freeze_panes( 1, 1 );
        my $t1001width = $model->{targetRevenue}
          && $model->{targetRevenue} =~ /dcp132/i;
        $wsheet->set_column( 0, 0,   $t1001width ? 64 : 50 );
        $wsheet->set_column( 1, 250, $t1001width ? 24 : 20 );
        $model->{titleWrites}{$wbook} = [];
        $wbook->{titleWriter} ||=
          sub { push @{ $model->{titleWrites}{$wbook} }, [@_]; };

        unless ( $model->{table1000} ) {
            $model->{table1000} = Dataset(
                number        => 1000,
                dataset       => $model->{dataset},
                name          => 'Company, charging year, data version',
                cols          => Labelset( list => [qw(Company Year Version)] ),
                defaultFormat => 'puretexthard',
                data          => [ 'no company', 'no year', 'no data version' ],
                usePlaceholderData => 1,
                forwardLinks       => {},
                appendTo           => $model->{inputTables},
            );
            push @{ $model->{edcmTables} },
              Stack(
                name =>
                  'EDCM input data ⇒1100. Company, charging year, data version',
                sources => [ $model->{table1000} ],
              ) if $model->{edcmTables};
        }
        my $inputDataNotes = $model->inputDataNotes;
        push @{ $model->{sheetLinks}{$wbook} }, $inputDataNotes;
        $_->wsWrite( $wbook, $wsheet )
          foreach $inputDataNotes,
          sort { ( $a->{number} || 9999 ) <=> ( $b->{number} || 9999 ) }
          @{ $model->{inputTables} };
        require Spreadsheet::WriteExcel::Utility;
        my ( $sh, $ro, $co ) = $model->{table1000}->wsWrite( $wbook, $wsheet );
        $sh = $sh->get_name;
        $model->{idAppend}{$wbook} =
            qq%&" for "&'$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co )
          . qq%&" in "&'$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co + 1 )
          . qq%&" ("&'$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co + 2 )
          . '&")"';
        $model->{nickNames}{$wbook} =
          qq%="$model->{nickName}"$model->{idAppend}{$wbook}%
          if $model->{nickName};
      };

    if ( $model->{embeddedModelM}
        && UNIVERSAL::can( $model->{embeddedModelM}, 'worksheetsAndClosures' ) )
    {
        my @mwac = $model->{embeddedModelM}->worksheetsAndClosures($wbook);
        while (@mwac) {
            my $sheet   = shift @mwac;
            my $closure = shift @mwac;
            next if $sheet =~ /^(?:Index|Result)/;
            next if $sheet =~ /^Input/ && !$model->{noSingleInputSheet};
            push @wsheetsAndClosures, "M($sheet)", $closure;
        }
    }

    push @wsheetsAndClosures, 'CDCM Revenues' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber} = 10;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   60 );
        $wsheet->set_column( 1, 250, 30 );
        my $dataSheet = delete $wbook->{dataSheet};
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'CDCM Revenues' ),
          @{ $model->{inputTable1001} };
        $wbook->{dataSheet} = $dataSheet;
      }
      if $model->{inputTable1001};

    push @wsheetsAndClosures,

      'Preprocessing' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber} ||= ++$wbook->{lastSheetNumber};
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes(
            name => 'Preprocessing and sanitisation of input data' );
      }

      if $model->{preprocessing};

    push @wsheetsAndClosures,

      'LAFs' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->lafNotes, @{ $model->{routeing} };
      },

      'DRM' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 24 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->networkModelNotes, @{ $model->{networkModel} };
      },

      $model->{noSM} ? () : (
        'SM' => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->set_column( 0, 0,   50 );
            $wsheet->set_column( 1, 250, 24 );
            $_->wsWrite( $wbook, $wsheet )
              foreach $model->serviceModelNotes,
              @{ $model->{serviceModels} };
        }
      ),

      'Loads' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->loadsNotes, @{ $model->{loadProfiles} },
          @{ $model->{volumeData} };
      },

      'Multi' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->print_area('A:V');
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->multiNotes, @{ $model->{timeOfDayResults} };
      },

      'SMD' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->smlNotes, @{ $model->{forecastSml} };
      },

      'AMD' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->amlNotes, @{ $model->{forecastAml} };
      },

      ( $model->{opAlloc} ? 'Opex' : 'Otex' ) => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->operatingNotes,
          @{ $model->{operatingExpenditure} };
      },

      'Contrib' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 24 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->contributionNotes, @{ $model->{contributions} };
      },

      'Yard' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->yardstickNotes, @{ $model->{yardsticks} };
      },

      'Standing' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->standingNotes, @{ $model->{standingResults} },
          $model->{unauth}
          ? @{ $model->{unauthorisedDemand} }
          : ();
      },

      $model->{fixedCap} && $model->{fixedCap} =~ /nosheet/ ? () : (
        (
            $model->{tariffs}
              && $model->{tariffs} =~ /dcp179|pc12hh|pc34hh/i
            ? 'AggCap'
            : 'NHH'
        ) => sub {
            my ($wsheet) = @_;
            $wsheet->freeze_panes( 1, 1 );
            $wsheet->set_landscape;
            $wsheet->set_column( 0, 0,   50 );
            $wsheet->set_column( 1, 250, 16 );
            $_->wsWrite( $wbook, $wsheet )
              foreach $model->standingNhhNotes, @{ $model->{standingNhh} };
        }
      ),

      'Reactive' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet )
          for $model->reactiveNotes, @{ $model->{reactiveResults} };
      },

      'Aggreg' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->aggregationNotes,
          @{ $model->{preliminaryAggregation} };
      },

      'Revenue' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 21 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->revenueNotes, @{ $model->{revenueMatching} };
      };

    push @wsheetsAndClosures,

      'Scaler' => sub {
        my ($wsheet) = @_;
        $wsheet->set_landscape;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 21 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->scalerNotes, @{ $model->{assetScaler} },
          @{ $model->{adderResults} };
      }

      unless $model->{scaler} && $model->{scaler} =~ /adder/i;

    push @wsheetsAndClosures,

      'Adder' => sub {
        my ($wsheet) = @_;
        $wsheet->set_landscape;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 21 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->adderNotes, @{ $model->{adderResults} };
      }

      if $model->{scaler} && $model->{scaler} =~ /adder/i;

    push @wsheetsAndClosures,

      'Adjust' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 28 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->roundingNotes, @{ $model->{roundingResults} },
          $model->{model100} ? @{ $model->{postPcdApplicationResults} }
          : (), @{ $model->{revenueSummaryTables} }, $model->{model100} ? ()
          : $model->{postPcdApplicationResults}
          ? @{ $model->{postPcdApplicationResults} }
          : ();

      },

      'Tariffs' => sub {
        my ($wsheet) = @_;
        unless ( $model->{compact} ) {
            $wsheet->freeze_panes( 1, 1 );
            $wsheet->fit_to_pages( 1, 1 );
        }
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );
        $wbook->{lastSheetNumber} = 36 if $wbook->{lastSheetNumber} < 36;
        push @{ $model->{sheetLinks}{$wbook} },
          my $notes = Notes( name => 'Tariffs' );
        $_->wsWrite( $wbook, $wsheet )
          foreach $notes,
          @{ $model->{tariffSummary} };
        if ( $model->{checksum_1_7} ) {
            require Spreadsheet::WriteExcel::Utility;
            my ( $sh, $ro, $co ) =
              $model->{checksum_1_7}->wsWrite( $wbook, $wsheet );
            $sh = $sh->get_name;
            my $cell = qq%'$sh'!%
              . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co );
            my $checksumText = qq%" [checksum "&TEXT($cell,"000 0000")&"]"%;
            $model->{checksumAppend}{$wbook} =
              qq%&IF(ISNUMBER($cell),$checksumText,"")%;
        }
      }

      unless $model->{unroundedTariffAnalysis}
      && $model->{unroundedTariffAnalysis} =~ /modelg/i;

    push @wsheetsAndClosures,

      'Components' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );

        $model->componentNotes->wsWrite( $wbook, $wsheet );

        my $dataSheet = delete $wbook->{dataSheet};
        my $logger    = delete $wbook->{logger};
        my $cset      = $model->{tariffComponentMap};
        $cset->wsWrite( $wbook, $wsheet );
        my $r = $cset->{columns}[0]{$wbook}{row};
        $wsheet->set_row( $_ + $r, 48 )
          foreach 0 .. $cset->{columns}[0]->lastRow;
        $wbook->{logger}    = $logger    if $logger;
        $wbook->{dataSheet} = $dataSheet if $dataSheet;

      }

      if $model->{components};

    push @wsheetsAndClosures,

      'Change' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->fit_to_pages( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );

        $_->wsWrite( $wbook, $wsheet ) foreach Notes(
            name  => 'Statistics (including estimated average tariff changes)',
            lines => [
                split /\n/,
                <<'EOL'
This sheet is for information only.  It can be deleted without affecting any calculations elsewhere in the model.
Information in this sheet should not be relied upon for any commercial purpose.
The only outputs from the model that are intended to comply with the methodology are in the Tariff sheet.
EOL
            ]
          ),
          @{ $model->{overallSummary} };
      }

      if $model->{summary} && $model->{summary} =~ /change/i;

    push @wsheetsAndClosures,

      'Summary' => sub {
        my ($wsheet) = @_;
        $wsheet->set_landscape;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->fit_to_pages( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );
        $wbook->{lastSheetNumber} = 37 if $wbook->{lastSheetNumber} < 37;
        my $notes = Notes(
            name  => 'Summary',
            lines => [
                split /\n/,
                <<'EOL'
This sheet is for information only.  It can be deleted without affecting any calculations elsewhere in the model.
EOL
            ]
        );
        $_->wsWrite( $wbook, $wsheet )
          foreach $notes,
          @{ $model->{overallSummary} };
      }

      if $model->{summary} && $model->{summary} !~ /change/i;

    push @wsheetsAndClosures,

      'M-ATW' => sub {
        return if $wbook->{findForwardLinks};
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->set_column( 0, 0,   40 );
        $wsheet->set_column( 1, 250, 20 );

        my $logger  = delete $wbook->{logger};
        my $noLinks = $wbook->{noLinks};
        $wbook->{noLinks} = 1;

        my @pairs =
          $model->{niceTariffMatrices}
          ->( sub { local ($_) = @_; !/(?:LD|Q)NO/i } );
        my @tables;

        my $count = 0;
        foreach (@pairs) {
            push @tables, $_ if $_->{columns} && $_->{name};
        }

        $wsheet->{nextFree} = 4 + @tables unless $model->{compact};
        $count = 0;
        my @breaks;
        foreach (@pairs) {
            push @breaks, $wsheet->{nextFree}
              if $_->{columns} && $_->{name} && $wsheet->{nextFree};
            $_->wsWrite( $wbook, $wsheet );
        }
        $wsheet->set_h_pagebreaks(@breaks);

        my $notes = Notes(
            name => 'Tariff matrices'
              . (
                     $model->{portfolio}
                  || $model->{boundary} ? ' (all-the-way tariffs)' : ''
              ),
            lines => [
                split /\n/, <<'EOL', -1
This sheet provides matrices breaking down each tariff component into its elements.
This sheet is for information only.  It can be deleted without affecting any calculations elsewhere in the model.
EOL
            ],
            sourceLines => \@tables
        );

        delete $wbook->{noLinks};
        $notes->wsWrite( $wbook, $wsheet, $model->{compact} ? () : ( 0, 0 ) );
        $logger->log($notes) if $logger;
        $wbook->{logger} = $logger if $logger;
        $wbook->{noLinks} = $noLinks;
      }

      if $model->{matrices};

    push @wsheetsAndClosures,

      'M-'
      . (    $model->{portfolio}
          && $model->{portfolio} =~ /qno/i ? 'QNO' : 'LDNO' ) => sub {
        return if $wbook->{findForwardLinks};
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->set_column( 0, 0,   40 );
        $wsheet->set_column( 1, 250, 20 );

        my $logger  = delete $wbook->{logger};
        my $noLinks = $wbook->{noLinks};
        $wbook->{noLinks} = 1;

        my @pairs =
          $model->{niceTariffMatrices}
          ->( sub { local ($_) = @_; /(?:LD|Q)NO/i } );
        my @tables;

        my $count = 0;
        foreach (@pairs) {
            push @tables, $_ if $_->{columns} && $_->{name};
        }

        $wsheet->{nextFree} = 4 + @tables unless $model->{compact};
        $count = 0;
        my @breaks;
        foreach (@pairs) {
            push @breaks, $wsheet->{nextFree}
              if $_->{columns} && $_->{name} && $wsheet->{nextFree};
            $_->wsWrite( $wbook, $wsheet );
        }
        $wsheet->set_h_pagebreaks(@breaks);

        my $notes = Notes(
            name  => 'Tariff matrices for embedded network tariffs',
            lines => [
                split /\n/, <<'EOL', -1
This sheet provides matrices breaking down each tariff component into its elements.
This sheet is for information only.  It can be deleted without affecting any calculations elsewhere in the model.
EOL
            ],
            sourceLines => \@tables
        );

        delete $wbook->{noLinks};
        $notes->wsWrite( $wbook, $wsheet, $model->{compact} ? () : ( 0, 0 ) );
        $logger->log($notes) if $logger;
        $wbook->{logger} = $logger if $logger;
        $wbook->{noLinks} = $noLinks;
      }

      if $model->{matrices} and $model->{portfolio} || $model->{boundary};

    push @wsheetsAndClosures,

      'M-Rev' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );

        my $noLinks = $wbook->{noLinks};
        $wbook->{noLinks} = 1;
        $_->wsWrite( $wbook, $wsheet ) foreach Notes(
            name  => 'Revenue matrix',
            lines => [
                split /\n/,
                <<'EOL'
This sheet is for information only.  It can be deleted without affecting any calculations elsewhere in the model.
EOL
            ],
          ),
          $model->revenueMatrices;
        $wbook->{noLinks} = $noLinks;
      }

      if $model->{matrices} && $model->{matrices} =~ /big/i;

    push @wsheetsAndClosures,

      'Comp' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber} = 41 unless $wbook->{lastSheetNumber} > 40;
        unless ( $model->{compact} ) {
            $wsheet->freeze_panes( 1, 1 );
            $wsheet->fit_to_pages( 1, 0 );
            $wsheet->set_column( 0, 0,   56 );
            $wsheet->set_column( 1, 250, 16 );
        }
        my $notes = Notes(
            name  => 'Comparisons',
            lines => [
                split /\n/,
                <<'EOL'
This sheet is for information only.  It can be deleted without affecting any calculations elsewhere in the model.
EOL
            ]
        );
        $_->wsWrite( $wbook, $wsheet )
          foreach $notes, @{ $model->{comparisonTables} };
      }

      if $model->{comparisonTables};

    push @wsheetsAndClosures,

      'CData' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $wsheet ) foreach Notes(
            name  => 'Additional calculations for tariff comparisons',
            lines => [
                split /\n/,
                <<'EOL'
This sheet is for information only.  It can be deleted without affecting any calculations elsewhere in the model.
EOL
            ]
          ),
          @{ $model->{consultationInput} };
      },

      'CTables' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet ) foreach Notes(
            name  => 'Tariff comparisons',
            lines => [
                split /\n/,
                <<'EOL'
This sheet is for information only.  It can be deleted without affecting any calculations elsewhere in the model.
EOL
            ]
          ),
          @{ $model->{consultationTables} };
      }

      if $model->{summary} && $model->{summary} =~ /consul/i;

    push @wsheetsAndClosures,

      'Stats' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber} = 42 unless $wbook->{lastSheetNumber} > 41;
        unless ( $model->{compact} ) {
            $wsheet->freeze_panes( 1, 1 );
            $wsheet->fit_to_pages( 1, 0 );
            $wsheet->set_column( 0, 0,   56 );
            $wsheet->set_column( 1, 250, 16 );
        }
        my $notes = Notes(
            name  => 'Statistics',
            lines => [
                split /\n/,
                <<'EOL'
This sheet is for information only.  It can be deleted without affecting any calculations elsewhere in the model.
EOL
            ]
        );
        $_->wsWrite( $wbook, $wsheet )
          foreach $notes, @{ $model->{statisticsTables} };
      }

      if $model->{statisticsTables};

    push @wsheetsAndClosures,

      'Info' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->fit_to_pages( 1, 0 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );
        my $notes = Notes(
            name  => 'Other information',
            lines => [
                split /\n/,
                <<'EOL'
This sheet is for information only.  It can be deleted without affecting any calculations in the model.
EOL
            ]
        );
        $_->wsWrite( $wbook, $wsheet )
          foreach $notes,
          @{ $model->{informationTables} };
        $wsheet->{protectionOptions}{select_locked_cells} = 1;
      }

      if $model->{informationTables};

    push @wsheetsAndClosures,

      '⇒EDCM' => sub {
        my ($wsheet) = @_;
        $wbook->{lastSheetNumber} = 42
          unless $wbook->{lastSheetNumber} > 42;
        $wsheet->set_landscape;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->fit_to_pages( 1, 1 );
        $wsheet->set_column( 0, 0,   48 );
        $wsheet->set_column( 1, 250, 16 );
        if ( ref $model->{edcmTables}[0] eq 'ARRAY' ) {
            my $col = shift @{ $model->{edcmTables} };
            push @{ $model->{edcmTables} }, Columnset(
                name          => 'EDCM input data ⇒1101. Financial information',
                singleRowName => 'Financial information',
                columns       => [
                    map {
                        $col->[$_]
                          || Constant( name => 'Placeholder', data => [], );
                    } 0 .. 5
                ],
            );
        }
        my $notes = Notes(
            name  => 'Data for EDCM model',
            lines => ['This sheet is for information only.']
        );
        $_->wsWrite( $wbook, $wsheet )
          foreach $notes,
          sort { $a->{name} cmp $b->{name} } @{ $model->{edcmTables} };
      }

      if $model->{edcmTables};

    push @wsheetsAndClosures,
      'UTA' => sub {
        my ($wsheet) = @_;
        $wbook->{lastSheetNumber} = 42
          unless $wbook->{lastSheetNumber} > 42;
        $wsheet->fit_to_pages( 1, 0 );
        $wsheet->set_column( 0, 0,   48 );
        $wsheet->set_column( 1, 250, 16 );
        my $notes = Notes( name => 'Unrounded tariff analysis', );
        $_->wsWrite( $wbook, $wsheet ) foreach $notes, @{ $model->{utaTables} };
      }
      if $model->{utaTables} && !$model->{modelgTables};

    push @wsheetsAndClosures,
      'G(Details)' => sub {
        my ($wsheet) = @_;
        $wbook->{lastSheetNumber} = 42
          unless $wbook->{lastSheetNumber} > 42;
        $wsheet->fit_to_pages( 1, 0 );
        $wsheet->set_column( 0, 0,   48 );
        $wsheet->set_column( 1, 250, 16 );
        my $notes = Notes( name => 'Details of Model G calculations', );
        $_->wsWrite( $wbook, $wsheet )
          foreach $notes, @{ $model->{modelgTables} };
      }
      if $model->{modelgTables};

    push @wsheetsAndClosures,
      'G(Summary)' => sub {
        my ($wsheet) = @_;
        $wsheet->fit_to_pages( 1, 0 );
        $wsheet->set_column( 0, 0,   56 );
        $wsheet->set_column( 1, 250, 14 );
        my $noLinks = $wbook->{noLinks};
        $wbook->{noLinks} = 1;
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Summary of Model G calculations', ),
          @{ $model->{modelgSummary} };
        $wbook->{noLinks} = $noLinks;
      }
      if $model->{modelgSummary};

    push @wsheetsAndClosures,
      'G(Results)' => sub {
        my ($wsheet) = @_;
        $wbook->{lastSheetNumber} = 43
          unless $wbook->{lastSheetNumber} > 43;
        $wsheet->fit_to_pages( 1, 0 );
        $wsheet->set_column( 0, 0,   48 );
        $wsheet->set_column( 1, 250, 16 );
        my $notes = Notes( name => 'Model G results', );
        $_->wsWrite( $wbook, $wsheet )
          foreach $notes, @{ $model->{modelgResults} };
      }
      if $model->{modelgResults};

    my $frontSheet = SpreadsheetModel::Book::FrontSheet->new(
        model => $model,
        $model->{legacy201} ? ( name => 'Overview' ) : (),
        copyright =>
          'Copyright 2009-2011 Energy Networks Association Limited and others. '
          . 'Copyright 2011-2017 Franck Latrémolière, Reckon LLP and others.'
    );

    push @wsheetsAndClosures,

      ( $model->{model100} ? 'Overview' : 'Index' ) => (
        $model->{compact} ? sub {
            my ($wsheet) = @_;
            $frontSheet->technicalNotes->wsWrite( $wbook, $wsheet );
        }
        : $frontSheet->closure($wbook)
      );

    if ( $model->{compact} ) {
        my %suffixes = (
            '⇒EDCM' => '',
            $model->{summary} && $model->{summary} =~ /arp/i
            ? (
                Tariffs => ' (2)$',
                Summary => '$',
              )
            : ( Tariffs => '$', ),
        );
        my $reallyCompact = $model->{compact} !~ /not/i;
        for ( my $i = 0 ; $i < @wsheetsAndClosures ; $i += 2 ) {
            my $suffix = $suffixes{ $wsheetsAndClosures[$i] };
            if ( defined $suffix ) {
                $wsheetsAndClosures[$i] .= $suffix;
            }
            elsif ($reallyCompact) {
                $wsheetsAndClosures[$i] = "CDCM/$wsheetsAndClosures[$i]";
            }
        }
    }

    if ( $model->{extraModelsToProcessFirst} ) {
        unshift @wsheetsAndClosures, $_->worksheetsAndClosures($wbook)
          foreach @{ $model->{extraModelsToProcessFirst} };
    }

    @wsheetsAndClosures;

}

sub modelIdentification {
    my ( $model, $wb, $ws ) = @_;
    return $model->{identification} if $model->{identification};
    my ( $w, $r, $c ) = $model->{table1000}->wsWrite( $wb, $ws );
    $model->{identification} = [
        map {
                q%'%
              . $w->get_name . q%'!%
              . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $r,
                $c + $_ )
        } 0 .. 2
    ];
}

sub configNotes {
    Notes(
        lines => <<'EOL'
Model configuration

This sheet enables some names and labels to be configured.  It does not affect calculations.

The list of tariffs, number of timebands and structure of network levels can only be configured when the model is built.

Voltage and network levels are defined as follows:
* 132kV means voltages of at least 132kV. For the purposes of this workbook, 132kV is not included within EHV.
* EHV, in this workbook, means voltages of at least 22kV and less than 132kV.
* HV means voltages of at least 1kV and less than 22kV.
* LV means voltages below 1kV.
EOL
    );
}

sub inputDataNotes {
    my ($model) = @_;
    Notes(
        name => $model->{noSingleInputSheet} ? 'Input data (part)'
        : 'Input data',
        $model->{noSingleInputSheet} || $model->{noLLFCs} ? ()
        : ( lines => 'This sheet contains all the input data, '
              . 'except LLFCs which are entered directly into the Tariff sheet.'
        )
    );
}

sub networkModelNotes {
    Notes(
        lines => <<'EOL'
Network model

This sheet collects data from a network model and calculates aggregated annuitised unit costs from these data.
EOL
    );
}

sub serviceModelNotes {
    Notes(
        lines => <<'EOL'
Service models

This sheet collects and processes data from the service models.
EOL
    );
}

sub lafNotes {
    Notes(
        lines => <<'EOL'
Loss adjustment factors and network use matrices

This sheet calculates matrices of loss adjustment factors and of network use factors.
These matrices map out the extent to which each type of user uses each level of the network, and are used throughout the workbook.
EOL
    );
}

sub loadsNotes {
    my ($model) = @_;
    Notes(
        lines => [
            <<'EOL'
Load characteristics

This sheet compiles information about the assumed characteristics of network users.

A load factor represents the average load of a user or user group, relative to the maximum load level of that user or
user group. Load factors are numbers between 0 and 1.

A coincidence factor represents the expectation value of the load of a user or user group at the time of system maximum load,
relative to the maximum load level of that user or user group.  Coincidence factors are numbers between 0 and 1.

EOL
            , $model->{hasGenerationCapacity}
            ? <<'EOL'
An F factor, for a generator, is the expectation value of output at the time of system maximum load, relative to installed generation capacity.
F factors are user inputs in respect of generators which are credited on the basis of their installed capacity.

EOL
            : (),
            <<'EOL'
A load coefficient is the expectation value of the load of a user or user group at the time of system maximum load, relative to the average load level of that user or user group.
For demand users, the load coefficient is a demand coefficient and can be calculated as the ratio of the coincidence factor to the load factor.
EOL
        ]
    );
}

sub useNotes {
    Notes(
        lines => <<'EOL'
Network use

This sheet combines the volume forecasts and network use matrices in order to estimate the extent to which the network will be used in the charging year.
EOL
    );
}

sub smlNotes {
    Notes(
        lines => <<'EOL'
Forecast simultaneous maximum load
EOL
    );
}

sub amlNotes {
    Notes(
        lines => <<'EOL'
Forecast aggregate maximum load
EOL
    );
}

sub operatingNotes {
    my ($model) = @_;
    Notes(
        lines => [
            $model->{opAlloc}
            ? (
                'Operating expenditure',
                '',
'This sheet calculates elements of tariff components that recover operating expenditure excluding network rates.'
              )
            : 'Other expenditure'
        ]
    );
}

sub contributionNotes {
    Notes( lines => <<'EOL');
Customer contributions

This sheet calculates factors used to take account of the costs deemed to be covered by connection charges.
EOL
}

sub yardstickNotes {
    Notes(
        lines => <<'EOT'
Yardsticks

This sheet calculates average p/kWh and p/kW/day charges that would apply if no costs were recovered through capacity or fixed charges.
EOT
    );
}

sub multiNotes {
    Notes(
        lines => <<'EOL'
Load characteristics for multiple unit rates
EOL
    );
}

sub standingNotes {
    Notes(
        lines => <<'EOL'
Allocation to standing charges

This sheet reallocates some costs from unit charges to fixed or capacity charges, for demand users only.
EOL
    );
}

sub standingNhhNotes {
    Notes(
        lines => <<'EOL'
Standing charges as fixed charges

This sheet allocates standing charges to fixed charges for non half hourly settled demand users.
EOL
    );
}

sub reactiveNotes {
    my ($model) = @_;
    Notes(
        name  => 'Reactive power unit charges',
        lines => [
            $model->{reactive} && $model->{reactive} =~ /band/i
            ? ( 'The calculations in this sheet are '
                  . 'based on steps 1-6 (Ofgem). '
                  . 'This gives banded reactive power unit charges.' )
            : ()
        ]
    );
}

sub aggregationNotes {
    Notes(
        lines => <<'EOL'
Aggregation

This sheet aggregates elements of tariffs excluding revenue matching and final adjustments and rounding.
EOL
    );
}

sub revenueNotes {
    Notes(
        lines => <<'EOL'
Revenue shortfall or surplus
EOL
    );

=head Development note

Matching starts with a summary of charges which has reactive unit charges.

It does not really need to, but doing it that way enables the table of reactive unit
charges against $allTariffsByEndUsers to look good (rather than look like a silly duplicate
at the bottom of the reactive sheet).

=cut

}

sub scalerNotes {
    Notes(
        lines => <<'EOL'
Revenue matching

This sheet modifies tariffs so that the total expected net revenues matches the target.
EOL
    );
}

sub adderNotes {
    Notes(
        lines => <<'EOL'
Adder
EOL
    );
}

sub roundingNotes {
    Notes( lines => <<'EOL');
Tariff component adjustment and rounding
EOL
}

sub componentNotes {
    Notes(
        name  => 'Tariff components and rules',
        lines => [ split /\n/, <<'EOL']
This sheet is for user information only.  It summarises the rules coded in this model to calculate each components of each tariff, and allows some names and labels to be configured.

The following shorthand is used:
PAYG: higher unit rates applicable to tariffs with no standing charges (as opposed to "Standard").
Yardstick: single unit rates (as opposed to "1", "2" etc.)

EOL
    );
}

1;
