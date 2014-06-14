package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others.
Copyright 2011-2014 Franck Latrémolière, Reckon LLP and others.

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
use SpreadsheetModel::Shortcuts 'Notes';

sub sheetPriority {
    my ( $model, $sheet ) = @_;
    return (
        (
            grep { $sheet eq $_ }
              $model->{frontSheets}
            ? @{ $model->{frontSheets} }
            : qw(Index Overview)
        ) ? ( $sheet =~ /^(?:Overview|Index)$/is ? 2 : 1 ) : 0
    ) unless $_[0]{arp};
    my $score = {
        'Index$'       => 80,
        'Assumptions$' => 60,
        'Schedule 15$' => 50,
        'Statistics$'  => 40,
        'Tariffs$'     => 30,
    }->{$sheet};
    $score = 10 if !$score && $sheet =~ /\$$/;
    $score;
}

sub worksheetsAndClosures {
    my ( $model, $wbook ) = @_;

    my @wsheetsAndClosures;

    push @wsheetsAndClosures,

      'Input' => sub {
        my ($wsheet) = @_;

        # reset in case of building several models in a single workbook
        delete $wbook->{highestAutoTableNumber};
        $wbook->{lastSheetNumber} = 19;

        $wsheet->{sheetNumber} = 11;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->{sheetNumber} ||= ++$wbook->{lastSheetNumber};
        my $t1001width =
             $model->{targetRevenue}
          && $model->{targetRevenue} =~ /DCP132/i
          && $model->{targetRevenue} !~ /DCP132longlabels/i;
        $wsheet->set_column( 0, 0,   $t1001width ? 64 : 50 );
        $wsheet->set_column( 1, 250, $t1001width ? 24 : 20 );
        $wsheet->{nextFree} = 2;
        my ( $sh, $ro, $co ) = (
            $model->{table1000} = Dataset(
                number        => 1000,
                dataset       => $model->{dataset},
                name          => 'Company, charging year, data version',
                cols          => Labelset( list => [qw(Company Year Version)] ),
                defaultFormat => 'texthard',
                data =>
                  [ 'Illustrative company', 'Year', 'Illustrative dataset' ]
            )
        )->wsWrite( $wbook, $wsheet );
        $sh = $sh->get_name;
        $wbook->{titleAppend} =
            qq%" for "&'$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co )
          . qq%&" in "&'$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co + 1 )
          . qq%&" ("&'$sh'!%
          . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co + 2 )
          . '&")"';

        if ( $model->{nickName} ) {
            use bytes;
            $model->{nickName} =
              qq%="$model->{nickName}"&$wbook->{titleAppend}%;
        }
        $_->wsWrite( $wbook, $wsheet )
          foreach sort { ( $a->{number} || 9999 ) <=> ( $b->{number} || 9999 ) }
          @{ $model->{inputTables} };
        push @{ $model->{sheetLinks} },
          my $inputDataNotes = $model->inputDataNotes;
        my $nextFree = delete $wsheet->{nextFree};
        my $width    = 1;
        foreach ( @{ $model->{inputTables} } ) {
            my $w = 0;
            $w += 1 + $_->lastCol
              foreach $_->{columns} ? @{ $_->{columns} } : $_;
            $width = $w if $w > $width;
        }
        $wsheet->print_area( 0, 0, $nextFree - 1, $width );
        $inputDataNotes->wsWrite( $wbook, $wsheet );
        $wsheet->{nextFree} = $nextFree;
      };

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
              foreach $model->serviceModelNotes, @{ $model->{serviceModels} };
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
          foreach $model->operatingNotes, @{ $model->{operatingExpenditure} };
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

      ( $model->{tariffs} =~ /dcp179|pc12hh|pc34hh/i ? 'AggCap' : 'NHH' ) =>
      sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->standingNhhNotes, @{ $model->{standingNhh} };
      },

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
          $model->{model100} ? @{ $model->{postPcdApplicationResults} } : (),
          @{ $model->{revenueSummaryTables} },
          $model->{model100} ? ()
          : $model->{postPcdApplicationResults}
          ? @{ $model->{postPcdApplicationResults} }
          : ();

      };

    push @wsheetsAndClosures,

      'Tariffs' => sub {
        my ($wsheet) = @_;
        unless ( $model->{arp} ) {
            $wsheet->freeze_panes( 1, 1 );
            $wsheet->fit_to_pages( 1, 1 );
        }
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );
        $wbook->{lastSheetNumber} = 36 if $wbook->{lastSheetNumber} < 36;
        push @{ $wbook->{prohibitedTableNumbers} }, 3701 if $model->{pcd};
        push @{ $model->{sheetLinks} }, my $notes = Notes( name => 'Tariffs' );
        $_->wsWrite( $wbook, $wsheet )
          foreach $notes,
          @{ $model->{tariffSummary} };

      };

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
        unless ( $model->{arp} ) {
            $wsheet->freeze_panes( 1, 1 );
            $wsheet->fit_to_pages( 1, 1 );
            $wsheet->set_column( 0, 0,   50 );
            $wsheet->set_column( 1, 250, 20 );
        }
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
        $wbook->{noLinks} = 1 if $model->{matrices} =~ /big|nol/i;

        my @pairs =
          $model->{niceTariffMatrices}->( sub { local ($_) = @_; !/LDNO/i } );
        my @tables;

        my $count = 0;
        foreach (@pairs) {
            push @tables, $_ if $_->{name};
        }

        $wsheet->{nextFree} = 4 + @tables;
        $count = 0;
        my @breaks;
        foreach (@pairs) {
            push @breaks, $wsheet->{nextFree} if $_->{name};
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
                split /\n/,
                <<'EOL'
This sheet provides matrices breaking down each tariff component into its elements.
This sheet is for information only.  It can be deleted without affecting any calculations elsewhere in the model.
EOL
            ],
            sourceLines => \@tables
        );

        delete $wbook->{noLinks};
        $notes->wsWrite( $wbook, $wsheet, 0, 0 );
        $logger->log($notes) if $logger;
        $wbook->{logger} = $logger if $logger;
        $wbook->{noLinks} = $noLinks;

      }

      if $model->{matrices};

    push @wsheetsAndClosures,

      'M-LDNO' => sub {
        return if $wbook->{findForwardLinks};
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->set_column( 0, 0,   40 );
        $wsheet->set_column( 1, 250, 20 );

        my $logger = $wbook->{logger};
        delete $wbook->{logger};
        my $noLinks = $wbook->{noLinks};
        $wbook->{noLinks} = 1 if $model->{matrices} =~ /big|nol/i;

        my @pairs =
          $model->{niceTariffMatrices}->( sub { local ($_) = @_; /LDNO/i } );
        my @tables;

        my $count = 0;
        foreach (@pairs) {
            push @tables, $_ if $_->{name};
        }

        $wsheet->{nextFree} = 4 + @tables;
        $count = 0;
        my @breaks;
        foreach (@pairs) {
            push @breaks, $wsheet->{nextFree} if $_->{name};
            $_->wsWrite( $wbook, $wsheet );
        }
        $wsheet->set_h_pagebreaks(@breaks);

        my $notes = Notes(
            name  => 'Tariff matrices for embedded network tariffs',
            lines => [
                split /\n/,
                <<'EOL'
This sheet provides matrices breaking down each tariff component into its elements.
This sheet is for information only.  It can be deleted without affecting any calculations elsewhere in the model.
EOL
            ],
            sourceLines => \@tables
        );

        delete $wbook->{noLinks};
        $notes->wsWrite( $wbook, $wsheet, 0, 0 );
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

        # Set noLinks for this sheet, since
        # any links would be to unnumbered tariff matrix tables.

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

      'Stats' => sub {
        my ($wsheet) = @_;
        $wsheet->set_landscape;
        unless ( $model->{arp} ) {
            $wsheet->freeze_panes( 1, 1 );
            $wsheet->fit_to_pages( 1, 1 );
            $wsheet->set_column( 0, 0,   64 );
            $wsheet->set_column( 1, 250, 20 );
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

      ( $model->{model100} ? 'Overview' : 'Index' ) => sub {
        my ($wsheet) = @_;
        unless ( $model->{arp} ) {
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->fit_to_pages( 1, 2 );
            $wsheet->set_column( 0, 0,   30 );
            $wsheet->set_column( 1, 1,   105 );
            $wsheet->set_column( 2, 250, 30 );
            $_->wsWrite( $wbook, $wsheet )
              foreach $model->topNotes, $wbook->colourCode,
              $wbook->{logger};
        }
        $model->technicalNotes->wsWrite( $wbook, $wsheet );
      };

    return @wsheetsAndClosures unless $model->{arp};

    for ( my $i = 0 ; $i < @wsheetsAndClosures ; $i += 2 ) {
        if ( $wsheetsAndClosures[$i] eq 'Tariffs' ) {
            $wsheetsAndClosures[$i] .= '$';
        }
        else {
            $wsheetsAndClosures[$i] = "CDCM/$wsheetsAndClosures[$i]";
        }
    }

    $model->{sharedData}
      ? $model->{sharedData}
      ->worksheetsAndClosuresMulti( $model, $wbook, @wsheetsAndClosures )
      : @wsheetsAndClosures;

}

sub modelIdentification {
    my ( $model, $wb, $ws ) = @_;
    return $model->{identification} if $model->{identification};
    my ( $w, $r, $c ) = $model->{table1000}->wsWrite( $wb, $ws );
    $model->{identification} =
        q%='%
      . $w->get_name . q%'!%
      . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $r, $c + 1 );
}

sub technicalNotes {
    my ($model) = @_;
    require POSIX;
    Notes(
        name       => '',
        rowFormats => ['caption'],
        lines      => [
            'Technical notes, configuration and code identification',
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
        name  => 'Overview',
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

            <<'EOL',

Copyright 2009-2011 Energy Networks Association Limited and others. Copyright 2011-2014 Franck Latrémolière, Reckon LLP and others. 
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

        ]
    );
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
    Notes(
        lines => <<'EOL'
Input data

This sheet contains all the input data (except LLFCs which can be entered directly into the Tariff sheet).
EOL
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
