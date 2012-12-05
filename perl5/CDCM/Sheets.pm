package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2012 DCUSA Limited and others. All rights reserved.

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
use CDCM::ModelNotes;

sub frontSheets {
    my ($model) = @_;
    $model->{model100}
      ? qw(Overview Input Tariffs Summary M-ATW M-Rev)
      : qw(Overview Input);
}

sub worksheetsAndClosures {
    my ( $model, $wbook ) = @_;

    $wbook->{lastSheetNumber} = 19;

    my @wsheetsAndClosures;

    push @wsheetsAndClosures,

      'Config' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber} = 11;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->fit_to_pages( 1, 1 );
        $wsheet->set_column( 0, 0, 50 );
        $wsheet->set_column( 1, 1, 100 );
        $model->configNotes->wsWrite( $wbook, $wsheet );
      }

      if $model->{detailedCosts};

    push @wsheetsAndClosures,

      'Input' => sub {
        my ($wsheet) = @_;
        $wsheet->{sheetNumber} = 11;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->{sheetNumber} ||= ++$wbook->{lastSheetNumber};
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );
        $wsheet->{nextFree} = 2;
        my $te = Dataset(
            number        => 1000,
            dataset       => $model->{dataset},
            name          => 'Company, charging year, data version',
            cols          => Labelset( list => [qw(Company Year Version)] ),
            defaultFormat => 'texthard',
            data =>
              [ 'Illustrative company', '9000/9001', 'Illustrative dataset' ]
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
          foreach sort { ( $a->{number} || 9999 ) <=> ( $b->{number} || 9999 ) }
          @{ $model->{inputTables} };
        my $nextFree = delete $wsheet->{nextFree};
        $model->inputDataNotes->wsWrite( $wbook, $wsheet );
        $wsheet->{nextFree} = $nextFree;
      }

      if $model->{inputData};

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
        $wsheet->{_black_white} = 1;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->lafNotes, @{ $model->{routeing} };
      },

      'DRM' => sub {
        my ($wsheet) = @_;
        $wsheet->{_black_white} = 1;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 24 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->networkModelNotes, @{ $model->{networkModel} };
      },

      $model->{noSM} ? () : (
        'SM' => sub {
            my ($wsheet) = @_;
            $wsheet->{_black_white} = 1;
            $wsheet->freeze_panes( 1, 0 );
            $wsheet->set_column( 0, 0,   50 );
            $wsheet->set_column( 1, 250, 24 );
            $_->wsWrite( $wbook, $wsheet )
              foreach $model->serviceModelNotes, @{ $model->{serviceModels} };
        }
      ),

      'Loads' => sub {
        my ($wsheet) = @_;
        $wsheet->{_black_white} = 1;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->loadsNotes, @{ $model->{loadProfiles} },
          @{ $model->{volumeData} };
      },

      'Multi' => sub {
        my ($wsheet) = @_;
        $wsheet->{_black_white} = 1;
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
        $wsheet->{_black_white} = 1;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->smlNotes, @{ $model->{forecastSml} };
      },

      'AMD' => sub {
        my ($wsheet) = @_;
        $wsheet->{_black_white} = 1;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->amlNotes, @{ $model->{forecastAml} };
      },

      ( $model->{opAlloc} ? 'Opex' : 'Otex' ) => sub {
        my ($wsheet) = @_;
        $wsheet->{_black_white} = 1;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->operatingNotes, @{ $model->{operatingExpenditure} };
      },

      'Contrib' => sub {
        my ($wsheet) = @_;
        $wsheet->{_black_white} = 1;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 24 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->contributionNotes, @{ $model->{contributions} };
      },

      'Yard' => sub {
        my ($wsheet) = @_;
        $wsheet->{_black_white} = 1;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->yardstickNotes, @{ $model->{yardsticks} };
      },

      'Standing' => sub {
        my ($wsheet) = @_;
        $wsheet->{_black_white} = 1;
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

      'NHH' => sub {
        my ($wsheet) = @_;
        $wsheet->{_black_white} = 1;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->standingNhhNotes, @{ $model->{standingNhh} };
      },

      'Reactive' => sub {
        my ($wsheet) = @_;
        $wsheet->{_black_white} = 1;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet )
          for $model->reactiveNotes, @{ $model->{reactiveResults} };
      },

      'Aggreg' => sub {
        my ($wsheet) = @_;
        $wsheet->{_black_white} = 1;
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
        $wsheet->{_black_white} = 1;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 21 );
        $_->wsWrite( $wbook, $wsheet )
          foreach $model->revenueNotes, @{ $model->{revenueMatching} };
      };

    push @wsheetsAndClosures,

      'Scaler' => sub {
        my ($wsheet) = @_;
        $wsheet->{_black_white} = 1;
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
        $wsheet->{_black_white} = 1;
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
        $wsheet->{_black_white} = 1;
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
        0 and $wsheet->activate;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->fit_to_pages( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );
        $wbook->{lastSheetNumber} = 36 if $wbook->{lastSheetNumber} < 36;
        push @{ $wbook->{prohibitedTableNumbers} }, 3701 if $model->{pcd};
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => 'Tariffs' ), @{ $model->{tariffSummary} };

      };

    push @wsheetsAndClosures,

      'Components' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );

        $model->componentNotes->wsWrite( $wbook, $wsheet );

        my $i      = $wbook->{inputSheet};
        my $d      = $wbook->{dataSheet};
        my $logger = $wbook->{logger};
        delete $wbook->{$_} foreach qw(inputSheet dataSheet logger);

        my $cset = $model->{tariffComponentMap};
        $cset->wsWrite( $wbook, $wsheet );
        my $r = $cset->{columns}[0]{$wbook}{row};
        $wsheet->set_row( $_ + $r, 48 )
          foreach 0 .. $cset->{columns}[0]->lastRow;

        $wbook->{logger}     = $logger if $logger;
        $wbook->{inputSheet} = $i      if $i;
        $wbook->{dataSheet}  = $d      if $d;

      }

      if $model->{components};

    push @wsheetsAndClosures,

      'Change' => sub {
        my ($wsheet) = @_;
        $wsheet->{_black_white} = 1;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->fit_to_pages( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );

        $_->wsWrite( $wbook, $wsheet ) foreach Notes(
            name => 'Statistics (including estimated average tariff changes)',
            $model->{oneSheet} ? () : (
                lines => [
                    split /\n/,
                    <<'EOL'
This sheet is for information only.  It can be deleted without affecting any calculations elsewhere in the model.
Information in this sheet should not be relied upon for any commercial purpose.
The only outputs from the model that are intended to comply with the methodology are in the Tariff sheet.
EOL
                ]
            )
          ),
          @{ $model->{overallSummary} };

      }

      if $model->{summary} && $model->{summary} =~ /change/i;

    push @wsheetsAndClosures,

      'Summary' => sub {
        my ($wsheet) = @_;
        $wsheet->{_black_white} = 1;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->fit_to_pages( 1, 1 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );

        $_->wsWrite( $wbook, $wsheet ) foreach Notes(
            name => 'Summary statistics',
            $model->{oneSheet} ? () : (
                lines => [
                    split /\n/,
                    <<'EOL'
This sheet is for information only.  It can be deleted without affecting any calculations elsewhere in the model.
EOL
                ]
            )
          ),
          @{ $model->{overallSummary} };

      }

      if $model->{summary} && $model->{summary} !~ /change/i;

    push @wsheetsAndClosures,

      'M-ATW' => sub {
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

        $wsheet->{nextFree} = 4 + @tables unless $model->{oneSheet};
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
            $model->{oneSheet} ? () : (
                lines => [
                    split /\n/,
                    <<'EOL'
This sheet provides matrices breaking down each tariff component into its elements.
This sheet is for information only.  It can be deleted without affecting any calculations elsewhere in the model.
EOL
                ]
            ),
            sourceLines => \@tables
        );

        delete $wbook->{noLinks};
        $notes->wsWrite( $wbook, $wsheet, $model->{oneSheet} ? () : ( 0, 0 ) );
        $logger->log($notes) if $logger;
        $wbook->{logger} = $logger if $logger;
        $wbook->{noLinks} = $noLinks;

      }

      if $model->{matrices};

    push @wsheetsAndClosures,

      'M-LDNO' => sub {
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

        $wsheet->{nextFree} = 4 + @tables unless $model->{oneSheet};
        $count = 0;
        my @breaks;
        foreach (@pairs) {
            push @breaks, $wsheet->{nextFree} if $_->{name};
            $_->wsWrite( $wbook, $wsheet );
        }
        $wsheet->set_h_pagebreaks(@breaks);

        my $notes = Notes(
            name => 'Tariff matrices for embedded network tariffs',
            $model->{oneSheet} ? () : (
                lines => [
                    split /\n/,
                    <<'EOL'
This sheet provides matrices breaking down each tariff component into its elements.
This sheet is for information only.  It can be deleted without affecting any calculations elsewhere in the model.
EOL
                ]
            ),
            sourceLines => \@tables
        );

        delete $wbook->{noLinks};
        $notes->wsWrite( $wbook, $wsheet, $model->{oneSheet} ? () : ( 0, 0 ) );
        $logger->log($notes) if $logger;
        $wbook->{logger} = $logger if $logger;
        $wbook->{noLinks} = $noLinks;

      }

      if $model->{matrices} and $model->{portfolio} || $model->{boundary};

    push @wsheetsAndClosures,

      'M-Rev' => sub {
        my ($wsheet) = @_;
        $wsheet->{_black_white} = 1;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->set_landscape;
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );

        # Set noLinks for this sheet, since
        # any links would be to unnumbered tariff matrix tables.

        my $noLinks = $wbook->{noLinks};
        $wbook->{noLinks} = 1;
        $_->wsWrite( $wbook, $wsheet ) foreach Notes(
            name => 'Revenue matrix',
            $model->{oneSheet} ? () : (
                lines => [
                    split /\n/,
                    <<'EOL'
This sheet is for information only.  It can be deleted without affecting any calculations elsewhere in the model.
EOL
                ]
            ),
          ),
          $model->revenueMatrices;
        $wbook->{noLinks} = $noLinks;
      }

      if $model->{matrices} && $model->{matrices} =~ /big/i;

    push @wsheetsAndClosures,

      'CData' => sub {
        my ($wsheet) = @_;
        $wsheet->{_black_white} = 1;
        $wsheet->freeze_panes( 1, 1 );
        $wsheet->{sheetNumber} ||= ++$wbook->{lastSheetNumber};
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $wsheet ) foreach Notes(
            name => 'Additional calculations for tariff comparisons',
            $model->{oneSheet} ? () : (
                lines => [
                    split /\n/,
                    <<'EOL'
This sheet is for information only.  It can be deleted without affecting any calculations elsewhere in the model.
EOL
                ]
            )
          ),
          @{ $model->{consultationInput} };
      },

      'CTables' => sub {
        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->set_column( 0, 0,   50 );
        $wsheet->set_column( 1, 250, 16 );
        $_->wsWrite( $wbook, $wsheet ) foreach Notes(
            name => 'Tariff comparisons',
            $model->{oneSheet} ? () : (
                lines => [
                    split /\n/,
                    <<'EOL'
This sheet is for information only.  It can be deleted without affecting any calculations elsewhere in the model.
EOL
                ]
            )
          ),
          @{ $model->{consultationTables} };
      }

      if $model->{summary} && $model->{summary} =~ /consul/i;

    push @wsheetsAndClosures,

      Overview => sub {

        my ($wsheet) = @_;
        $wsheet->freeze_panes( 1, 0 );
        $wsheet->fit_to_pages( 1, 2 );
        $wsheet->set_column( 0, 0,   30 );
        $wsheet->set_column( 1, 1,   120 );
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
                  . POSIX::strftime(
                    '%a %e %b %Y %H:%M:%S',
                    $model->{localTime} ? @{ $model->{localTime} } : localtime
                  )
                  . ( $ENV{SERVER_NAME} ? " by $ENV{SERVER_NAME}" : '' ),
            ]
        )->wsWrite( $wbook, $wsheet );

      };

    @wsheetsAndClosures;

}

1;
