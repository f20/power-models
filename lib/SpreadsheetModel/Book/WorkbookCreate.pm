package SpreadsheetModel::Book::WorkbookCreate;

=head Copyright licence and disclaimer

Copyright 2008-2014 Franck Latrémolière, Reckon LLP and others.

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
use SpreadsheetModel::Logger;
use File::Spec::Functions qw(catfile);

sub create {

    my ( $module, $fileName, $instructions, $settings ) = @_;
    my @optionArray =
      ref $instructions eq 'ARRAY' ? @$instructions : $instructions;
    my @localTime   = localtime;
    my $streamMaker = $settings->{streamMaker};
    my $tmpDir;
    $streamMaker ||= sub {
        my ($fn) = @_;
        unless ($fn) {
            binmode STDOUT;
            return \*STDOUT;
        }
        my $finalFile = $fn;
        my $tempFile  = $fn;
        my $closer;
        $tmpDir = '~$tmp-' . $$ unless $^O =~ /win32/i;
        if ($tmpDir) {
            mkdir $tmpDir;
            chmod 0770, $tmpDir;

            $fn =~ s#.*/##s;
            $tempFile = catfile( $tmpDir, $fn );
            $closer = sub {
                rename $tempFile, $finalFile;
                rmdir $tmpDir;
            };

        }
        open my $handle, '>', $tempFile;
        binmode $handle;
        $handle, $closer, $finalFile;
    };

    ( my $handle, my $closer, $fileName ) = $streamMaker->($fileName);
    my $wbook = $module->new($handle);
    my @exports = grep { $settings->{$_} && /^Export/ } keys %$settings;
    my $exporter;
    if (@exports) {
        eval {
            require SpreadsheetModel::Export::Controller;
            $exporter =
              SpreadsheetModel::Export::Controller->new( $fileName, $wbook );
        };
        warn "@exports: $@" if $@;
    }

    # Work around taint issue with IO::File
    $wbook->set_tempdir($tmpDir) if $tmpDir && $module !~ /xlsx/i;

    $wbook->setFormats( $optionArray[0] );
    my @models;
    my ( %allClosures, @wsheetShowOrder, %wsheetActive, %wsheetPassword,
        %sheetDisplayName, @forwardLinkFindingRun, $multiModelSharing );

    foreach my $i ( 0 .. $#optionArray ) {
        if ( my $dataset = $optionArray[$i]{dataset} ) {
            if ( my $yaml = $dataset->{yaml} ) {
                require YAML;    # for deferred parsing
                my @parsed = YAML::Load($yaml);
                if ( @parsed > 1 ) {
                    foreach my $section (@parsed) {
                        while ( my ( $tab, $dat ) = each %$section ) {
                            next unless ref $dat eq 'ARRAY';
                            for ( my $col = 0 ; $col < @$dat ; ++$col ) {
                                my $cd = $dat->[$col];
                                next unless ref $cd eq 'HASH';
                                while ( my ( $row, $v ) = each %$cd ) {
                                    $dataset->{$tab}[$col]{$row} = $v;
                                }
                            }
                        }
                    }
                }
                else {
                    %$dataset = %{ $parsed[0] };
                }
            }
            if ( my $prev = $optionArray[$i]{dataset}{baseDataset} ) {
                unless ( $prev > $i ) {
                    push @{ $optionArray[ $i - $prev ]{requestsToSeeModel} },
                      sub {
                        $optionArray[$i]{sourceModel} = $_[0];
                      };
                }
            }
            else {
                foreach my $overrides (
                    grep { $_ }
                    map  { $optionArray[$i]{$_} }
                    qw(dataOverride ~datasetOverride)
                  )
                {
                    $dataset = Storable::dclone($dataset);
                    foreach my $override (
                        ref $overrides eq 'ARRAY' ? @$overrides : $overrides )
                    {
                        foreach my $itable ( keys %$override ) {
                            for (
                                my $icolumn = 1 ;
                                $icolumn < @{ $override->{$itable} } ;
                                ++$icolumn
                              )
                            {
                                foreach my $irow (
                                    keys %{ $override->{$itable}[$icolumn] } )
                                {
                                    $dataset->{$itable}[$icolumn]{$irow} =
                                      $override->{$itable}[$icolumn]{$irow};
                                }
                            }
                        }
                    }
                }
            }
            $dataset->{usePlaceholderData} ||= $optionArray[$i]{illustrative}
              if $optionArray[$i]{illustrative};
            $optionArray[$i]{dataset} = $dataset;
        }
    }

    foreach ( 0 .. $#optionArray ) {
        my $options = $optionArray[$_];
        my $modelCount = $_ ? ".$_" : '';
        $modelCount = '.' . ( 1 + $_ ) if @optionArray > 1;
        $options->{PerlModule}
          ->setUpMultiModelSharing( \$multiModelSharing, $options,
            \@optionArray )
          if $#optionArray
          && UNIVERSAL::can( $options->{PerlModule}, 'setUpMultiModelSharing' );
        my $model = $options->{PerlModule}->new(%$options);
        map { $_->($model); } @{ $options->{requestsToSeeModel} }
          if $options->{requestsToSeeModel};
        $forwardLinkFindingRun[$_] = $model if $options->{forwardLinks};
        $options->{revisionText} ||= '';
        $wbook->{titlePrefix} =
            $options->{titlePrefix} eq 'revision'
          ? $options->{revisionText}
          : $options->{titlePrefix}
          if $options->{titlePrefix};
        $model->{localTime} = \@localTime;
        $SpreadsheetModel::ShowDimensions = $options->{showDimensions}
          if $options->{showDimensions};
        $options->{logger} = new SpreadsheetModel::Logger(
            name            => '',
            showFinalTables => $model->{forwardLinks},
            showDetails     => $model->{debug},
        );

        my $canPriority = $model->can('sheetPriority');
        my @pairs       = $model->worksheetsAndClosures($wbook);
        $options->{wsheetRunOrder} = [];
        while ( ( local $_, my $closure ) = splice @pairs, 0, 2 ) {
            my $priority = $canPriority ? $model->sheetPriority($_)
              || 0 : /^(?:Index|Overview)$/is ? 1 : 0;
            my $fullName = $_ . $modelCount;
            $sheetDisplayName{$fullName} =
              m#(.*)/# ? $1 . $modelCount : /(.*)\$$/ ? $1 : $_ . $modelCount;
            push @{ $options->{wsheetRunOrder} }, $fullName;
            push @{ $wsheetShowOrder[$priority] }, $fullName;
            $allClosures{$fullName} = $closure;
            undef $wsheetActive{$_}
              if $options->{activeSheets} && /$options->{activeSheets}/;
            $wsheetPassword{$fullName} = $options->{password}
              if $options->{protect};
        }

    }

    my %wsheet;

    for ( my $i = $#wsheetShowOrder ; $i >= 0 ; --$i ) {
        my %byDisplayName;
        foreach ( @{ $wsheetShowOrder[$i] } ) {
            my $dn = $sheetDisplayName{$_};
            $wsheet{$_} = $byDisplayName{$dn} ||= $wbook->add_worksheet($dn);
        }
    }

    $wbook->{$_} = $wsheet{$_} foreach keys %wsheet;
    foreach ( 0 .. $#optionArray ) {
        my $options = $optionArray[$_];
        my $modelCount = $_ ? ".$_" : '';
        $modelCount = '.' . ( 1 + $_ ) if @optionArray > 1;
        $wbook->{dataSheet} = $wsheet{ 'Input' . $modelCount };
        delete $wbook->{highestAutoTableNumber};

        if ( $forwardLinkFindingRun[$_] ) {
            open my $h2, '>', '/dev/null';
            my $wb2 = $module->new($h2);
            $wb2->setFormats($options);
            $wb2->{findForwardLinks} = 1;
            my @wsheetsAndClosures2 =
              $forwardLinkFindingRun[$_]->worksheetsAndClosures($wb2);
            my %closures2 = @wsheetsAndClosures2;
            my @sheetNames2 = @wsheetsAndClosures2[ grep { !( $_ % 2 ) }
              0 .. $#wsheetsAndClosures2 ];
            $wb2->{$_} = $wb2->add_worksheet($_) foreach @sheetNames2;
            $closures2{$_}->( $wb2->{$_} )
              foreach grep { !/Overview|Index/i } @sheetNames2;
            $wb2->close;
        }

        $wbook->{$_} = $options->{$_} foreach grep { exists $options->{$_} } qw(
          copy
          debug
          forwardLinks
          linesAsComment
          logger
          mergedRanges
          noLinks
          rowHeight
          tolerateMisordering
          validation
        );

        foreach ( @{ $options->{wsheetRunOrder} } ) {
            my $ws = $wsheet{$_};
            delete $ws->{sheetNumber};
            delete $ws->{lastTableNumber};
            $allClosures{$_}->($ws);
            $ws->activate if exists $wsheetActive{$_};
            $ws->fit_to_pages( 1, 0 );
            $ws->hide_gridlines(2);
            $ws->protect( $wsheetPassword{$_}, $ws->{protectionOptions} )
              if exists $wsheetPassword{$_};
            $ws->set_footer("&F");
            $ws->set_header("&L&A&C&R&P of &N");
            $ws->set_paper(9);
            $ws->insert_image( 0, 0, $options->{watermarkFile} )
              if $options->{watermarkFile};
        }

        if ($exporter) {
            $exporter->setModel( $modelCount, $options );
            $exporter->$_() foreach @exports;
        }

    }

    $multiModelSharing->finish
      if UNIVERSAL::can( $multiModelSharing, 'finish' );

    $wbook->close;
    $closer->() if $closer;
    0;

}

1;
