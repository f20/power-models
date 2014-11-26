package SpreadsheetModel::WorkbookCreate;

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

sub bgCreate {
    my ( $module, $fileName, @arguments ) = @_;
    my $pid = fork;
    return $pid, $fileName if $pid;
    $0 = "perl: $fileName";
    $module->create( $fileName, @arguments );
    exit 0 if defined $pid;

    # NB: if you need to avoid exit, then do something like:
    #     $ENV{PATH} = '';
    #     eval { File::Temp::cleanup(); };
    #     exec '/bin/test';

}

sub create {

    my ( $module, $fileName, $instructions, $settings ) = @_;
    my @optionArray =
      ref $instructions eq 'ARRAY' ? @$instructions : $instructions;
    my @localTime = localtime;
    my $tmpDir;
    my $streamMaker = $settings->{streamMaker};
    $streamMaker ||= sub {
        my ($fn) = @_;
        unless ($fn) {
            binmode STDOUT;
            return \*STDOUT;
        }
        $tmpDir = '~$tmp-' . $$ unless $^O =~ /win32/i;
        mkdir $tmpDir and chmod 0770, $tmpDir if $tmpDir;
        open my $handle, '>', $tmpDir ? catfile( $tmpDir, $fn ) : $fn;
        binmode $handle;
        $handle, sub {
            if ($tmpDir) {
                my $finalFile = $fn;
                if (
                    $fn !~ m#/#
                    and (
                        my ($folder) =
                        grep { -d $_ && -w _; } qw(~$models models.tmp)
                    )
                  )
                {
                    $finalFile = catfile( $folder, $fn );
                }
                rename catfile( $tmpDir, $fn ), $finalFile;
                rmdir $tmpDir;
            }
        };
    };
    my ( $handle, $closer ) = $streamMaker->($fileName);
    my $wbook = $module->new($handle);
    $wbook->set_tempdir($tmpDir)
      if $tmpDir && $module !~ /xlsx/i;  # work around taint issue with IO::File

    my @exports = grep { $settings->{$_} && /^Export/ } keys %$settings;
    my $exporter;
    if (@exports) {
        eval {
            require SpreadsheetModel::WorkbookExport;
            $exporter =
              SpreadsheetModel::WorkbookExport->new( $fileName, $wbook );
        };
        warn "@exports: $@" if $@;
    }

    $wbook->setFormats( $optionArray[0] );
    my @models;
    my ( %allClosures, @wsheetShowOrder, %wsheetActive, %wsheetPassword );
    my @forwardLinkFindingRun;
    my $multiModelSharing;
    my %sheetDisplayName;

    foreach my $i ( 0 .. $#optionArray ) {
        if ( my $dataset = $optionArray[$i]{dataset} ) {
            $wbook->{noData} = !$optionArray[$i]{illustrative};
            if ( my $yaml = $dataset->{yaml} ) {
                require YAML;    # for deferred parsing
                my $parsed = YAML::Load($yaml);
                %$dataset = %$parsed;
            }
            if ( my $prev = $optionArray[$i]{dataset}{baseDataset} ) {
                unless ( $prev > $i ) {
                    push @{ $optionArray[ $i - $prev ]{requestsToSeeModel} },
                      sub {
                        $optionArray[$i]{sourceModel} = $_[0];
                      };
                }
            }
            elsif ( my $overrides = $optionArray[$i]{dataOverride} ) {
                my $dataset = Storable::dclone( $optionArray[$i]{dataset} );
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
                $optionArray[$i]{dataset} = $dataset;
            }
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
        0 and $wbook->{titlePrefix} ||= $options->{revisionText};
        $model->{localTime} = \@localTime;
        $SpreadsheetModel::ShowDimensions = $options->{showDimensions}
          if $options->{showDimensions};
        $options->{logger} = new SpreadsheetModel::Logger(
            name            => 'List of data tables',
            finalTablesBold => $model->{forwardLinks},
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
              if $options->{protect} && !/^(?:Index|Overview)$/is;
        }

    }

    my %wsheet;

    for ( my $i = $#wsheetShowOrder ; $i >= 0 ; --$i ) {
        my %byDisplayName;
        foreach ( @{ $wsheetShowOrder[$i] } ) {
            my $dn = $sheetDisplayName{$_};
            my $ws = $byDisplayName{$dn};
            unless ($ws) {
                $ws = $wbook->add_worksheet($dn);
                $ws->set_paper(9);
                $ws->fit_to_pages( 1, 0 );
                $ws->set_header("&L&A&C&R&P of &N");
                $ws->set_footer("&F");
                $ws->hide_gridlines(2);
                $ws->protect( $wsheetPassword{$_} )
                  if exists $wsheetPassword{$_};
                $ws->activate if exists $wsheetActive{$_};
            }
            $wsheet{$_} = $byDisplayName{$dn} = $ws;
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

        $wbook->{$_} = $options->{$_}
          foreach grep { exists $options->{$_} }
          qw(copy debug forwardLinks hideFormulas logger noLinks rowHeight validation);

        foreach ( @{ $options->{wsheetRunOrder} } ) {
            delete $wsheet{$_}{sheetNumber};
            delete $wsheet{$_}{lastTableNumber};
            $allClosures{$_}->( $wsheet{$_} );
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

}

1;
