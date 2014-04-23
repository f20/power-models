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

    # NB: if we do not want to rely on "exit", then do something like:
    #     $ENV{PATH} = '';
    #     eval { File::Temp::cleanup(); };
    #     exec '/bin/test';

    $pid;
}

sub _applyDataOverride {
    my ( $dataset, $override ) = @_;
    foreach my $itable ( keys %$override ) {
        for (
            my $icolumn = 1 ;
            $icolumn < @{ $override->{$itable} } ;
            ++$icolumn
          )
        {
            foreach my $irow ( keys %{ $override->{$itable}[$icolumn] } ) {
                $dataset->{$itable}[$icolumn]{$irow} =
                  $override->{$itable}[$icolumn]{$irow};
            }
        }
    }
    $dataset;
}

sub create {

    my ( $module, $fileName, $instructions, %settings ) = @_;
    my @optionArray =
      ref $instructions eq 'ARRAY' ? @$instructions : $instructions;
    my @localTime = localtime;
    my $tmpDir;
    my $streamMaker = $settings{streamMaker};
    $streamMaker ||= sub {
        $tmpDir = '~$tmp-' . $$ unless $^O =~ /win32/i;
        mkdir $tmpDir and chmod 0770, $tmpDir if $tmpDir;
        open my $handle, '>',
          $tmpDir ? catfile( $tmpDir, $fileName ) : $fileName;
        binmode $handle;
        $handle, sub {
            if ($tmpDir) {
                rename catfile( $tmpDir, $fileName ), $fileName;
                rmdir $tmpDir;
            }
        };
    };
    my ( $handle, $closer ) = $streamMaker->($fileName);
    my $wbook = $module->new($handle);
    $wbook->set_tempdir($tmpDir)
      if $tmpDir && $module !~ /xlsx/i;  # work around taint issue with IO::File
    $wbook->setFormats( $optionArray[0] );
    my @models;
    my ( %allClosures, @wsheetShowOrder, %wsheetActive, %wsheetPassword );
    my @forwardLinkFindingRun;
    my $multiModelSharing;

    foreach ( 0 .. $#optionArray ) {
        my $options = $optionArray[$_];
        if ( $options->{dataset} ) {
            $wbook->{noData} = !$options->{illustrative};
            if ( my $yaml = delete $options->{dataset}{yaml} )
            {                            # deferred parsing
                require YAML;
                $options->{dataset} = YAML::Load($yaml);
            }
            $options->{dataset} =
              _applyDataOverride( Storable::dclone( $options->{dataset} ),
                $options->{dataOverride} )
              if $options->{dataOverride} && $options->{dataset};
        }
        my $modelCount = $_ ? ".$_" : '';
        $modelCount = '.' . ( 1 + $_ ) if @optionArray > 1;
        $options->{PerlModule}
          ->setUpMultiModelSharing( \$multiModelSharing, $options,
            \@optionArray )
          if UNIVERSAL::can( $options->{PerlModule}, 'setUpMultiModelSharing' );
        my $model = $options->{PerlModule}->new(%$options);
        $forwardLinkFindingRun[$_] = $model if $options->{forwardLinks};
        $options->{revisionText} ||= '';
        0 and $wbook->{titlePrefix} ||= $options->{revisionText};
        $model->{localTime} = \@localTime;
        $SpreadsheetModel::ShowDimensions = $options->{showDimensions}
          if $options->{showDimensions};
        $options->{logger} = new SpreadsheetModel::Logger(
            name            => 'List of data tables',
            finalTablesBold => $model->{forwardLinks},
        );

        my $canPriority = $model->can('sheetPriority');
        my @pairs       = $model->worksheetsAndClosures($wbook);
        $options->{wsheetRunOrder} = [];
        while ( ( local $_, my $closure ) = splice @pairs, 0, 2 ) {
            my $priority = $canPriority ? $model->sheetPriority($_)
              || 0 : /^(?:Index|Overview)$/is ? 1 : 0;
            my $fullName = /(.*)\$$/ ? $1 : $_ . $modelCount;
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
        foreach ( @{ $wsheetShowOrder[$i] } ) {
            my $ws = $wbook->add_worksheet($_);
            $ws->activate if exists $wsheetActive{$_};
            $ws->set_paper(9);
            $ws->fit_to_pages( 1, 0 );
            $ws->set_header("&L&A&C&R&P of &N");
            $ws->set_footer("&F");
            $ws->hide_gridlines(2);
            $ws->protect( $wsheetPassword{$_} ) if exists $wsheetPassword{$_};
            $wsheet{$_} = $ws;
        }
    }

    $wbook->{$_} = $wsheet{$_} foreach keys %wsheet;
    foreach ( 0 .. $#optionArray ) {
        my $options = $optionArray[$_];
        my $modelCount = $_ ? ".$_" : '';
        $modelCount = '.' . ( 1 + $_ ) if @optionArray > 1;
        $wbook->{dataSheet} = $wsheet{ 'Input' . $modelCount };

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
          qw(copy debug forwardLinks logAll logger noLinks rowHeight validation);

        $allClosures{$_}->( $wsheet{$_} )
          foreach @{ $options->{wsheetRunOrder} };

        my $dumpLoc = $fileName;
        $dumpLoc =~ s/\.xlsx?$//i;
        $dumpLoc .= $modelCount;

        if ( $settings{ExportHtml} ) {
            require SpreadsheetModel::ExportHtml;
            mkdir $dumpLoc;
            chmod 0770, $dumpLoc;
            SpreadsheetModel::ExportHtml::writeHtml( $options->{logger},
                "$dumpLoc/" );
        }

        if ( $settings{ExportText} ) {
            require SpreadsheetModel::ExportText;
            SpreadsheetModel::ExportText::writeText( $options, "$dumpLoc-" );
        }

        if ( $settings{ExportGraphviz} ) {
            require SpreadsheetModel::ExportGraphviz;
            my $dir = "$dumpLoc-graphs";
            mkdir $dir;
            chmod 0770, $dir;
            SpreadsheetModel::ExportGraphviz::writeGraphs(
                $options->{logger}{objects},
                $wbook, "$dir/" );
        }

        if ( $settings{ExportYaml} || $settings{ExportPerl} ) {
            my @objects = grep { defined $_ } @{ $options->{logger}{objects} };
            my $objNames = join( "\n",
                $options->{logger}{realRows}
                ? @{ $options->{logger}{realRows} }
                : map { "$_->{name}" } @objects );
            my @coreObj =
              map { UNIVERSAL::can( $_, 'getCore' ) ? $_->getCore : "$_"; }
              @objects;
            if ( $settings{ExportYaml} ) {
                require YAML;
                open my $fh, '>', "$dumpLoc.$$";
                binmode $fh, ':utf8';
                print {$fh} YAML::Dump(
                    {
                        '.' => $objNames,
                        map { ( ref $_ ? $_->{name} : $_, $_ ); } @coreObj
                    }
                );
                close $fh;
                rename "$dumpLoc.$$", "$dumpLoc.yaml";
            }
            if ( $settings{ExportPerl} ) {
                require Data::Dumper;
                my %counter;
                local $_ =
                  Data::Dumper->new( [ $objNames, @coreObj ] )->Indent(1)
                  ->Names(
                    [
                        'tableNames',
                        map {
                            my $n =
                              ref $_
                              ? $_->{name}
                              : $_;
                            $n =~ s/[^a-z0-9]+/_/gi;
                            $n =~ s/^([0-9]+)[0-9]{2}/$1/s;
                            "t$n" . ( $counter{$n}++ ? "_$counter{$n}" : '' );
                        } @coreObj
                    ]
                  )->Dump;
                s/\\x\{([0-9a-f]+)\}/chr (hex ($1))/eg;
                open my $fh, '>', "$dumpLoc.$$";
                binmode $fh, ':utf8';
                print {$fh} $_;
                close $fh;
                rename "$dumpLoc.$$", "$dumpLoc.pl";
            }
        }

    }

    $multiModelSharing->finish
      if UNIVERSAL::can( $multiModelSharing, 'finish' );

    $wbook->close;
    $closer->();

}

sub writeColourCode {
    my ( $wbook, $wsheet, $row ) = @_;
    $row ||= 1;
    $wsheet->write_string(
        ++$row, 2,
        'Colour coding',
        $wbook->getFormat('thc')
    );
    $wsheet->write_string( ++$row, 2, 'Input data',
        $wbook->getFormat('0.000hard') );
    $wsheet->write_string(
        ++$row, 2,
        'Constant value',
        $wbook->getFormat('0.000con')
    );
    $wsheet->write_string(
        ++$row, 2,
        'Formula: calculation',
        $wbook->getFormat('0.000soft')
    );
    $wsheet->write_string(
        ++$row, 2,
        'Formula: copy',
        $wbook->getFormat('0.000copy')
    );
    $wsheet->write_string(
        ++$row, 2,
        'Void cell in input data table',
        $wbook->getFormat('unused')
    );
    $wsheet->write_string(
        ++$row, 2,
        'Void cell in other table',
        $wbook->getFormat('unavailable')
    );
    $wsheet->write_string(
        ++$row, 2,
        'Unlocked cell for notes',
        $wbook->getFormat('scribbles')
    );

    unless ( $wsheet->{nextFree} && $wsheet->{nextFree} > ++$row ) {
        $wsheet->{nextFree} = ++$row;
    }
}

sub writeColourCodeSimple {
    my ( $wbook, $wsheet, $row ) = @_;
    $row ||= 1;
    $wsheet->write_string(
        ++$row, 2,
        'Colour coding',
        $wbook->getFormat('thc')
    );
    $wsheet->write_string( ++$row, 2, 'Input data',
        $wbook->getFormat('0.000hard') );
    $wsheet->write_string( ++$row, 2, 'Calculation',
        $wbook->getFormat('0.000soft') );
    $wsheet->write_string(
        ++$row, 2,
        'Data from tariff model',
        $wbook->getFormat('0.000copy')
    );
}

1;
