package ModelM::MonsterBook;

=head Copyright licence and disclaimer

Copyright 2009-2011 The Competitive Networks Association and others. All rights reserved.

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
use POSIX ();
use SpreadsheetModel::Logger;
use File::Spec::Functions qw(catfile);
use SpreadsheetModel::Shortcuts ':all';
use ModelM::Master;
use ModelM::Sheets;
use ModelM::Options;

sub bgCreate {
    my ( $module, @arguments ) = @_;
    if ( my $pid = fork ) {
        return $pid, $arguments[0];
    }
    else {
        $0 = "perl: $arguments[0]";
        $module->create(@arguments);
        exit 0;
    }
}

sub create {
    my ( $module, $fileName, $options ) = @_;
    my @options   = @$options;
    my @localTime = localtime;
    $fileName = POSIX::strftime( '%Y-%m-%dT%H-%M-%S', @localTime ) . '.xls'
      unless $fileName && $fileName =~ /\.xlsx?$/i;
    my $actualModule =
      $fileName =~ /\.xlsx$/i
      ? 'SpreadsheetModel::WorkbookXLSX'
      : 'SpreadsheetModel::Workbook';
    my $tmpDir = '~$tmp-' . $$;
    mkdir $tmpDir;
    open my $handle, '>', catfile( $tmpDir, $fileName );
    my $wbook = $actualModule->new($handle);
    $wbook->set_tempdir($tmpDir)
      unless $actualModule =~ /xlsx/i;   # work around taint issue with IO::File
    $wbook->setFormats( $options[0] );
    my @models;
    my $optionsColumns;
    my @allCoreNames;
    my %allClosures;
    my $model;
    my $completeSharedOptionsSheet;

    if ( @options > 1 ) {                # Make the shared Controller
        $optionsColumns = ( bless $options[0], 'ModelM' )->allocationRules;
        my $ws = $wbook->add_worksheet('Options');
        $ws->fit_to_pages( 1, 0 );
        $ws->set_header("&L&A&R&P of &N");
        $ws->set_footer("&F");
        $ws->hide_gridlines(2);
        $ws->freeze_panes( 1, 0 );
        $ws->set_column( 0, 0,   48 );
        $ws->set_column( 1, 250, 20 );
        $_->wsWrite( $wbook, $ws ) foreach Notes(
            name => 'Controller'
              . (
                $options[0]{revisionText}
                ? " (version: $options[0]{revisionText})"
                : ''
              ),
            lines => [
                <<'EOL',

Copyright 2011 The Competitive Networks Association and contributors. All
rights reserved.

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
EOL
                <<EOL,

UNLESS STATED OTHERWISE, THIS WORKBOOK IS ONLY A PROTOTYPE FOR TESTING
PURPOSES AND ALL THE DATA IN THIS MODEL ARE FOR ILLUSTRATION ONLY.
EOL
            ]
        ), @$optionsColumns;
        $completeSharedOptionsSheet = sub {
            my $companies =
              Labelset( list => [ map { "=Input.$_!B8" } 1 .. @options ] );
            $_->wsWrite( $wbook, $ws ) foreach map {
                my ( $masterRow, $cols, $format, $tableName ) = @$_;
                Constant(
                    name          => "From $tableName",
                    defaultFormat => $format,
                    rows          => $companies,
                    cols          => Labelset(
                        list =>
                          [ map { "=Results.1!$_" . ($masterRow); } @$cols ]
                    ),
                    data => [
                        map {
                            my $col = $_;
                            [
                                map { "=Results.$_!$col" . ( $masterRow + 1 ); }
                                  1 .. @options
                            ];
                        } @$cols
                    ]
                );
              } (
                [
                    $model->{impactTables}[0]{columns}[0]{$wbook}{row},
                    [qw(B C D)], '%copy', $model->{impactTables}[0]{name},
                ],
                [
                    $model->{impactTables}[1]{columns}[0]{$wbook}{row},
                    [qw(B C)], '%copy', $model->{impactTables}[1]{name},
                ],
                [
                    $model->{impactTables}[2]{columns}[0]{$wbook}{row},
                    [qw(B C D E)], '%copy', $model->{impactTables}[2]{name},
                ],
                [
                    $model->{impactTables}[3]{columns}[0]{$wbook}{row},
                    [qw(B C D E)], '%softpm', $model->{impactTables}[3]{name},
                ],
              );
        };
    }

    foreach ( 0 .. $#options ) {
        my $options = $options[$_];
        $options->{optionsColumns} = $optionsColumns if $optionsColumns;
        my $modelCount = $_ ? ".$_" : '';
        $modelCount = '.' . ( 1 + $_ ) if @options > 1;
        $options->{showDimensions} =
          defined $SpreadsheetModel::ShowDimensions
          ? $SpreadsheetModel::ShowDimensions
          : 0
          unless exists $options->{showDimensions};
        delete $options->{inputData} if $options->{oneSheet};
        $wbook->{copy}       = $options->{copy};
        $wbook->{debug}      = $options->{debug};
        $wbook->{logAll}     = $options->{logAll};
        $wbook->{noLinks}    = $options->{noLinks};
        $wbook->{validation} = $options->{validation};
        $wbook->{noData}     = !$options->{illustrative};
        $wbook->{titlePrefix} ||= $options->{revisionText} ||= '';
        $options->{localTime}             = \@localTime;
        $SpreadsheetModel::ShowDimensions = $options->{showDimensions};
        $options->{logger}                = new SpreadsheetModel::Logger(
            name  => 'List of data tables in this workbook',
            lines => [
                'This table lists the data tables '
                  . '(inputs and calculations) in the model.  '
                  . 'The link is to the first data cell of each table.',
                '',
            ]
        );
        $model = new ModelM(%$options);
        my @wsheetsAndClosures = $model->worksheetsAndClosures($wbook);
        my %closure            = @wsheetsAndClosures;
        my @wsheetNames =
          @wsheetsAndClosures[ grep { !( $_ % 2 ) } 0 .. $#wsheetsAndClosures ];
        $options->{wsheetNames} = [ map { $_ . $modelCount } @wsheetNames ];
        @wsheetNames = (
            ( grep { /^(Overview|Options)/ } @wsheetNames ),
            ( grep { !/^(Overview|Options)/ } @wsheetNames )
        );
        push @allCoreNames, grep {
            my $nn = $_;
            !grep { $nn eq $_ } @allCoreNames
        } @wsheetNames;
        $allClosures{ $_ . $modelCount } = $closure{$_}
          foreach grep { $closure{$_} } @allCoreNames;
    }
    my %wsheet;
    foreach my $cn (@allCoreNames) {
        foreach ( 0 .. $#options ) {
            my $options = $options[$_];
            my $modelCount = $_ ? ".$_" : '';
            $modelCount = '.' . ( 1 + $_ ) if @options > 1;
            if ( $allClosures{ $cn . $modelCount } ) {
                my $ws;
                if ( ref $options->{oneSheet} ) { $ws = $options->{oneSheet}; }
                else {
                    $ws = $wbook->add_worksheet( $cn . $modelCount );
                    $ws->activate if $cn . $modelCount eq 'Results';
                    $ws->fit_to_pages( 1, 0 );
                    $ws->set_header("&L&A&C$options->{revisionText}&R&P of &N");
                    $ws->set_footer("&F");
                    $ws->hide_gridlines(2);
                    $ws->protect( $options->{password} )
                      if $options->{protect} && $cn !~ /^Overview/;
                    $options->{oneSheet} = $ws if $options->{oneSheet};
                }
                $wsheet{ $cn . $modelCount } = $ws;
            }
        }
    }
    $wbook->{$_} = $wsheet{$_} foreach keys %wsheet;
    foreach ( 0 .. $#options ) {
        my $options = $options[$_];
        my $modelCount = $_ ? ".$_" : '';
        $modelCount = '.' . ( 1 + $_ ) if @options > 1;
        if ( $options->{inputData} ) {
            $wbook->{inputSheet} = $wsheet{ 'Input' . $modelCount }
              if $options->{inputData} =~ /inputSheet/;
            $wbook->{dataSheet} = $wsheet{ 'Input' . $modelCount }
              if $options->{inputData} =~ /dataSheet/;
        }
        $wbook->{logger}          = $options->{logger};
        $wbook->{lastSheetNumber} = 13;
        $allClosures{$_}->( $wsheet{$_} ) foreach @{ $options->{wsheetNames} };
    }
    delete $wbook->{logger};
    $completeSharedOptionsSheet->() if $completeSharedOptionsSheet;
    $wbook->close;
    rename catfile( $tmpDir, $fileName ), $fileName;
    rmdir $tmpDir;
}

1;
