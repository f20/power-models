﻿package SpreadsheetModel::Book::Manufacturing;

# Copyright 2011-2025 Franck Latrémolière and others.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice,
# this list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
# this list of conditions and the following disclaimer in the documentation
# and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
# EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
# THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

use warnings;
use strict;
use utf8;
use YAML;
require SpreadsheetModel::Book::Validation;
use File::Spec::Functions qw(catfile);

sub factory {
    my ( $class, @factorySettings ) = @_;
    my $self     = bless {}, $class;
    my $settings = {@factorySettings};
    my ( @rulesets, @datasets, %dataByDatasetName, %ruleOverrides,
        %dataOverrides, %rulesDataSettings, %finalRulesDataSettings );

    $self->{resetSettings} = sub {
        $settings = {@factorySettings};
    };

    $self->{resetRules} = sub {
        @rulesets      = ();
        %ruleOverrides = ();
    };

    $self->{resetData} = sub {
        @datasets          = ();
        %dataOverrides     = ();
        %dataByDatasetName = ();
    };

    $self->{setting} = sub {
        %$settings = ( %$settings, @_ );
    };

    $self->{setRule} = sub {
        %ruleOverrides = ( %ruleOverrides, @_ );
    };

    my $xdataParser;
    $self->{xdataParser} = sub {
        return $xdataParser if $xdataParser;
        require SpreadsheetModel::Data::ParseXdata;
        $xdataParser =
          SpreadsheetModel::Data::ParseXdata->new( \%dataOverrides,
            $self->{setRule}, \&jsonMachineMaker );
    };

    $self->{xdataKey} = sub {
        return unless %dataOverrides;
        my $key = rand();
        $dataOverrides{hash} = 'hashing-error';
        eval {
            my $digestMachine =
              SpreadsheetModel::Book::Validation::digestMachine();
            $key = $digestMachine->add( Dump( \%dataOverrides ) )->digest;
            $dataOverrides{hash} =
              substr( $digestMachine->add($key)->hexdigest, 5, 8 );
        };
        warn "Data overrides hashing error: $@" if $@;
        $key;
    };

    my $processStream = $self->{processStream} = sub {

        my ( $blob, $fileName ) = @_;

        if ( ref $blob eq 'GLOB' ) {    # file handle
            if ( $fileName =~ /\.[ct]sv$/is ) {
                require SpreadsheetModel::Data::ParseCsv;
                SpreadsheetModel::Data::ParseCsv::parseCsvInputData(
                    \%dataByDatasetName, $blob, $fileName );
                return;
            }
            local undef $/;
            binmode $blob, ':utf8';
            $blob = <$blob>;
            $/    = "\n";
        }

        my @objects;
        if ( ref $blob ) {
            @objects = $blob;
        }
        elsif ( $blob =~ /^---/s ) {
            @objects =
                 length($blob) < 4_096
              || defined $fileName && $fileName =~ /%/
              || $blob =~ /\n---/ ? Load($blob) : { yaml => $blob };
        }
        else {
            eval { @objects = jsonMachineMaker()->decode($blob); };
            warn "$fileName: $@" if $@;
        }

        my ( $templateDefinedFlag, $singleRuleset );
        while (@objects) {
            local $_ = shift @objects;
            next unless ref $_ eq 'HASH';
            if ( keys %$_ == 1 && $_->{rulesOverrides} ) {
                $self->{setRule}->( %{ $_->{rulesOverrides} } );
            }
            elsif ( defined $_->{template} ) {
                push @rulesets, $_;
                if ( $_->{template} eq '%' ) {
                    $singleRuleset = $_;
                }
                else {
                    $templateDefinedFlag = 1;
                }
            }
            elsif ( $templateDefinedFlag && defined $_->{nickName} ) {
                push @rulesets, $_;
            }
            elsif ( defined $fileName
                && $fileName =~ /([^\\\/]*%[^\\\/]*)\.(?:yml|yaml|json)$/is )
            {
                $_->{template} = $1;
                if (@objects) {
                    $_->{'~datasetIllustrative'} = [@objects];
                    @objects = ();
                }
                push @rulesets, $_;
            }
            else {
                my $datasetName;
                if (
                    defined $fileName
                    && $fileName =~ m#([0-9]+-[0-9]+[a-zA-Z0-9+-]*)?
                        [/\\]*([^/\\]+)\.(?:yml|yaml|json)$#six
                  )
                {
                    $datasetName = $2;
                    $datasetName .= "-$1" if $1;
                }
                push @datasets, {
                    defined $datasetName
                    ? (
                        '~datasetName' => $datasetName,
                        keys %$_
                        ? (
                            defined $fileName
                            ? (
                                '~datasetSource' => {
                                    file => scalar(
                                        $fileName =~ s#.*/models/#…/models/#,
                                        $fileName
                                    ),
                                    validation => eval {
                                        require Encode;
                                        require
                                          SpreadsheetModel::Book::Validation;
                                        SpreadsheetModel::Book::Validation::digestMachine(
                                        )->add( Encode::encode_utf8($blob) )
                                          ->hexdigest;
                                    }
                                      || 'Digest not working',
                                }
                              )
                            : ()
                          )
                        : ( '~datasetSource' => 'Empty dataset' ),
                      )
                    : (),
                    defined $singleRuleset
                    ? ( '~rulesRef' => $singleRuleset )
                    : (),
                    dataset => $_,
                };
            }
        }

    };

    my $addFile = $self->{addFile} = sub {
        my ($path) = @_;
        return
          if $settings->{fileFilter} && !$settings->{fileFilter}->($path);
        local $_ = $path;
        my $dh;
        if (/\.(ygz|ybz|bz2|gz)$/si) {
            s/'/'"'"'/g;
            unless ( open $dh, join ' ',
                ( $1 =~ /bz/ ? 'bzcat' : qw(gunzip -c) ),
                "'$_'", '|' )
            {
                warn "No such compressed file: $path\n";
                return;
            }
        }
        else {
            unless ( open $dh, '<', $path ) {
                warn "No such file: $path\n";
                return;
            }
        }
        $processStream->( $dh, $path );
    };

    $self->{addFolder} = sub {
        my ($path) = @_;
        my @datasetsStored = @datasets;
        @datasets = ();
        my $dirh;
        opendir $dirh, $path;
        $addFile->( catfile( $path, $_ ) )
          foreach grep { !/^\./s; } readdir $dirh;
        closedir $dirh;
        $path     = $1 if $path =~ m#([^/\\]+)#si;
        @datasets = (
            @datasetsStored,
            {
                '~datasetName' => $path,
                datasetArray   => [@datasets],
            }
        );
    };

    # This applies rules overrides, loads relevant code, and, if so
    # configured, attributes a revision number to the resulting rules.
    # Returns an array of rulesets (or nothing if there is a problem).
    my $validate = $self->{validate} = sub {
        my ( $validatedLibs, $revisionsDatabasePath ) =
          $settings->{validate} ? @{ $settings->{validate} } : [];
        $validatedLibs = [ grep { -d $_; } $validatedLibs ]
          unless 'ARRAY' eq ref $validatedLibs;

        if (%ruleOverrides) {
            foreach (@rulesets) {
                $_->{template} .= '+' if $_->{template};
                my %hash = ( %$_, %ruleOverrides );
                delete $hash{$_}
                  foreach grep { !defined $hash{$_}; } keys %hash;
                $_ = \%hash;
            }
        }

        $settings->{safetyCheck}->(@rulesets) if $settings->{safetyCheck};
        foreach (@rulesets) {
            return unless _loadModules( $_, "$_->{PerlModule}::Master" );
            if ( $_->{PerlModule}->can('requiredModulesForRuleset') ) {
                return
                  unless _loadModules( $_,
                    $_->{PerlModule}->requiredModulesForRuleset($_) );
            }
            $_->{protect}    = 1 unless exists $_->{protect};
            $_->{validation} = 'lenientnomsg'
              unless exists $_->{validation};
        }

        require SpreadsheetModel::Book::WorkbookCreate;
        require SpreadsheetModel::Book::WorkbookFormats;

        my $sourceCodeDigest =
          SpreadsheetModel::Book::Validation::sourceCodeDigest($validatedLibs);

        my ($db);
        if ( $revisionsDatabasePath
            && require SpreadsheetModel::Book::RevisionNumbering )
        {
            $db = SpreadsheetModel::Book::RevisionNumbering->connect(
                $revisionsDatabasePath)
              or warn "Cannot connect to $revisionsDatabasePath";
        }

        foreach (@rulesets) {
            $_->{'~codeValidation'} = $sourceCodeDigest;
            delete @{$_}{ 'revisionText', grep { /^\./s } keys %$_ };
            my $template = delete $_->{template};
            $_->{revisionText} = $db->revisionText( Dump($_) ) if $db;
            $_->{template}     = $template if defined $template;
        }

        @rulesets;

    };

    my $scorer = sub {
        my ( $rule, $data ) = @_;
        my $score = 0;
        return -666 unless $rule->{PerlModule};
        return -42
          if defined $data->{'~rulesRef'} && $data->{'~rulesRef'} != $rule;
        my $scoringModule = "$rule->{PerlModule}::PickBest";
        eval "require $scoringModule";
        $score += $scoringModule->score( $rule, $1 )
          if $scoringModule->can('score')
          and $data->{'~datasetName'}
          and $data->{'~datasetName'} =~ /(20[0-9]{2}-[0-9]{2})/;
        $score -= 1_000_000
          if $scoringModule->can('wantTables')
          and keys %{ $data->{dataset} }
          and grep {
                  !$data->{dataset}{$_}
              and !$data->{dataset}{yaml}
              || $data->{dataset}{yaml} !~ /^$_:/m
          } $scoringModule->wantTables($rule);
        $score;
    };

    my ( $xlsModule, $xlsxModule, $workbookModule );

    $workbookModule = sub {
        if ( $_[0] ) {
            unless ($xlsModule) {
                eval {
                    local $SIG{__DIE__} = \&Carp::confess;
                    require SpreadsheetModel::Book::WorkbookXLS;
                    $xlsModule = 'SpreadsheetModel::Book::WorkbookXLS';
                };
                warn $@ if $@;
            }
            $xlsModule ||= $workbookModule->();
        }
        else {
            unless ($xlsxModule) {
                eval {
                    local $SIG{__DIE__} = \&Carp::confess;
                    require SpreadsheetModel::Book::WorkbookXLSX;
                    $xlsxModule = 'SpreadsheetModel::Book::WorkbookXLSX';
                };
                warn $@ if $@;
            }
            $xlsxModule ||= $workbookModule->(1);
        }
    };

    $self->{fileList} = sub {

        if (%dataByDatasetName) {
            while ( my ( $book, $data ) = each %dataByDatasetName ) {
                $book =~ s/([+-]r[0-9]+)$//is;
                $data->{numTariffs} = 2
                  if $book =~ s/-(?:LRIC|LRICsplit|FCP)?$//is;
                push @datasets,
                  {
                    dataset        => $data,
                    '~datasetName' => $book,
                  };
                if ( !@rulesets || $settings->{dumpInputYaml} ) {
                    my $blob     = Dump($data);
                    my $fileName = "$book.yml";
                    warn "Writing $book data\n";
                    open my $h, '>', $fileName . $$;
                    binmode $h, ':utf8';
                    print {$h} $blob;
                    close $h;
                    rename $fileName . $$, $fileName;
                }
            }
            %dataByDatasetName = ();
        }

        return unless @rulesets;

        if ( $settings->{dataMerge} ) {
            my %byDatasetName;
            foreach (@datasets) {
                push @{ $byDatasetName{ $_->{'~datasetName'} || $_ } }, $_;
            }
            @datasets = ();
            foreach my $name ( sort keys %byDatasetName ) {
                my $heap = $byDatasetName{$name};
                push @datasets,
                  {
                    '~datasetName'   => $name,
                    '~datasetSource' =>
                      [ map { $_->{'~datasetSource'}; } @$heap ],
                    dataset => {
                        yaml => join '',
                        map { $_->{dataset}{yaml} || Dump( $_->{dataset} ); }
                          @$heap
                    },
                  };
            }
        }
        elsif ( !@datasets ) {
            @datasets = {
                dataset        => {},
                '~datasetName' => 'Blank',
            };
        }

        if (%dataOverrides) {
            my $overrides = {%dataOverrides};
            my $suffix =
              defined $overrides->{hash} ? delete $overrides->{hash} : '';
            foreach (@datasets) {
                $_->{'~datasetOverride'} = $overrides;
                $_->{'~datasetName'} .= $suffix
                  if defined $_->{'~datasetName'};
            }
        }

        $settings->{adjustDatasets}->( \@datasets )
          if $settings->{adjustDatasets};

        $validate->();

        my $extension = $workbookModule->( $settings->{xls} )->fileExtension;

        my $addToList = sub {
            my ( $data, $rule ) = @_;
            my $spreadsheetFile = $rule->{template};
            $spreadsheetFile =~ s/^%-/%-$rule->{PerlModule}-/
              unless $spreadsheetFile =~ /-$rule->{PerlModule}-/;
            $spreadsheetFile =~ s/%%/
                require SpreadsheetModel::Data::DnoAreas;
                SpreadsheetModel::Data::DnoAreas::normaliseDnoName(
                    $data->{'~datasetName'} =~ m#(.*)-20[0-9]{2}-[0-9]+#
                );
              /eg;
            $spreadsheetFile =~ s/%/$data->{'~datasetName'}/g;
            if ( $rulesDataSettings{$spreadsheetFile} ) {
                $rulesDataSettings{$spreadsheetFile} = [
                    undef,
                    [
                          $rulesDataSettings{$spreadsheetFile}[0]
                        ? $rulesDataSettings{$spreadsheetFile}
                        : @{ $rulesDataSettings{$spreadsheetFile}[1] },
                        [ $rule, $data, $settings ]
                    ],
                    $settings
                ];
            }
            else {
                $rulesDataSettings{$spreadsheetFile} =
                  [ $rule, $data, $settings ];
            }
        };

        if ( $settings->{customPicker} ) {
            $settings->{customPicker}->( $addToList, \@datasets, \@rulesets );
        }
        elsif ( $settings->{pickBestRules} ) {
            foreach my $data (@datasets) {
                my @scored;
                foreach my $rule (@rulesets) {
                    push @scored, [ $rule, $scorer->( $rule, $data ) ];
                }
                if (@scored) {
                    @scored = sort { $b->[1] <=> $a->[1] } @scored;
                    $addToList->( $data, $scored[0][0] );
                }
            }
        }
        else {
            foreach my $data (@datasets) {
                foreach my $rule (@rulesets) {
                    $addToList->( $data, $rule )
                      if $settings->{allowInconsistentRules}
                      || $scorer->( $rule, $data ) >= 0;
                }
            }
        }

        while ( my ( $file, $instructions ) = each %rulesDataSettings ) {
            my $spreadsheetFile = $file;
            if (
                my @revisionTexts =
                grep { $_; } map { $_->{revisionText}; } $instructions->[0]
                || map { $_->[0]; } @{ $instructions->[1] }
              )
            {
                $spreadsheetFile .= '-' unless $spreadsheetFile =~ /[+-]$/s;
                my $prev = '';
                my $count;
                foreach ( @revisionTexts, '' ) {
                    if ( $_ eq $prev ) {
                        ++$count;
                    }
                    else {
                        $spreadsheetFile .= $prev;
                        $spreadsheetFile .= 'x' . $count if $count > 1;
                        $prev  = $_;
                        $count = 1;
                    }
                }
            }
            $spreadsheetFile .= $extension;
            $finalRulesDataSettings{$spreadsheetFile} =
              $rulesDataSettings{$file};
        }

        keys %finalRulesDataSettings;

    };

    $self->{run} = sub {
        my ( $executor, $progressReporter ) = @_;
        $progressReporter->( 0 + keys %finalRulesDataSettings )
          if $progressReporter;
        my $increment =
          %finalRulesDataSettings ? 1.0 / keys %finalRulesDataSettings : 0;
        my $progress = 0;
        while ( my ( $fileName, $rds ) = each %finalRulesDataSettings ) {
            $fileName = catfile( $rds->[2]{folder}, $fileName )
              if $rds->[2]{folder};
            my $module       = $workbookModule->( $rds->[2]{xls} );
            my $continuation = $rds->[2]{PostProcessing};
            if ($executor) {
                $executor->run( $module, 'create', $fileName, $rds,
                    $continuation, 1 );
            }
            else {
                warn "create $fileName\n";
                $module->create( $fileName, @$rds );
                $continuation->($fileName) if $continuation;
                warn "finished $fileName\n";
            }
            $progressReporter->() if $progressReporter;
        }
        if ($executor) {
            if ( my @errors = $executor->complete ) {
                my $wrong = (
                      @errors > 1
                    ? @errors . " things have"
                    : 'Something has'
                ) . ' gone wrong.';
                warn "$wrong\n";
                warn sprintf( "%3d❗️ %s\n", $_ + 1, $errors[$_][0] )
                  foreach 0 .. $#errors;
                return 0 + @errors;
            }
        }
    };

    $self;

}

sub runAllWithFiles {
    my ( $self, @files ) = @_;
    $self->{addFile}->($_) foreach @files;
    $self->{fileList}->();
    $self->{run}->();
}

sub parseXdata {
    my $self = shift;
    $self->{xdataParser}->()->doParseXdata(@_);
}

my $jsonMachine;

sub jsonMachineMaker {
    return $jsonMachine if $jsonMachine;
    foreach (qw(JSON JSON::PP)) {
        return $jsonMachine = $_->new
          if eval "require $_";
    }
    die 'No JSON module';
}

sub _loadModules {
    my $ruleset = shift;
    return 1 unless @_;
    my %savedINC = %INC;
    foreach (@_) {
        eval "require $_";
        if ($@) {
            %INC = %savedINC;
            my $for =
              defined $ruleset->{template}
              ? " for $ruleset->{template}"
              : '';
            die <<EOW;
Cannot load $_$for:
$@
EOW
        }
    }
    1;
}

1;
