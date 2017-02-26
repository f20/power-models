package SpreadsheetModel::Book::Manufacturing;

=head Copyright licence and disclaimer

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
use YAML;
require SpreadsheetModel::Book::Validation;
use File::Spec::Functions qw(catfile);

sub factory {
    my ( $class, @factorySettings ) = @_;
    my $self = bless {}, $class;
    my $settings = {@factorySettings};
    my ( %ruleOverrides, %dataOverrides );
    my ( @rulesets, @datasets, %dataByDatasetName );
    my %rulesDataSettings;

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
        require SpreadsheetModel::Data::XdataParser;
        $xdataParser =
          SpreadsheetModel::Data::XdataParser->new( \%dataOverrides,
            $self->{setRule} );
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
            if ( $fileName =~ /\.(?:csv|dta)$/is ) {
                require SpreadsheetModel::Data::ParseInputData;
                SpreadsheetModel::Data::ParseInputData::parseCsvDtaInputData(
                    \%dataByDatasetName, $blob, $fileName );
                return;
            }
            local undef $/;
            binmode $blob, ':utf8';
            $blob = <$blob>;
        }

        my @objects;
        if ( ref $blob ) {
            @objects = $blob;
        }
        elsif ( $blob =~ /^---/s ) {
            @objects = length($blob) < 4_196 && $fileName !~ /^\+/s
              || defined $fileName
              && $fileName =~ /%/ ? Load($blob) : { yaml => $blob };
        }
        else {
            eval { @objects = _jsonMachine()->decode($blob); };
            warn "$fileName: $@" if $@;
        }

        my @remainingObjects = grep { ref $_ eq 'HASH' } @objects;
        while (@remainingObjects) {
            local $_ = shift @remainingObjects;
            if ( exists $_->{template} ) {
                push @rulesets, $_;
            }
            elsif ( defined $fileName
                && $fileName =~ /([^\\\/]*%[^\\\/]*)\.(?:yml|yaml|json)$/is )
            {
                $_->{template} = $1;
                if (@remainingObjects) {
                    $_->{'~datasetIllustrative'} = [@remainingObjects];
                    @remainingObjects = ();
                }
                push @rulesets, $_;
            }
            elsif (
                defined $fileName
                && $fileName =~ m#([0-9]+-[0-9]+[a-zA-Z0-9-]*)?
                        [/\\]*\+\.(?:yml|yaml|json)$#six
              )
            {
                $dataByDatasetName{'+'}{ $1 ? "-$1" : '' } = [ $_, $fileName ];
            }
            else {
                my $datasetName;
                if (
                    defined $fileName
                    && $fileName =~ m#([0-9]+-[0-9]+[a-zA-Z0-9-]*)?
                        [/\\+]*([^/\\+]+)\.(?:yml|yaml|json)$#six
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
                                        $fileName =~ s#.*/models/#…/models#,
                                        $fileName
                                    ),
                                    validation => eval {
                                        require Encode;
                                        require
                                          SpreadsheetModel::Book::Validation;
                                        SpreadsheetModel::Book::Validation::digestMachine
                                          ->add( Encode::encode_utf8($blob) )
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
                    dataset => $_,
                };
            }
        }

    };

    $self->{addFile} = sub {
        ( local $_ ) = @_;
        my $dh;
        if (/\.(ygz|ybz|bz2|gz)$/si) {
            local $_ = $_;
            s/'/'"'"'/g;
            unless ( open $dh, join ' ',
                ( $1 =~ /bz/ ? 'bzcat' : qw(gunzip -c) ),
                "'$_'", '|' )
            {
                warn "No such compressed file: $_\n";
                return;
            }
        }
        else {
            unless ( open $dh, '<', $_ ) {
                warn "No such file: $_\n";
                return;
            }
        }
        $processStream->( $dh, $_ );
    };

    # This applies rules overrides, loads relevant code,
    # and, where configured to do so,
    # attributes a revision number to the resulting rules.
    # Returns an array of rulesets, or nothing if there is a problem.
    my $validate = $self->{validate} = sub {
        my ( $validatedLibs, $dbString ) = @_;
        $validatedLibs = [ grep { -d $_; } $validatedLibs ]
          unless 'ARRAY' eq ref $validatedLibs;

        if (%ruleOverrides) {
            foreach (@rulesets) {
                $_->{template} .= '+' if $_->{template};
                $_ = { %$_, %ruleOverrides };
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
            $_->{protect} = 1 unless exists $_->{protect};
            $_->{validation} = 'lenientnomsg'
              unless exists $_->{validation};
        }

        require SpreadsheetModel::Book::WorkbookCreate;
        require SpreadsheetModel::Book::WorkbookFormats;

        my $sourceCodeDigest =
          SpreadsheetModel::Book::Validation::sourceCodeDigest($validatedLibs);

        my ($db);
        if ( $dbString && require SpreadsheetModel::Book::RevisionNumbering ) {
            $db = SpreadsheetModel::Book::RevisionNumbering->connect($dbString)
              or warn "Cannot connect to $dbString";
        }

        foreach (@rulesets) {
            $_->{'~codeValidation'} = $sourceCodeDigest;
            delete @{$_}{ 'revisionText', grep { /^\./s } keys %$_ };
            my $template = delete $_->{template};
            $_->{revisionText} = $db->revisionText( Dump($_) ) if $db;
            $_->{template} = $template if defined $template;
        }

        @rulesets;

    };

    my $scorer = sub {
        my ( $rule, $data ) = @_;
        my $score = 0;
        return -666 unless $rule->{PerlModule};
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

        return unless @rulesets && @datasets;

        if ( $settings->{dataMerge} ) {
            my %byDatasetName;
            foreach (@datasets) {
                push @{ $byDatasetName{ $_->{'~datasetName'} || $_ } }, $_;
            }
            @datasets = ();
            while ( my ( $name, $heap ) = each %byDatasetName ) {
                push @datasets,
                  {
                    '~datasetName' => $name,
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

        if ( $settings->{extraDataYears} ) {
            foreach ( @{ $settings->{extraDataYears} } ) {
                my ( $sourceYear, $targetYear, $dataset ) = @$_;
                foreach (@datasets) {
                    my $name = $_->{'~datasetName'} or next;
                    $name =~ s/$sourceYear/$targetYear/ or next;
                    push @datasets,
                      {
                        '~datasetName'   => $name,
                        '~datasetSource' => "Additional data year $targetYear",
                        dataset          => $dataset,
                      };
                }
            }
        }

        if (%dataOverrides) {
            my $overrides = {%dataOverrides};
            my $suffix    = '-' . delete $overrides->{hash};
            foreach (@datasets) {
                $_->{'~datasetOverride'} = $overrides;
                $_->{'~datasetName'} .= $suffix
                  if defined $_->{'~datasetName'};
            }
        }

        $validate->( @{ $settings->{validate} } ) if $settings->{validate};

        my $extension = $workbookModule->( $settings->{xls} )->fileExtension;

        my $addToList = sub {
            my ( $data, $rule ) = @_;
            my $spreadsheetFile = $rule->{template};
            $spreadsheetFile =~ s/^%-/%-$rule->{PerlModule}-/
              unless $spreadsheetFile =~ /$rule->{PerlModule}/;
            if ( $rule->{revisionText} ) {
                $spreadsheetFile .= '-' unless $spreadsheetFile =~ /[+-]$/s;
                $spreadsheetFile .= $rule->{revisionText};
            }
            $spreadsheetFile =~ s/%%/
                require SpreadsheetModel::Data::DnoAreas;
                SpreadsheetModel::Data::DnoAreas::normaliseDnoName(
                    $data->{'~datasetName'}=~m#(.*)-20[0-9]{2}-[0-9]+#
                );
              /eg;
            $spreadsheetFile =~ s/%/$data->{'~datasetName'}/g;
            $spreadsheetFile .= $extension;
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

        if ( $settings->{pickBestRules} ) {
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
        return keys %rulesDataSettings if wantarray;

    };

    $self->{run} = sub {
        my ( $executor, $progressReporter ) = @_;
        $progressReporter->( 0 + keys %rulesDataSettings )
          if $progressReporter;
        my $increment = %rulesDataSettings ? 1.0 / keys %rulesDataSettings : 0;
        my $progress = 0;
        while ( my ( $fileName, $rds ) = each %rulesDataSettings ) {
            $fileName = catfile( $rds->[2]{folder}, $fileName )
              if $rds->[2]{folder};
            my $rulesData    = _mergeRulesData( @{$rds}[ 0, 1 ] );
            my $module       = $workbookModule->( $rds->[2]{xls} );
            my $continuation = $rds->[2]{PostProcessing};
            if ($executor) {
                $executor->run( $module, 'create', $fileName,
                    [ $rulesData, $rds->[2] ],
                    $continuation, 1 );
            }
            else {
                warn "create $fileName\n";
                $module->create( $fileName, $rulesData, $rds->[2] );
                $continuation->($fileName) if $continuation;
                warn "finished $fileName\n";
            }
            $progressReporter->() if $progressReporter;
        }
        if ($executor) {
            if ( my $errorCount = $executor->complete ) {
                die(
                    (
                        $errorCount > 1
                        ? "$errorCount things have"
                        : 'Something has'
                    )
                    . ' gone wrong'
                );
            }
        }
    };

    $self;

}

sub parseXdata {
    my $self = shift;
    $self->{xdataParser}->()->parseXdata(@_);
}

my $_jsonMachine;

sub _jsonMachine {
    return $_jsonMachine if $_jsonMachine;
    foreach (qw(JSON JSON::PP)) {
        return $_jsonMachine = $_->new
          if eval "require $_";
    }
    die 'No JSON module';
}

sub _mergeRulesData {
    return [ map { _mergeRulesData(@$_); } @{ $_[1] } ]
      if !$_[0] && ref $_[1] eq 'ARRAY';
    my %options = map { %$_ } @_;
    my $extraNotice = delete $options{extraNotice};
    my @keys =
      grep { exists $options{$_}; }
      qw(
      password
      template
      dataset
      ~datasetOverride
    );
    my @removed = map { delete $options{$_}; } @keys;
    $options{$_} = '***'
      foreach grep { /^(?:password|\~datasetOverride)$/s; } @keys;
    $options{yaml} = Dump( \%options );

    if ( defined $extraNotice ) {
        $options{extraNotice} =
          'ARRAY' eq ref $extraNotice
          ? join( "\n", @$extraNotice )
          : $extraNotice;
    }
    for ( my $i = 0 ; $i < @keys ; ++$i ) {
        $options{ $keys[$i] } = $removed[$i];
    }
    \%options;
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
            warn <<EOW;
Cannot load $_$for:
$@
EOW
            return;
        }
    }
    1;
}

1;
