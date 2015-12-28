package Ancillary::Manufacturing;

=head Copyright licence and disclaimer

Copyright 2011-2015 Franck Latrémolière, Reckon LLP and others.

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
require Storable;
require Ancillary::Validation;
use File::Spec::Functions qw(catfile);

sub factory {
    my ( $class, %factorySettings ) = @_;
    my $self     = bless {}, $class;
    my $threads1 = 0;
    my %settings = %factorySettings;
    my ( %ruleOverrides, %dataOverrides );
    my ( @rulesets, @datasets, %deferredData );
    my %rulesDataSettings;

    $self->{resetSettings} = sub {
        %settings = %factorySettings;
    };

    $self->{resetRules} = sub {
        @rulesets      = ();
        %ruleOverrides = ();
    };

    $self->{resetData} = sub {
        @datasets      = ();
        %dataOverrides = ();
    };

    $self->{setting} = sub {
        %settings = ( %settings, @_ );
    };

    my $setRule = $self->{setRule} = sub {
        %ruleOverrides = ( %ruleOverrides, @_ );
    };

    my $applyDataOverride = sub {
        foreach ( grep { ref $_ eq 'HASH' } @_ ) {
            while ( my ( $tab, $dat ) = each %$_ ) {
                if ( ref $dat eq 'HASH' ) {
                    while ( my ( $row, $rd ) = each %$dat ) {
                        next unless ref $rd eq 'ARRAY';
                        for ( my $col = 0 ; $col < @$rd ; ++$col ) {
                            $dataOverrides{$tab}[ $col + 1 ]{$row} =
                              $rd->[$col];
                        }
                    }
                }
                elsif ( ref $dat eq 'ARRAY' ) {
                    for ( my $col = 0 ; $col < @$dat ; ++$col ) {
                        my $cd = $dat->[$col];
                        next unless ref $cd eq 'HASH';
                        while ( my ( $row, $v ) = each %$cd ) {
                            $dataOverrides{$tab}[$col]{$row} = $v;
                        }
                    }
                }
            }
        }
    };

    $self->{xdata} = sub {
        foreach (@_) {
            if (/^---\r?\n/s) {
                my @y = eval { Load $_; };
                warn $@ if $@;
                $applyDataOverride->(@y);
                next;
            }
            local $_ = $_;
            if (s/\{(.*)\}//s) {
                foreach ( grep { $_ } split /\}\s*\{/s, $1 ) {
                    my $d = _jsonMachine()->decode( '{' . $_ . '}' );
                    next unless ref $d eq 'HASH';
                    $setRule->(
                        map { %$_; } grep { $_; }
                          map  { delete $_[0]{$_}; }
                          grep { /^rules?$/is; }
                          keys %$d
                    );
                    $applyDataOverride->($d);
                }
            }
            while (s/(\S.*\|.*\S)//m) {
                my ( $tab, $col, @more ) = split /\|/, $1, -1;
                next unless $tab;
                if ( $tab =~ /^rules$/is ) {
                    $setRule->( $col, @more );
                }
                elsif ( @more == 1 ) {
                    $dataOverrides{$tab}{$col} = $more[0];
                }
                elsif (@more == 2
                    && $tab =~ /^[0-9]+$/s
                    && $col
                    && $col =~ /^[0-9]+$/s )
                {
                    $dataOverrides{$tab}[$col]{ $more[0] } = $more[1];
                }
            }
        }

        return unless %dataOverrides;
        my $key = rand();
        $dataOverrides{hash} = 'hashing-error';
        eval {
            my $digestMachine = Ancillary::Validation::digestMachine();
            $key = $digestMachine->add( Dump( \%dataOverrides ) )->digest;
            $dataOverrides{hash} =
              substr( $digestMachine->add($key)->hexdigest, 5, 8 );
        };
        warn "Data overrides hashing error: $@" if $@;
        $key;

    };

    my $processStream = $self->{processStream} = sub {

        my ( $blob, $fileName ) = @_;

        if ( $fileName =~ /\.dta$/is ) {    # $blob must be a file handle
            require Parse::Stata::DtaReader;
            warn "Reading $_ with Parse::Stata::DtaReader\n";
            my $dta = Parse::Stata::DtaReader->new($blob);
            my ( @table, @column );
            for ( my $i = 1 ; $i < $dta->{nvar} ; ++$i ) {
                if ( $dta->{varlist}[$i] =~ /t([0-9]+)c([0-9]+)/ ) {
                    $table[$i]  = $1;
                    $column[$i] = $2;
                }
                else {
                    $table[$i]  = $dta->{varlist}[$i];
                    $column[$i] = 0;
                }
            }
            while ( my @row = $dta->readRow ) {
                my $book = $row[0];
                my $line = 'Value';
                $line = $row[1] if $row[1] && $row[1] =~ /^[0-9]+$/s;
                $deferredData{$book}{ $table[$_] }[ $column[$_] ]{$line} =
                  $row[$_]
                  foreach grep { $table[$_] } 1 .. $#table;
            }
            return;
        }

        if ( ref $blob eq 'GLOB' ) {
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

        foreach ( grep { ref $_ eq 'HASH' } @objects ) {
            if ( exists $_->{template} ) {
                push @rulesets, $_;
            }
            elsif ( defined $fileName
                && $fileName =~ /([^\\\/]*%[^\\\/]*)\.(?:yml|yaml|json)$/is )
            {
                $_->{template} = $1;
                push @rulesets, $_;
            }
            elsif (
                defined $fileName
                && $fileName =~ m#([0-9]+-[0-9]+[a-zA-Z0-9-]*)?
                        [/\\]*\+\.(?:yml|yaml|json)$#six
              )
            {
                $deferredData{'+'}{ $1 ? "-$1" : '' } = [ $_, $fileName ];
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
                                    file       => $fileName,
                                    validation => eval {
                                        require Encode;
                                        Ancillary::Validation::digestMachine()
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

    my $validate = $self->{validate} = sub {
        my ( $perl5dir, $dbString ) = @_;

        if (%ruleOverrides) {
            foreach (@rulesets) {
                $_->{template} .= '+' if $_->{template};
                $_ = { %$_, %ruleOverrides };
            }
        }

        $settings{safetyCheck}->(@rulesets) if $settings{safetyCheck};
        foreach (@rulesets) {
            _loadModules( $_, "$_->{PerlModule}::Master" ) || return;
            $_->{PerlModule}->can('requiredModulesForRuleset')
              and _loadModules( $_,
                $_->{PerlModule}->requiredModulesForRuleset($_) )
              || return;
            $_->{protect} = 1 unless exists $_->{protect};
            $_->{validation} = 'lenientnomsg' unless exists $_->{validation};
        }

        require SpreadsheetModel::WorkbookCreate;
        require SpreadsheetModel::WorkbookFormats;

        my $sourceCodeDigest =
          Ancillary::Validation::sourceCodeDigest($perl5dir);

        my ($db);
        if ( $dbString && require Ancillary::RevisionNumbering ) {
            $db = Ancillary::RevisionNumbering->connect($dbString)
              or warn "Cannot connect to $dbString";
        }

        foreach (@rulesets) {
            $_->{'~codeValidation'} = $sourceCodeDigest;
            delete $_->{'.'};
            $_->{revisionText} = $db->revisionText( Dump($_) ) if $db;
        }

    };

    my $pickBestScorer = sub {
        my ( $metadata, $rule ) = @_;
        my $score = 0;
        $score += 9000 if $metadata->[0] eq $rule->{PerlModule};
        my $scoringModule = "$rule->{PerlModule}::PickBest";
        $score += $scoringModule->score( $rule, $metadata )
          if eval "require $scoringModule";
    };

    my ( $xlsModule, $xlsxModule, $workbookModule );

    $workbookModule = sub {
        if ( $_[0] ) {
            unless ($xlsModule) {
                eval {
                    local $SIG{__DIE__} = \&Carp::confess;
                    require SpreadsheetModel::Workbook;
                    $xlsModule = 'SpreadsheetModel::Workbook';
                };
                warn $@ if $@;
            }
            $xlsModule ||= $workbookModule->();
        }
        else {
            unless ($xlsxModule) {
                eval {
                    local $SIG{__DIE__} = \&Carp::confess;
                    require SpreadsheetModel::WorkbookXLSX;
                    $xlsxModule = 'SpreadsheetModel::WorkbookXLSX';
                };
                warn $@ if $@;
            }
            $xlsxModule ||= $workbookModule->(1);
        }
    };

    $self->{fileList} = sub {

        if (%deferredData) {
            while ( my ( $book, $data ) = each %deferredData ) {
                if ( $book eq '+' ) {
                    my %nameset;
                    require Ancillary::DnoAreas;
                    map {
                        undef
                          $nameset{ Ancillary::DnoAreas::normaliseDnoName(
                                $_->[0] ) }{ $_->[1]
                              || '' };
                      } grep { $_->[0] }
                      map    { [m#(.*)(-20[0-9]{2}-[0-9]+)#] }
                      grep { $_ } map { $_->{'~datasetName'} } @datasets;
                    foreach my $dno ( keys %nameset ) {
                        foreach my $suffix ( sort keys %$data ) {
                            next if exists $nameset{$dno}{$suffix};
                            local $_ = $data->{$suffix}[1];
                            s/\+\./+$dno./;
                            push @datasets,
                              {
                                '~datasetName'   => $dno . $suffix,
                                dataset          => $data->{$suffix}[0],
                                '~datasetSource' => { file => $_ },
                              };
                        }
                    }
                }
                else {
                    $book =~ s/(?:-LRIC|-LRICsplit|-FCP)?([+-]r[0-9]+)?$//is;
                    $data->{numTariffs} = 2;
                    my $blob     = Dump($data);
                    my $fileName = "$book.yml";
                    warn "Writing $book data\n";
                    open my $h, '>', $fileName . $$;
                    binmode $h, ':utf8';
                    print {$h} $blob;
                    close $h;
                    rename $fileName . $$, $fileName;
                    $processStream->( $blob, $fileName );
                }
            }
            %deferredData = ();
        }

        if (%dataOverrides) {
            my $suffix = '-' . delete $dataOverrides{hash};
            foreach (@datasets) {
                $_->{dataOverride2} = \%dataOverrides;
                $_->{'~datasetName'} .= $suffix
                  if defined $_->{'~datasetName'};
            }
        }

        $validate->( @{ $settings{validate} } ) if $settings{validate};

        my $extension = $workbookModule->( $settings{xls} )->fileExtension;
        my $addToList = sub {
            my ( $data, $rule ) = @_;
            my $spreadsheetFile = $rule->{template};
            $spreadsheetFile =~ s/-/-$rule->{PerlModule}-/
              unless $spreadsheetFile =~ /$rule->{PerlModule}/;
            if ( $rule->{revisionText} ) {
                $spreadsheetFile .= '-' unless $spreadsheetFile =~ /[+-]$/s;
                $spreadsheetFile .= $rule->{revisionText};
            }
            $spreadsheetFile =~ s/%%/
                require Ancillary::DnoAreas;
                Ancillary::DnoAreas::normaliseDnoName($data->{'~datasetName'}=~m#(.*)-20[0-9]{2}-[0-9]+#);
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
                        [ $rule, $data, \%settings ]
                    ]
                ];
            }
            else {
                $rulesDataSettings{$spreadsheetFile} =
                  [ $rule, $data, \%settings ];
            }
        };

        if ( $settings{dataMerge} ) {
            my @mergedData;
            my %byCompany;
            foreach my $data (@datasets) {
                if (   $data->{dataset}{yaml}
                    && $data->{'~datasetName'}
                    && $data->{'~datasetName'} =~ /^([A-Z-]*[A-Z])[ -]*(.*)/i )
                {
                    $byCompany{$1}{ $2 || 'undated' } = $data->{dataset}{yaml};
                }
                else {
                    push @mergedData, $data;
                }
            }
            foreach my $co ( keys %byCompany ) {
                my @k = sort keys %{ $byCompany{$co} };
                push @mergedData,
                  {
                    '~datasetName' => join( '-', $co, @k ),
                    dataset => { yaml => join '', @{ $byCompany{$co} }{@k} },
                  };
            }
            @datasets = @mergedData;
        }

        if ( $settings{pickBestRules} ) {
            foreach my $data (@datasets) {
                my @scored;
                my $metadata = [
                    $data->{'~datasetSource'} && $data->{'~datasetSource'}{file}
                    ? "FILLER/$data->{'~datasetSource'}{file}" =~
                      /([A-Z0-9-]+)\/(?:.*(20[0-9][0-9]-[0-9][0-9]))?.*/
                    : ''
                ];
                foreach my $rule (@rulesets) {
                    next if _notPossible( $rule, $data );
                    push @scored,
                      [ $rule, $pickBestScorer->( $metadata, $rule ) ];
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
                      unless _notPossible( $rule, $data );
                }
            }
        }
        return keys %rulesDataSettings if wantarray;
    };

    $self->{threads} = sub {
        my ($threads) = @_;
        $threads1 = $threads - 1
          if $threads && $threads > 0 && $threads < 1_000;
        1 + $threads1;
    };

    $self->{run} = sub {
        my @fileNames = keys %rulesDataSettings;
        my %instructionsSettings =
          map {
            $_ => [
                _mergeRulesData( @{ $rulesDataSettings{$_} }[ 0, 1 ] ),
                $rulesDataSettings{$_}[2]
            ];
          } @fileNames;
        my ($folder) =
          grep { -d $_ && -w _; } qw(models.tmp ~$models);

        if ( $threads1 && eval 'require Ancillary::ParallelRunning' ) {
            foreach (@fileNames) {
                Ancillary::ParallelRunning::waitanypid($threads1);
                Ancillary::ParallelRunning::backgroundrun(
                    $workbookModule->( $instructionsSettings{$_}[1]{xls} ),
                    'create',
                    defined $folder ? catfile( $folder, $_ ) : $_,
                    $instructionsSettings{$_},
                    $instructionsSettings{$_}[1]{PostProcessing}
                );
            }
            my $errorCount = Ancillary::ParallelRunning::waitanypid(0);
            die "Some ($errorCount) things have gone wrong" if $errorCount;
        }
        else {
            foreach (@fileNames) {
                $workbookModule->( $instructionsSettings{$_}[1]{xls} )
                  ->create( defined $folder ? catfile( $folder, $_ ) : $_,
                    , @{ $instructionsSettings{$_} } );
                $instructionsSettings{$_}[1]{PostProcessing}->($_)
                  if $instructionsSettings{$_}[1]{PostProcessing};
            }
        }
    };

    $self;

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
    my @keys =
      grep { exists $options{$_}; } qw(password dataset datasetOverride);
    my @removed = map { delete $options{$_}; } @keys;
    $options{password} = '***' if $keys[0] && $keys[0] eq 'password';
    $options{yaml} = Dump( \%options );
    for ( my $i = 0 ; $i < @keys ; ++$i ) {
        $options{ $keys[$i] } = $removed[$i];
    }
    \%options;
}

sub _notPossible {
    my ( $rule, $data ) = @_;
    $rule->{wantTables} && keys %{ $data->{dataset} }
      and grep {
              !$data->{dataset}{$_}
          and !$data->{dataset}{yaml}
          || $data->{dataset}{yaml} !~ /^$_:/m
      } split /\s+/, $rule->{wantTables};
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
              defined $ruleset->{template} ? " for $ruleset->{template}" : '';
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
