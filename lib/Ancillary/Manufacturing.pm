package Ancillary::Manufacturing;

=head Copyright licence and disclaimer

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
require Storable;
require YAML;
require Ancillary::Validation;

sub factory {
    my ( $class, %factorySettings ) = @_;
    my $self     = bless {}, $class;
    my $threads1 = 0;
    my %settings = %factorySettings;
    my ( @rulesets, %ruleOverrides, @datasets, @xdata, %dataAccumulator );
    my %rulesDataSettings;

    $self->{prepare} = sub {
        warn '{prepare} does nothing and is deprecated';
    };

    $self->{resetSettings} = sub {
        %settings = %factorySettings;
    };

    $self->{resetRules} = sub {
        @rulesets      = ();
        %ruleOverrides = ();
    };

    $self->{resetData} = sub {
        @datasets = ();
        @xdata    = ();
    };

    $self->{setting} = sub {
        %settings = ( %settings, @_ );
    };

    $self->{setRule} = sub {
        %ruleOverrides = ( %ruleOverrides, @_ );
    };

    $self->{xdata} = sub {
        push @xdata, @_;
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
            }
            while ( my @row = $dta->readRow ) {
                my $book = $row[0];
                next unless my $line = $table[1] ? 'Single-line CSV' : $row[1];
                $dataAccumulator{$book}{ $table[$_] }[ $column[$_] ]{$line} =
                  $row[$_]
                  foreach grep { $table[$_] } 1 .. $#table;
            }
            return;
        }

        if ( ref $blob eq 'GLOB' ) {
            binmode $blob, ':utf8';
            local undef $/;
            $blob = <$blob>;
        }

        my @objects;
        if ( ref $blob ) {
            @objects = $blob;
        }
        elsif ( $blob =~ /^---/s ) {
            @objects = length($blob) < 32_768
              || defined $fileName
              && $fileName =~ /%/ ? YAML::Load($blob) : { yaml => $blob };
        }
        else {
            eval { @objects = _jsonMachine()->decode($blob); };
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
            else {
                my $datasetName;
                if (
                    defined $fileName
                    && $fileName =~ m#([0-9]+-[0-9]+[a-zA-Z0-9-]*)?
                        [/\\]?([^/\\]+)\.(?:yml|yaml|json)$#six
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
                warn "Could not open file: $_";
                return;
            }
        }
        else {
            unless ( open $dh, '<', $_ ) {
                warn "Could not open file: $_";
                return;
            }
        }
        $processStream->( $dh, $_ );
    };

    my $overrideRules = $self->{overrideRules} = sub {
        my %override = @_;
        my $suffix = ( grep { !/^Export/ } keys %override ) ? '+' : '';
        foreach (@rulesets) {
            $_->{template} .= $suffix if $_->{template};
            $_ = { %$_, %override };
        }
    };

    my $overrideData = $self->{overrideData} =
      sub {    # this processes any kind of xdata, not just data overrides
        my $od;
        my $takeOutRules = sub {
            $overrideRules->( ref $_ eq 'ARRAY' ? @$_ : %$_ )
              foreach grep { $_; }
              map          { delete $_[0]{$_}; }
              grep         { /^rules?$/i; }
              keys %{ $_[0] };
        };
        foreach (@_) {
            if (s/\{(.*)\}//s) {
                foreach ( grep { $_ } split /\}\s*\{/s, $1 ) {
                    my $d = _jsonMachine()->decode( '{' . $_ . '}' );
                    next unless ref $d eq 'HASH';
                    $takeOutRules->($d);
                    while ( my ( $tab, $dat ) = each %$d ) {
                        if ( ref $dat eq 'HASH' ) {
                            while ( my ( $row, $rd ) = each %$dat ) {
                                next unless ref $rd eq 'ARRAY';
                                for ( my $col = 0 ; $col < @$rd ; ++$col ) {
                                    $od->{$tab}[ $col + 1 ]{$row} = $rd->[$col];
                                }
                            }
                        }
                        elsif ( ref $dat eq 'ARRAY' ) {
                            for ( my $col = 0 ; $col < @$dat ; ++$col ) {
                                my $cd = $dat->[$col];
                                next unless ref $cd eq 'HASH';
                                while ( my ( $row, $v ) = each %$cd ) {
                                    $od->{$tab}[$col]{$row} = $v;
                                }
                            }
                        }
                    }
                }
            }
            while (s/(\S.*\|.*\S)//m) {
                my ( $tab, $col, @more ) = split /\|/, $1, -1;
                if ( @more == 1 ) {
                    $od->{$tab}{$col} = $more[0];
                }
                if (   @more == 2
                    && $tab
                    && $col
                    && $tab =~ /^[0-9]+$/s
                    && $col =~ /^[0-9]+$/s )
                {
                    $od->{$tab}[$col]{ $more[0] } = $more[1];
                }
            }
            $takeOutRules->($od);
        }
        return unless $od && keys %$od;
        my ( $key, $hash ) = ( rand(), 'hashing-error' );
        eval {
            my $digestMachine = Ancillary::Validation::digestMachine();
            $key = $digestMachine->add( _jsonMachine()->encode($od) )->digest;
            $hash = substr( $digestMachine->add($key)->hexdigest, 5, 8 );
        };

        foreach (@datasets) {
            $_->{dataOverride} = $od;
            $_->{'~datasetName'} .= "-$hash" if defined $_->{'~datasetName'};
        }
        ( $key, $hash );
      };

    my $validate = $self->{validate} = sub {
        my ( $perl5dir, $dbString ) = @_;

        foreach (@rulesets) {
            die "$_->{PerlModule} looks unsafe"
              unless {
                CDCM      => 1,
                EDCM2     => 1,
                EUoS      => 1,
                ModelM    => 1,
                Quantiles => 1,
              }->{ $_->{PerlModule} };    # hack
            _loadModules( $_, "$_->{PerlModule}::Master" ) || return;
            $_->{PerlModule}->can('requiredModulesForRuleset')
              and _loadModules( $_,
                $_->{PerlModule}->requiredModulesForRuleset($_) )
              || return;
            $_->{protect} = 1 unless exists $_->{protect};
            $_->{validation} = 'lenientnomsg' unless exists $_->{validation};
        }

        require SpreadsheetModel::WorkbookCreate;

        my $sourceCodeDigest =
          Ancillary::Validation::sourceCodeDigest($perl5dir);

        # Omitted from validation: this file, and anything "require"d below.
        delete $sourceCodeDigest->{'Ancillary/Manufacturing.pm'};

        my ($db);
        if ( $dbString && require Ancillary::RevisionNumbering ) {
            $db = Ancillary::RevisionNumbering->connect($dbString)
              or warn "Cannot connect to $dbString";
        }

        foreach (@rulesets) {
            $_->{'~codeValidation'} = $sourceCodeDigest;
            delete $_->{'.'};
            $_->{revisionText} = $db->revisionText( YAML::Dump($_) ) if $db;
        }

    };

    my $scorer = sub {
        my ( $metadata, $rule ) = @_;
        my $score = 0;
        $score += 9000 if $metadata->[0] eq $rule->{PerlModule};
        my $scoringModule = "$rule->{PerlModule}::Scoring";
        $score += $scoringModule->score( $rule, $metadata )
          if eval "require $scoringModule";
    };

    my ( $xlsModule, $xlsxModule, $workbookModule );

    $workbookModule = sub {
        return $xlsModule ||= eval {
            require SpreadsheetModel::Workbook;
            'SpreadsheetModel::Workbook';
        } || $workbookModule->() if $_[0];
        $xlsxModule ||= eval {
            require SpreadsheetModel::WorkbookXLSX;
            'SpreadsheetModel::WorkbookXLSX';
        } || $workbookModule->(1);
    };

    $self->{fileList} = sub {

        if (%dataAccumulator) {
            while ( my ( $book, $data ) = each %dataAccumulator ) {
                $book =~ s/(?:-LRIC|-LRICsplit|-FCP)?([+-]r[0-9]+)?$//is;
                $data->{numTariffs} = 2;
                my $blob     = YAML::Dump($data);
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

        $overrideData->(@xdata)                 if @xdata;
        $overrideRules->(%ruleOverrides)        if %ruleOverrides;
        $validate->( @{ $settings{validate} } ) if $settings{validate};

        my $extension = $workbookModule->( $settings{xls} )->fileExtension;
        my $addToList = sub {
            my ( $data, $rule ) = @_;
            my $spreadsheetFile = $rule->{template};
            if ( $rule->{revisionText} ) {
                $spreadsheetFile .= '-' unless $spreadsheetFile =~ /[+-]$/s;
                $spreadsheetFile .= $rule->{revisionText};
            }
            $spreadsheetFile =~
              s/%%/($data->{'~datasetName'}=~m#(.*)-[0-9]{4}-[0-9]{2}#)[0]/eg;
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

        if ( $settings{pickBestRules} ) {
            foreach my $data (@datasets) {
                my @scored;
                my $metadata = [
                    $data->{'~datasetSource'} && $data->{'~datasetSource'}{file}
                    ? $data->{'~datasetSource'}{file} =~
                      /([A-Z0-9-]+)\/(?:.*(20[0-9][0-9]-[0-9][0-9]))?.*/
                    : ''
                ];
                foreach my $rule (@rulesets) {
                    next if _notPossible( $rule, $data );
                    push @scored, [ $rule, $scorer->( $metadata, $rule ) ];
                }
                if (@scored) {
                    @scored = sort { $b->[1] <=> $a->[1] } @scored;
                    $addToList->( $data, $scored[0][0] );
                }
            }
        }
        elsif ( $settings{groupByRule} ) {
            foreach my $rule (@rulesets) {
                foreach my $data (@datasets) {
                    $addToList->( $data, $rule )
                      unless _notPossible( $rule, $data );
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
                _merge( @{ $rulesDataSettings{$_} }[ 0, 1 ] ),
                $rulesDataSettings{$_}[2]
            ];
          } @fileNames;
        if ( $threads1 && eval 'require Ancillary::ParallelRunning' ) {
            foreach (@fileNames) {
                Ancillary::ParallelRunning::waitanypid($threads1);
                Ancillary::ParallelRunning::registerpid(
                    $workbookModule->( $instructionsSettings{$_}[1]{xls} )
                      ->bgCreate( $_, @{ $instructionsSettings{$_} } ),
                    $instructionsSettings{$_}[1]{PostProcessing}
                );
            }
            Ancillary::ParallelRunning::waitanypid(0);
        }
        else {
            foreach (@fileNames) {
                $workbookModule->( $instructionsSettings{$_}[1]{xls} )
                  ->create( $_, @{ $instructionsSettings{$_} } );
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
        return $_jsonMachine = $_->new->utf8->canonical(1) if eval "require $_";
    }
    die 'No JSON module';
}

sub _merge {
    return [ map { _merge(@$_); } @{ $_[1] } ]
      if !$_[0] && ref $_[1] eq 'ARRAY';
    my %options = map { %$_ } @_;
    $options{identification} ||= join ' ', map {
        if ( local $_ = $_ ) {
            tr/-/ /;
            s/ (20[0-9][0-9] [0-9][0-9])/\t$1/;
            $_;
        }
        else {
            ();
        }
    } @options{qw(~datasetName version)};
    my %opt = %options;
    delete $opt{$_} foreach qw(dataset datasetOverride);
    $opt{password} = "***" if $opt{password};
    $options{yaml} = YAML::Dump( \%opt );
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
