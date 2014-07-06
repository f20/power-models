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

sub factory {
    my ($class) = @_;
    my $self = bless {}, $class;
    my $threads1 = 2;
    my ( $workbookModule, @rulesets, @datasets, %files, @createOptions,
        %manufacturingSettings );

    $self->{processStream} = sub {
        my ( $blob, $fileName ) = @_;
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
            eval {
                require JSON;
                @objects = JSON::from_json($blob);
            };
            eval {
                require JSON::PP;
                require Encode;
                @objects = JSON::PP::decode_json( Encode::encode_utf8($blob) );
            } if $@;
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
                if ( defined $fileName
                    && $fileName =~
m#([0-9]+-[0-9]+[a-zA-Z0-9-]*)?[/\\]?([^/\\]+)\.(?:yml|yaml|json)$#si
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
                                        require Digest::SHA1;
                                        require Encode;
                                        Digest::SHA1::sha1_hex(
                                            Encode::encode_utf8($blob) );
                                    },
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

    $self->{useXLS} = sub {
        if ( eval 'require SpreadsheetModel::Workbook' ) {
            $workbookModule = 'SpreadsheetModel::Workbook';
        }
        else {
            warn 'Could not load SpreadsheetModel::Workbook';
            warn $@;
        }
        $self;
    };

    $self->{setSettings} = sub {
        %manufacturingSettings = ( %manufacturingSettings, @_ );
    };

    $self->{useXLSX} = sub {
        if ( eval 'require SpreadsheetModel::WorkbookXLSX' ) {
            $workbookModule = 'SpreadsheetModel::WorkbookXLSX';
        }
        else {
            warn 'Could not load SpreadsheetModel::WorkbookXLSX';
            warn $@;
        }
        $self;
    };

    $self->{overrideRules} = sub {
        my %override = @_;
        my $suffix = ( grep { !/^Export/ } keys %override ) ? '+' : '';
        foreach (@rulesets) {
            $_->{template} .= $suffix if $_->{template};
            $_ = { %$_, %override };
        }
    };

    $self->{overrideData} = sub {
        my $od;
        my $takeOutRules = sub {
            $self->{overrideRules}->( ref $_ eq 'ARRAY' ? @$_ : %$_ )
              foreach grep { $_; }
              map          { delete $_[0]{$_}; }
              grep         { /^rules?$/i; }
              keys %{ $_[0] };
        };
        foreach (@_) {
            if (s/\{(.*)\}//s) {
                foreach ( grep { $_ } split /\}\s*\{/s, $1 ) {
                    require JSON::PP;
                    my $d = JSON::PP::decode_json( '{' . $_ . '}' );
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
        my ( $key, $hash ) = ( rand(), 'error' );
        eval {
            require Digest::SHA1;
            require JSON::PP;
            $key =
              Digest::SHA1::sha1(
                JSON::PP->new->canonical(1)->utf8->encode($od) );
            $hash = substr( Digest::SHA1::sha1_hex($key), 5, 8 );
        };

        foreach (@datasets) {
            $_->{dataOverride} = $od;
            $_->{'~datasetName'} .= "-$hash" if defined $_->{'~datasetName'};
        }
        ( $key, $hash );
    };

    $self->{validate} = sub {
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

        $self->{useXLSX}->() unless $workbookModule;

        use Ancillary::Validation qw(sha1File sourceCodeSha1);
        my $sourceCodeSha1 = sourceCodeSha1($perl5dir);

        # Omitted from validation: this file, and anything "require"d below.
        delete $sourceCodeSha1->{'Ancillary/Manufacturing.pm'};

        my ($db);
        if ( $dbString && require Ancillary::RevisionNumbering ) {
            $db = Ancillary::RevisionNumbering->connect($dbString)
              or warn "Cannot connect to $dbString";
        }

        foreach (@rulesets) {
            $_->{'~codeValidation'} = $sourceCodeSha1;
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

    my $addToList = sub {
        my ( $data, $rule, $extras ) = @_;
        my $spreadsheetFile = $rule->{template};
        if ( $rule->{revisionText} ) {
            $spreadsheetFile .= '-' unless $spreadsheetFile =~ /[+-]$/s;
            $spreadsheetFile .= $rule->{revisionText};
        }
        $spreadsheetFile =~
          s/%%/($data->{'~datasetName'}=~m#(.*)-[0-9]{4}-[0-9]{2}#)[0]/eg;
        $spreadsheetFile =~ s/%/$data->{'~datasetName'}/g;
        if ( exists $manufacturingSettings{output} ) {
            $spreadsheetFile =~ tr/-/ /;
            $extras->{identification} = $spreadsheetFile;
            $spreadsheetFile = $manufacturingSettings{output} || '';
        }
        $spreadsheetFile .= eval { $workbookModule->fileExtension; }
          || ( $workbookModule =~ /xlsx/i ? '.xlsx' : '.xls' )
          if $spreadsheetFile;
        if ( $files{$spreadsheetFile} ) {
            $files{$spreadsheetFile} = [
                undef,
                [
                      $files{$spreadsheetFile}[0]
                    ? $files{$spreadsheetFile}
                    : @{ $files{$spreadsheetFile}[1] },
                    [ $rule, $data, $extras || () ]
                ]
            ];
        }
        else {
            $files{$spreadsheetFile} = [ $rule, $data, $extras || () ];
        }
    };

    $self->{fileList} = sub {
        if ( $manufacturingSettings{pickBestRules} ) {
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
        elsif ( $manufacturingSettings{groupByRule} ) {
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
        keys %files;
    };

    $self->{prepare} = sub {
        map {
            $files{$_}
              && ( $files{$_} = _merge( @{ $files{$_} } ) ) ? $_ : ();
        } @_;
    };

    $self->{addOptions} = sub {
        push @createOptions, @_;
    };

    $self->{run} = sub {
        foreach (@_) {
            $workbookModule->create( $_, $files{$_}, @createOptions );
            $manufacturingSettings{PostProcessing}->($_)
              if $manufacturingSettings{PostProcessing};
        }
    };

    $self->{setThreads} = sub {
        my ($threads) = @_;
        $threads1 = $threads - 1 if $threads > 0 && $threads < 1_000;
    };

    $self->{runParallel} = sub {
        require Ancillary::ParallelRunning or goto &{ $self->{run} };
        foreach (@_) {
            Ancillary::ParallelRunning::waitanypid($threads1);
            Ancillary::ParallelRunning::registerpid(
                $workbookModule->bgCreate( $_, $files{$_}, @createOptions ),
                $manufacturingSettings{PostProcessing} );
        }
        Ancillary::ParallelRunning::waitanypid(0);
    };

    $self;

}

sub _merge {
    if ( !$_[0] && ref $_[1] eq 'ARRAY' ) {
        my @result = map { _merge(@$_); } @{ $_[1] };
        return wantarray ? @result : \@result;
    }
    my %options = map { %$_ } @_;
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

sub _runInFolder {
    pipe local *IN, local *OUT;
    if ( my $pid = fork ) {
        close OUT;
        my @result = <IN>;
        chomp @result;
        waitpid $pid, 0;
        return "@_ says $?" if $? >> 8;
        return wantarray ? @result : $result[0];
    }
    open STDOUT, '>&OUT';
    close IN;
    chdir shift;
    exec @_;
    die "exec @_: $!";
}

1;
