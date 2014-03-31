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
    my @options;
    my ( $workbookModule, $fileExtension, @rulesets, @datasets, %files );

    my $processRuleset = sub {
        local $_ = $_[0];
        _loadModules( $_, "$_->{PerlModule}::Master" ) || return;
        $_->{PerlModule}->can('requiredModulesForRuleset')
          and
          _loadModules( $_, $_->{PerlModule}->requiredModulesForRuleset($_) )
          || return;
        $_->{protect} = 1 unless exists $_->{protect};
        $_->{validation} = 'lenientnomsg' unless exists $_->{validation};
        push @rulesets, $_;
    };

    $self->{processStream} = sub {
        my ( $fileHandle, $fileName ) = @_;
        binmode $fileHandle, ':utf8';
        local undef $/;
        my $blob    = <$fileHandle>;
        my @objects = ();
        if ( $blob =~ /^---/s ) {
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
                $processRuleset->($_);
            }
            elsif ( defined $fileName
                && $fileName =~ /([^\\\/]*%[^\\\/]*)\.(?:yml|yaml|json)$/is )
            {
                $_->{template} = $1;
                $processRuleset->($_);
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

    $self->{useXLSX} = sub {
        if ( eval 'require SpreadsheetModel::WorkbookXLSX' ) {
            $workbookModule ||= 'SpreadsheetModel::WorkbookXLSX';
            $fileExtension = '.xlsx';
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

        unless ( $workbookModule && $fileExtension ) {
            $workbookModule = 'SpreadsheetModel::Workbook';
            $fileExtension  = '.xls';
            require SpreadsheetModel::Workbook;
        }

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

        # Keep dataOverride, illustrative, template, version and suchlike.
        # The purpose of this revision number is to help find or produce rules
        # to reproduce the same model, not just to describe the modelling rules.
        # This also avoids a cloning operation.

    };

    $self->{fileList} = sub {
        foreach my $rule (@rulesets) {
            my @wantTables;
            @wantTables = split /\s+/, $rule->{wantTables}
              if $rule->{wantTables};
            foreach my $data (@datasets) {
                next
                  if keys %{ $data->{dataset} }
                  and grep {
                          !$data->{dataset}{$_}
                      and !$data->{dataset}{yaml}
                      || $data->{dataset}{yaml} !~ /^$_:/m
                  } @wantTables;
                my $spreadsheetFile = $rule->{template};
                if ( $rule->{revisionText} ) {
                    $spreadsheetFile .= '-' unless $spreadsheetFile =~ /[+-]$/s;
                    $spreadsheetFile .= $rule->{revisionText};
                }
                $spreadsheetFile =~ s/%/$data->{'~datasetName'}/;
                $spreadsheetFile .= $fileExtension;
                if ( $files{$spreadsheetFile} ) {
                    $files{$spreadsheetFile} = [
                        undef,
                        [
                              $files{$spreadsheetFile}[0]
                            ? $files{$spreadsheetFile}
                            : @{ $files{$spreadsheetFile}[1] },
                            [ $rule, $data ]
                        ]
                    ];
                }
                else {
                    $files{$spreadsheetFile} = [ $rule, $data ];
                }
            }
        }
        keys %files;
    };

    $self->{prepare} = sub {
        map {
            $files{$_}
              && ( $files{$_} = _mergeRuleData( @{ $files{$_} } ) ) ? $_ : ();
        } @_;
    };

    $self->{addOptions} = sub {
        push @options, @_;
    };

    $self->{run} = sub {
        $workbookModule->create( $_, $files{$_}, @options ) foreach @_;
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
                $workbookModule->bgCreate( $_, $files{$_}, @options ) );
        }
        Ancillary::ParallelRunning::waitanypid(0);
    };

    $self;

}

sub _mergeRuleData {
    my ( $rule, $data ) = @_;
    if ( !$rule && ref $data eq 'ARRAY' ) {
        my @result = map { _mergeRuleData(@$_); } @$data;
        return wantarray ? @result : \@result;
    }
    my %options = ( %$rule, %$data );
    my %opt = %options;
    delete $opt{$_} foreach qw(dataset datasetOverride);
    $opt{password} = "***" if $opt{password};
    $options{yaml} = YAML::Dump( \%opt );
    \%options;
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
