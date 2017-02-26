package SpreadsheetModel::CLI::CommandRunner;

=head Copyright licence and disclaimer

Copyright 2011-2017 Franck Latrémolière and others. All rights reserved.

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
use File::Glob qw(bsd_glob);
use File::Spec::Functions qw(abs2rel catdir catfile);
use Encode qw(decode_utf8);

use constant {
    C_HOMEDIR       => 0,
    C_VALIDATEDLIBS => 1,
};

sub makeModels {

    my $self = shift;

    require SpreadsheetModel::Book::Manufacturing;
    my $maker = SpreadsheetModel::Book::Manufacturing->factory(
        validate => [
            $self->[C_VALIDATEDLIBS],
            grep { -d $_ } catdir( $self->[C_HOMEDIR], 'X_Revisions' )
        ]
    );

    my $folder;
    my $setFolder = sub {
        my ($forceCreation) = @_;
        unless ( defined $folder ) {
            ($folder) = grep { -d $_ && -w _; } qw(~$models);
            if ( !defined $folder ) {
                require POSIX;
                $folder = POSIX::strftime( '%Y-%m-%d.tmp', localtime );
                if ( $forceCreation && !-e $folder ) {
                    warn "Created folder $folder to save models.\n";
                    mkdir $folder;
                }
                undef $folder unless -d $folder && -w _;
            }
        }
        if ( defined $folder ) {
            $maker->{setting}->( folder => $folder );
        }
    };

    my $executor;

    foreach ( map { decode_utf8 $_} @_ ) {
        if (/^-/s) {
            if ( $_ eq '-' ) {
                $maker->{processStream}->( \*STDIN );
            }
            elsif (/^-+(?:carp|confess)/is) {
                require Carp;
                $SIG{__DIE__} = \&Carp::confess;
            }
            elsif (/^-+(auto)?check/is) {
                $maker->{setRule}
                  ->( checksums => 'Line checksum 5; Table checksum 7' );
                if (/^-+autocheck(.*)/is) {
                    require SpreadsheetModel::Data::Autocheck;
                    $maker->{setting}->(
                        PostProcessing => makePostProcessor(
                            $1 ? "convert$1" : 'calc',
                            SpreadsheetModel::Data::Autocheck->new(
                                $self->[C_HOMEDIR]
                            )->makeWriterAndParserOptions,
                        )
                    );
                }
            }
            elsif (/^-+debug/is)   { $maker->{setRule}->( debug        => 1 ); }
            elsif (/^-+edcm/is)    { $maker->{setRule}->( edcmTables   => 1 ); }
            elsif (/^-+forward/is) { $maker->{setRule}->( forwardLinks => 1 ); }
            elsif (
                /^-+( graphviz|
                  html|
                  perl|
                  rtf|
                  te?xt|
                  tablemap|
                  ya?ml
                )/xis
              )
            {
                $maker->{setting}->( 'Export' . ucfirst( lc($1) ), 1 );
            }
            elsif (/^-+lib=(\S+)/is) {
                my @libs =
                  grep { -d $_; }
                  map { catdir( $_, $1 ); } @{ $self->[C_VALIDATEDLIBS] };
                if (@libs) {
                    lib->import(@libs);
                }
                else {
                    die "No lib found for $1";
                }
            }
            elsif (
                /^-+( numExtraLocations|
                  numExtraTariffs|
                  numLocations|
                  numSampleTariffs|
                  numTariffs
                )=([0-9]+)/xis
              )
            {
                $maker->{setRule}->( $1 => $2 );
            }
            elsif (/^-+tariffs=(.+)/is) {
                $maker->{setRule}->(
                    tariffs      => [ split /[^0-9]+/, $1 ],
                    vertical     => 1,
                    dataOverride => {
                        1190 => [ undef, { 'Enter TRUE or FALSE' => 'FALSE' } ]
                    },
                    ldnoRev => 0,
                );
            }
            elsif (/^-+orange/is) {
                $maker->{setRule}->( colour => 'orange' );
            }
            elsif (/^-+gold/is) {
                $maker->{setRule}->( colour => 'gold' );
            }
            elsif (/^-+illustrative/is) {
                $maker->{setRule}->( illustrative => 1 );
            }
            elsif (/^-+datamerge/is) {
                $maker->{setting}->( dataMerge => 1 );
            }
            elsif (/^-+extend=(20[0-9]{2})-(20[0-9]{2})/is) {
                $maker->{setting}->(
                    extraDataYears => [
                        map {
                            [
                                "-$1-02", "-$_-02",
                                {
                                    1000 => [
                                        {},
                                        {},
                                        {
                                            'Company charging year data version'
                                              => join( '/', $_, $_ + 1 ),
                                        },
                                        {
                                            'Company charging year data version'
                                              => 'Scenario',
                                        }
                                    ],
                                    baseDataset => 1,
                                }
                            ]
                        } $1 + 1 .. $2
                    ]
                );
            }
            elsif (/^-+pickall/is) {
                $maker->{setting}->( allowInconsistentRules => 1 );
            }
            elsif (/^-+pickbest/is) {
                $maker->{setting}->( pickBestRules => 1 );
            }
            elsif (/^-+password=(.+)/is) {
                $maker->{setRule}->( password => $1 );
            }
            elsif (/^-+password/is) {
                srand();
                $maker->{setRule}->( password => rand() );
            }
            elsif (/^-+(no|skip)protect/is) {
                $maker->{setRule}->( protect => 0 );
            }
            elsif (/^-+(right.*)/is) { $maker->{setRule}->( alignment => $1 ); }
            elsif (/^-+single/is) { $executor = 0; }
            elsif (/^-+([0-9]*)([tp])?$/is) {
                unless ($executor) {
                    if ( $2 ? $2 eq 't' : $^O =~ /win32/i ) {
                        require SpreadsheetModel::CLI::ExecutorThread;
                        $executor = SpreadsheetModel::CLI::ExecutorThread->new;
                    }
                    else {
                        require SpreadsheetModel::CLI::ExecutorFork;
                        $executor = SpreadsheetModel::CLI::ExecutorFork->new;
                    }
                }
                $executor->setThreads($1) if $1;
            }
            elsif (/^-+sqlite(.*)/is) {
                require SpreadsheetModel::Data::DataExtraction;
                $maker->{setting}->(
                    PostProcessing => makePostProcessor(
                        $1 ? "convert$1" : 'calc',
                        SpreadsheetModel::Data::DataExtraction::databaseWriter(
                        ),
                    )
                );
            }
            elsif (/^-+stats=?(.*)only/is) {
                $maker->{setRule}->( summary => $1 );
            }
            elsif (/^-+stats=?(.*)/is) {
                $maker->{setRule}
                  ->( summary => 'statistics' . ( $1 ? $1 : '' ), );
            }
            elsif (/^-+template(?:=(.+))?/is) {
                $maker->{setRule}->( template => $1 || ( time . '-' . $$ ) );
            }
            elsif (/^-+(?:folder|directory)=(.+)?/is) {
                $folder = $1;
            }
            elsif (/^-+xdata=?(.*)/is) {
                if ($1) {
                    if ( open my $fh, '<', $1 ) {
                        binmode $fh, ':utf8';
                        local undef $/;
                        $maker->parseXdata(<$fh>);
                    }
                    else {
                        $maker->parseXdata($1);
                    }
                }
                else {
                    local undef $/;
                    print "Enter xdata:\n";
                    $maker->parseXdata(<STDIN>);
                }
            }
            elsif (/^-+extraNotice=?(.*)/is) {
                if ($1) {
                    if ( open my $fh, '<', $1 ) {
                        binmode $fh, ':utf8';
                        local undef $/;
                        $maker->{setRule}->( extraNotice => <$fh> );
                    }
                    else {
                        $maker->{setRule}->( extraNotice => $1 );
                    }
                }
                else {
                    binmode STDIN, ':utf8';
                    local undef $/;
                    print "Enter extraNotice text:\n";
                    $maker->{setRule}->( extraNotice => <STDIN> );
                }
            }
            elsif (/^-+xls$/is)  { $maker->{setting}->( xls => 1 ); }
            elsif (/^-+xlsx$/is) { $maker->{setting}->( xls => 0 ); }
            elsif (/^-+new(data|rules|settings)/is) {
                $setFolder->(1);
                $maker->{fileList}->();
                $maker->{ 'reset' . ucfirst( lc($1) ) }->();
            }
            else {
                warn "Ignored option: $_\n";
            }
        }
        elsif ( -f $_ ) {
            $maker->{addFile}->( abs2rel($_) );
        }
        else {
            s/^\s+//s;
            s/\s+$//s;
            if ( -f $_ ) {
                $maker->{addFile}->( abs2rel($_) );
            }
            elsif (   # Windows cmd.exe does not expand globs, so we try it here
                my @list = grep { -f $_; } bsd_glob($_)
              )
            {
                $maker->{addFile}->( abs2rel($_) ) foreach @list;
            }
            else {
                my $file = catfile( $self->[C_HOMEDIR], 'models', $_ );
                if ( -f $file ) {
                    $maker->{addFile}->( abs2rel($file) );
                }
                elsif ( my @list = grep { -f $_; } bsd_glob($file) ) {
                    $maker->{addFile}->( abs2rel($_) ) foreach @list;
                }
                else {
                    warn "Ignored argument: $_";
                }
            }
        }
    }

    unless ( defined $executor ) {
        if ( $^O !~ /win32/i
            && eval 'require SpreadsheetModel::CLI::ExecutorFork' )
        {
            $executor = SpreadsheetModel::CLI::ExecutorFork->new;
        }
        elsif ( eval 'require SpreadsheetModel::CLI::ExecutorThread' ) {
            $executor = SpreadsheetModel::CLI::ExecutorThread->new;
        }
        else {
            warn "No multi-threading: $@";
        }
    }

    if ( my @files = $maker->{fileList}->() ) {
        $setFolder->( @files > 1 );
        my ( $progress, $total ) = ( 0, 0 );
        $maker->{run}->(
            $executor,
            sub {
                if ( $_[0] ) {

                  # Receive number of items to process.
                  # The progress bar is shown when starting to create each model
                  # so it deliberately never completes.
                    $total = 1 + $_[0];
                }
                elsif ( $total > 2 ) {
                    my $done = int( 76 * ++$progress / $total );
                    warn '|' x $done . '-' x ( 76 - $done ) . "\n";
                }
            }
        );
    }
    else {
        warn "Nothing to do.\n";
    }

}

1;
