package SpreadsheetModel::CLI::CommandRunner;

=head Copyright licence and disclaimer

Copyright 2011-2015 Franck Latrémolière and others. All rights reserved.

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

use constant {
    C_HOMEDIR       => 0,
    C_VALIDATEDLIBS => 1,
    C_DESTINATION   => 2,
};

sub makeModels {

    my $self = shift;

    my $folder;

    require SpreadsheetModel::Book::Manufacturing;
    my $maker = SpreadsheetModel::Book::Manufacturing->factory(
        validate => [
            $self->[C_VALIDATEDLIBS],
            grep { -d $_ } catdir( $self->[C_HOMEDIR], 'X_Revisions' )
        ]
    );

    unless ( $^O =~ /win32/i ) {
        if ( my $threads = `sysctl -n hw.ncpu 2>/dev/null`
            || `nproc 2>/dev/null` )
        {
            chomp $threads;
            $maker->{threads}->($threads);
        }
    }

    foreach ( map { decode_utf8 $_} @_ ) {
        if (/^-/s) {
            if ( $_ eq '-' ) {
                $maker->{processStream}->( \*STDIN );
            }
            elsif (/^-autodata/i) {
                $maker->{setting}->( autoData => 1 );
                $maker->{setRule}->( template => '%' );
            }
            elsif (/^-+(?:carp|confess)/is) {
                require Carp;
                $SIG{__DIE__} = \&Carp::confess;
            }
            elsif (/^-+(auto)?check/is) {
                $maker->{setRule}
                  ->( checksums => 'Line checksum 5; Table checksum 7' );
                if (/^-+autocheck/is) {
                    require SpreadsheetModel::Data::DataExtraction;
                    $maker->{setting}->(
                        PostProcessing => _makePostProcessor(
                            $maker->{threads}->(),
                            SpreadsheetModel::Data::DataExtraction::checksumWriter(
                            ),
                            'convert'
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
                  text|
                  tablemap|
                  yaml
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
                srand();
                $maker->{setRule}->( colour => 'gold', password => rand() );
            }
            elsif (/^-+illustrative/is) {
                $maker->{setRule}->( illustrative => 1, );
            }
            elsif (/^-+datamerge/is) {
                $maker->{setting}->( dataMerge => 1 );
            }
            elsif (/^-+pickbest/is) {
                $maker->{setting}->( pickBestRules => 1 );
            }
            elsif (/^-+password=(.+)/is) {
                $maker->{setRule}->( password => $1 );
            }
            elsif (/^-+(no|skip)protect/is) {
                $maker->{setRule}->( protect => 0 );
            }
            elsif (/^-+(right.*)/is) { $maker->{setRule}->( alignment => $1 ); }
            elsif (/^-+single/is) { $maker->{threads}->(1); }
            elsif (/^-+(sqlite.*)/is) {
                require SpreadsheetModel::Data::DataExtraction;
                $maker->{setting}->(
                    PostProcessing => _makePostProcessor(
                        $maker->{threads}->(),
                        SpreadsheetModel::Data::DataExtraction::databaseWriter(
                        ),
                        "convert$1"
                    )
                );
            }
            elsif (/^-+stats=?(.*)/is) {
                $maker->{setRule}
                  ->( summary => 'statistics' . ( $1 ? $1 : '' ), );
            }
            elsif (/^-+template(?:=(.+))?/is) {
                $maker->{setRule}->( template => $1 || ( time . "-$$" ) );
            }
            elsif (/^-+(?:folder|directory)=(.+)?/is) {
                $folder = $1;
            }
            elsif (/^-+([0-9]+)/is) {
                $maker->{threads}->($1);
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
            elsif (/^-+xls$/is)  { $maker->{setting}->( xls => 1 ); }
            elsif (/^-+xlsx$/is) { $maker->{setting}->( xls => 0 ); }
            elsif (/^-+new(data|rules|settings)/is) {
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
            else {
                my $file = catfile( $self->[C_HOMEDIR], 'models', $_ );
                if ( -f $file ) {
                    $maker->{addFile}->( abs2rel($file) );
                }
                elsif ( my @list = grep { -f $_; } glob($file) ) {
                    $maker->{addFile}->( abs2rel($_) ) foreach @list;
                }
                else {
                    warn "Ignored argument: $_";
                }
            }
        }
    }

    if ( my @files = $maker->{fileList}->() ) {
        unless ( defined $folder ) {
            $folder = _temporaryFolder();
            if ( !defined $folder && @files > 1 ) {
                mkdir 'models.tmp';
                $folder = _temporaryFolder();
            }
        }
        if ( defined $folder ) {
            $maker->{setting}->( folder => $folder );
        }
        warn( ( @files > 1 ? ( @files . ' models' ) : 'One model' )
            . ' to be saved'
              . ( defined $folder ? " to $folder folder" : '' )
              . ".\n" );
        $maker->{run}->();
    }
    else {
        warn "Nothing to do.\n";
    }

}

sub _temporaryFolder {
    my ($folder) =
      grep { -d $_ && -w _; } qw(models.tmp ~$models);
    $folder;
}

1;
