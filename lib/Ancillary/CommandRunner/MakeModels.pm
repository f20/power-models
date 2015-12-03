package Ancillary::CommandRunner;

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
    C_PERL5DIR => 0,
    C_HOMEDIR  => 1,
    C_FOLDER   => 2,
    C_LOG      => 3,
};

sub makeModels {

    my $self = shift;

    require Ancillary::Manufacturing;
    my $maker = Ancillary::Manufacturing->factory(
        validate => [
            $self->[C_PERL5DIR],
            grep { -d $_ } catdir( $self->[C_HOMEDIR], 'X_Revisions' )
        ]
    );

    if ( $^O !~ /win32/i ) {
        if ( my $threads = `sysctl -n hw.ncpu 2>/dev/null` || `nproc` ) {
            chomp $threads;
            $maker->{threads}->($threads);
        }
    }

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
                if (/^-+autocheck/is) {
                    require Compilation::DataExtraction;
                    $maker->{setting}->(
                        PostProcessing => _makePostProcessor(
                            $maker->{threads}->(),
                            Compilation::DataExtraction::checksumWriter(),
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
                my $d = catdir( $self->[C_PERL5DIR], $1 );
                if ( -d $d ) {
                    lib->import($d);
                }
                else {
                    die "Special lib $d not found";
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
            elsif (/^-+orange/is) { $maker->{setRule}->( colour => 'orange' ); }
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
                require Compilation::DataExtraction;
                $maker->{setting}->(
                    PostProcessing => _makePostProcessor(
                        $maker->{threads}->(),
                        Compilation::DataExtraction::databaseWriter(),
                        "convert$1"
                    )
                );
            }
            elsif (/^-+stats=?(.*)/is) {
                $maker->{setRule}->(
                    summary => 'statistics',
                    $1 ? ( statistics => $1 ) : (),
                );
            }
            elsif (/^-+template(?:=(.+))?/is) {
                $maker->{setRule}->( template => $1 || ( time . "-$$" ) );
            }
            elsif (/^-+([0-9]+)/is) {
                $maker->{threads}->($1);
            }
            elsif (/^-+xdata=(.*)/is) {
                if ($1) {
                    if ( open my $fh, '<', $1 ) {
                        binmode $fh, ':utf8';
                        local undef $/;
                        $maker->{xdata}->(<$fh>);
                    }
                    else {
                        $maker->{xdata}->($1);
                    }
                }
                else {
                    local undef $/;
                    print "Enter xdata:\n";
                    $maker->{xdata}->(<STDIN>);
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
                my $file = catfile( $self->[C_HOMEDIR], $_ );
                if ( -f $file ) {
                    $maker->{addFile}->( abs2rel($file) );
                }
                elsif ( my @list = eval "<$file>" ) {
                    $maker->{addFile}->( abs2rel($_) ) foreach @list;
                }
                else {
                    warn "Ignored argument: $_";
                }
            }
        }
    }

    my @files = $maker->{fileList}->();
    mkdir 'models.tmp' if @files > 1 and !-e 'models.tmp' and !-e '~$models';
    $maker->{run}->();

}

1;
