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
use Encode qw(decode_utf8);
use File::Spec::Functions qw(catfile);

sub ymlIndex {

    my ( $self, @args ) = @_;

    require YAML;
    require Digest::SHA;

    my $getData = grep { /data/i } @args;
    my @folders = grep { -d $_ } @args;
    my %rulesets;

    foreach my $folder ( @folders ? @folders : $self->[C_HOMEDIR] ) {
        local $_ = $folder;
        s/'/'"'"'/g;
        open LIST, qq^cd '$_'; find . -name '*.y*ml' |^;
        while (<LIST>) {
            chomp;
            next if m#/\~\$#;
            next unless s#^\./##s;
            next if !$getData && m#/Data-[0-9]{4}-[0-9]{2}#;
            my @a = YAML::LoadFile( catfile( $folder, $_ ) );
            if ( @a == 0 ) {
                warn "$_ contains no objects\n";
            }
            elsif ( @a == 1 ) {
                $rulesets{$_} = $a[0];
            }
            else {
                for ( my $no = 0 ; $no < @a ; ++$no ) {
                    $rulesets{"$_/$no"} = $a[$no];
                }
            }
        }
    }

    my %map;
    my %keys;
    my @rulesetList = sort keys %rulesets;
    foreach my $rulesetName (@rulesetList) {
        my $ruleset = $rulesets{$rulesetName};
        unless ( ref $ruleset eq 'HASH' ) {
            warn "The rules set $rulesetName is not a HASH";
            next;
        }
        while ( my ( $k, $v ) = each %$ruleset ) {
            next if $k eq '.';
            undef $keys{$k};
            my $fk = "$k?";
            my $fv = (
                 !defined $v ? '¿✗'
                : ref $v     ? '¿✓'
                  . ref($v) . '#'
                  . Digest::SHA::sha1_hex(
                    Encode::encode_utf8( YAML::Dump($v) )
                  )
                : "¿✓$v"
            );
            $map{$fk}{$fv}{hash} ||= { $k => $v };
            undef $map{$fk}{$fv}{useset}{$rulesetName};
        }
    }

    foreach my $k ( keys %keys ) {
        my $fk = "$k?";
        if ( my @miss = grep { !exists $rulesets{$_}{$k} } @rulesetList ) {
            my $fv = '¿✗';
            $map{$fk}{$fv}{hash} ||= {};
            undef $map{$fk}{$fv}{useset}{$_} foreach @miss;
        }
    }

    my @fKeys = sort keys %map;

    if ( @rulesetList < 64 ) {

        foreach my $fk (@fKeys) {
            foreach my $fv ( keys %{ $map{$fk} } ) {
                my $x = delete $map{$fk}{$fv}{useset};
                $map{$fk}{$fv}{usemap} = unpack 'q', pack 'b64',
                  join '', map { exists $x->{$_} ? 1 : 0; } @rulesetList;
            }
        }

      RESTART: for ( my $i1 = 0 ; $i1 < @fKeys ; ++$i1 ) {
            for ( my $i2 = $i1 + 1 ; $i2 < @fKeys ; ++$i2 ) {
                my $fk1 = $fKeys[$i1];
                my $fk2 = $fKeys[$i2];
                my %possible;
                my @val1 = sort keys %{ $map{$fk1} };
                my @val2 = sort keys %{ $map{$fk2} };
                foreach my $fv1 (@val1) {
                    foreach my $fv2 (@val2) {
                        if ( my $combo =
                            $map{$fk1}{$fv1}{usemap} &
                            $map{$fk2}{$fv2}{usemap} )
                        {
                            $possible{ $fv1 . $fv2 } = {
                                usemap => $combo,
                                hash   => {
                                    %{ $map{$fk1}{$fv1}{hash} },
                                    %{ $map{$fk2}{$fv2}{hash} }
                                }
                            };
                        }
                    }
                }
                my @possibleList = keys %possible;
                if ( @possibleList == 1 || @possibleList < @val1 + @val2 - 1 ) {
                    delete $map{$fk1};
                    delete $map{$fk2};
                    if ( @possibleList == 1 ) {
                        $map{''} = { '' => values %possible };
                    }
                    else {
                        $map{ $fk1 . $fk2 } = \%possible;
                    }
                    @fKeys = sort keys %map;
                    goto RESTART;
                }
            }
        }
    }

    my %key;

    if ( my $h = delete $map{''} ) {
        $h = $h->{''}{hash};
        $map{$_} = $h->{$_} foreach keys %$h;
    }
    my $menuCounter = 0;
    foreach my $fk (@fKeys) {
        next if $fk eq '';
        my $menuLetter    = chr( 64 + ++$menuCounter );
        my $optionCounter = 0;
        foreach my $fv ( sort keys %{ $map{$fk} } ) {
            my $usemap = delete $map{$fk}{$fv}{usemap};
            if ( my $h = delete $map{$fk}{$fv}{hash} ) {
                $map{$fk}{$fv}{$_} = $h->{$_} foreach keys %$h;
            }
            my $optionId = $menuLetter . '=' . chr( 96 + ++$optionCounter );
            $map{$fk}{$optionId} = delete $map{$fk}{$fv};
            if ($usemap) {
                my $mask = 1;
                foreach (@rulesetList) {
                    $key{$_} = $key{$_} ? "$key{$_}&$optionId" : $optionId
                      if $mask & $usemap;
                    $mask <<= 1;
                }
            }
        }
        $map{$menuLetter} = delete $map{$fk};
    }

    print YAML::Dump( \%map, [ map { "$key{$_} $_" } sort keys %key ] );

}

1;
