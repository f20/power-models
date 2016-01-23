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

    my @folders = grep { -d $_ } @args;
    my @files = grep { /\.ya?ml$/is && -f $_ } @args;
    @folders = ( $self->[C_HOMEDIR] ) unless @folders + @files;
    my $name = join '; ', @folders, @files;

    foreach (@folders) {
        s/'/'"'"'/g;
        open LIST, qq^find '$_' -type f -name '%*.y*ml' |^;
        while (<LIST>) {
            chomp;
            push @files, $_;
        }
    }

    my %rulesets;

    foreach (@files) {
        my @a = YAML::LoadFile($_);
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

    my %map;

    my %keys;
    my @rulesetList = sort keys %rulesets;
    my $block       = 0;
    my $bit         = 1;
    foreach my $rulesetName (@rulesetList) {
        my $ruleset = $rulesets{$rulesetName};
        unless ( ref $ruleset eq 'HASH' ) {
            warn "The rules set $rulesetName is not a HASH";
            next;
        }
        while ( my ( $k, $v ) = each %$ruleset ) {
            next if $k eq '.' || $k eq 'nickName';
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
            $map{$fk}{$fv}{usemap}[$block] |= $bit;
        }
        unless ( $bit <<= 1 ) {
            ++$block;
            $bit = 1;
        }
    }

    foreach my $k ( keys %keys ) {
        my $fv    = '¿✗';
        my $fk    = "$k?";
        my $block = 0;
        my $bit   = 1;
        foreach (@rulesetList) {
            if ( !exists $rulesets{$_}{$k} ) {
                $map{$fk}{$fv}{hash} ||= {};
                $map{$fk}{$fv}{usemap}[$block] |= $bit;
            }
            unless ( $bit <<= 1 ) {
                ++$block;
                $bit = 1;
            }
        }
    }

    my @fKeys = sort keys %map;

    #   if ( @rulesetList < 64 ) {
    #       foreach my $fk (@fKeys) {
    #           foreach my $fv ( keys %{ $map{$fk} } ) {
    #               my $x = delete $map{$fk}{$fv}{useset};
    #               $map{$fk}{$fv}{usemap} = unpack 'q', pack 'b64',
    #                 join '', map { exists $x->{$_} ? 1 : 0; } @rulesetList;
    #           }
    #       }

  RESTART: for ( my $i1 = 0 ; $i1 < @fKeys ; ++$i1 ) {
        for ( my $i2 = $i1 + 1 ; $i2 < @fKeys ; ++$i2 ) {
            my $fk1 = $fKeys[$i1];
            my $fk2 = $fKeys[$i2];
            my %possible;
            my @val1 = sort keys %{ $map{$fk1} };
            my @val2 = sort keys %{ $map{$fk2} };
            next if @val1 == 1 && @val2 > 1 || @val2 == 1 && @val1 > 1;
            foreach my $fv1 (@val1) {
                foreach my $fv2 (@val2) {
                    my @intersection;
                    my $intersectionFlag;
                    for ( my $i = 0 ;
                        $i < @{ $map{$fk1}{$fv1}{usemap} } ; ++$i )
                    {
                        $intersectionFlag = 1
                          if $intersection[$i] =
                          ( $map{$fk1}{$fv1}{usemap}[$i] || 0 ) &
                          ( $map{$fk2}{$fv2}{usemap}[$i] || 0 );
                    }
                    if ($intersectionFlag) {
                        $possible{ $fv1 . $fv2 } = {
                            usemap => \@intersection,
                            hash   => {
                                %{ $map{$fk1}{$fv1}{hash} },
                                %{ $map{$fk2}{$fv2}{hash} }
                            }
                        };
                    }
                }
            }
            my @possibleList = keys %possible;
            my $test = @possibleList == 1 || @possibleList < @val1 + @val2 - 1;
            warn join ' : ', 0 + @val1, 0 + @val2, 0 + @possibleList,
              $test ? 'Yes' : 'No'
              if @possibleList > @val1
              && @possibleList > @val2
              && @possibleList < @val1 * @val2;
            if ($test) {
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

    my %key = map { ( $_ => [ $rulesets{$_}{'.'} ] ); } @rulesetList;

    if ( my $h = delete $map{''} ) {
        $h = $h->{''}{hash};
        $map{$_} = $h->{$_} foreach keys %$h;
    }
    my $menuCounter = 0;
    foreach my $fk (@fKeys) {
        next if $fk eq '';
        my $menuId =
          0.005 + 0.01 * ++$menuCounter;    # '.' . chr( 64 + ++$menuCounter );
        $menuId =~ s/^0//s;
        $menuId =~ s/5$//s;
        my $optionCounter = 0;
        my @options       = sort keys %{ $map{$fk} };
        my $alphaCounter  = @options < 26;
        foreach my $fv (@options) {
            my $usemap = delete $map{$fk}{$fv}{usemap};
            if ( my $h = delete $map{$fk}{$fv}{hash} ) {
                $map{$fk}{$fv}{$_} = $h->{$_} foreach keys %$h;
            }
            my $optionId =
              $alphaCounter
              ? chr( 96 + ++$optionCounter )
              : 0.0001 * ++$optionCounter;
            $map{$fk}{$optionId} = delete $map{$fk}{$fv};
            if ($usemap) {
                my $mask  = 1;
                my $index = 0;
                foreach (@rulesetList) {
                    push @{ $key{$_} }, $map{$fk}{$optionId}
                      if $usemap->[$index] && ( $mask & $usemap->[$index] );
                    unless ( $mask <<= 1 ) {
                        ++$index;
                        $mask = 1;
                    }
                }
            }
        }
        $map{$menuId} = delete $map{$fk};
    }

    $map{';'} = \%key;
    $map{'.'} = $name;

    print YAML::Dump( \%map );

}

1;
