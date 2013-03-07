package Chedam;

=head Copyright licence and disclaimer

Copyright 2013 Franck Latrémolière and others. All rights reserved.

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

sub toDot {
    my ($data) = @_;
    my $widthFactor;
    $widthFactor =
      ( $data->{MW_GSP_132kV} || 0 ) +
      ( $data->{MW_GSP_EHV}   || 0 ) +
      ( $data->{MW_GSP_HV}    || 0 );
    $widthFactor ||=
      ( $data->{MVA_GSP_132kV} || 0 ) +
      ( $data->{MVA_GSP_EHV}   || 0 ) +
      ( $data->{MVA_GSP_HV}    || 0 );
    $widthFactor = $widthFactor ? 40.0 / $widthFactor : 0.02;
    my $lineBreak = '\n';
    my $dotCode   = join "\n",
      'digraph {',
      'graph [nodesep=1.5,ranksep=1.5,size="6,8",lp="top",'
      . 'fontname="Arial",fontsize="24"'
      . ',label="'
      . $data->{name} . '"' . '];',
      'node [shape=ellipse,style=filled,fillcolor="#ccffcc",'
      . 'width=2,height=1,fontname="Arial",fontsize="24"];', (
        map {
            my $style = $data->{"km_$_"} ? '' : ',style="filled,dotted"';
            my $pretty =
              $data->{"km_$_"}
              ? ( ( /^[LH]V/ ? ' &mdash; ' : $lineBreak )
                . int( $data->{"km_$_"} + 0.5 )
                  . ' km' )
              : '';
            $pretty =~ s/([0-9]+)([0-9]{3})/$1,$2/;
            qq%Net_$_ [label="$_ network$pretty"$style];%;
        } qw(132kV EHV HV LV),
      ),
      'node [shape=rectangle,style=filled,fillcolor="#ffffcc"];',
      'edge [color="#ff6633"'
      . ( 0 ? ',arrowhead=none' : ',arrowsize="0.75"' )
      . '];', 'subgraph clusterGSPG {', 'graph [label=""];', (
        map {
            if ( $data->{"MW_GSP_$_"} ) {
                my $pretty1 =
                  $data->{"count_GSP_$_"}
                  ? ( 0.1 * int( 0.5 + 10 * $data->{"count_GSP_$_"} )
                      . $lineBreak )
                  : '';
                my $pretty2 = int( $data->{"MW_GSP_$_"} + 0.5 ) . ' MW';
                $pretty2 =~ s/([0-9]+)([0-9]{3})/$1,$2/;
                qq%GSP_$_ [label="${pretty1}GSP $_$lineBreak$pretty2"];%;
            }
            elsif ( $data->{"MVA_GSP_$_"} ) {
                my $pretty1 =
                  $data->{"count_GSP_$_"}
                  ? ( 0.1 * int( 0.5 + 10 * $data->{"count_GSP_$_"} )
                      . $lineBreak )
                  : '';
                my $pretty2 = int( $data->{"MVA_GSP_$_"} + 0.5 ) . ' MVA';
                $pretty2 =~ s/([0-9]+)([0-9]{3})/$1,$2/;
                qq%GSP_$_ [label="${pretty1}GSP $_$lineBreak$pretty2"];%;
            }
            else {
                qq%GSP_$_ [label="GSP $_${lineBreak} ",style="filled,dotted"];%;
            }
        } qw(132kV EHV HV)
      ),
      '}', (
        map {
            my $pretty3 = $_;
            $pretty3 =~ s#Pri_(.+)#$1/HV#;
            if ( $data->{"MW_$_"} ) {
                my $pretty1 =
                  $data->{"count_$_"}
                  ? ( 0.1 * int( 0.5 + 10 * $data->{"count_$_"} ) . $lineBreak )
                  : '';
                my $pretty2 = int( $data->{"MW_$_"} + 0.5 ) . ' MW';
                $pretty2 =~ s/([0-9]+)([0-9]{3})/$1,$2/;
                qq%$_ [label="$pretty1$pretty3$lineBreak$pretty2"];%;
            }
            elsif ( $data->{"MVA_$_"} ) {
                my $pretty1 =
                  $data->{"count_$_"}
                  ? ( 0.1 * int( 0.5 + 10 * $data->{"count_$_"} ) . $lineBreak )
                  : '';
                my $pretty2 = int( $data->{"MVA_$_"} + 0.5 ) . ' MVA';
                $pretty2 =~ s/([0-9]+)([0-9]{3})/$1,$2/;
                qq%$_ [label="$pretty1$pretty3$lineBreak$pretty2"];%;
            }
            else {
                qq%$_ [label="$pretty3${lineBreak} ",style="filled,dotted"];%;
            }
        } qw(BSP Pri_EHV Pri_132kV)
      ),
      (
        map {
            if ( $data->{"MW_$_"} ) {
                my $pretty1 =
                  $data->{"count_$_"}
                  ? ( 0.1 * int( 0.5 + 10 * $data->{"count_$_"} ) . $lineBreak )
                  : '';
                my $pretty2 = int( $data->{"MW_$_"} + 0.5 ) . ' MW';
                $pretty2 =~ s/([0-9]+)([0-9]{3})/$1,$2/;
                my $style =
                  $data->{"MW_$_"}
                  ? ( 'setlinewidth('
                      . ( $data->{"MW_$_"} * $widthFactor )
                      . ')' )
                  : 'dotted';
                (
                    qq%Site_$_ [label="$pretty1$_ sites$lineBreak$pretty2"];%,
                    qq%Net_$_->Site_$_ [style="$style"];%
                );
            }
            elsif ( $data->{"MVA_$_"} ) {
                my $pretty1 =
                  $data->{"count_$_"}
                  ? ( 0.1 * int( 0.5 + 10 * $data->{"count_$_"} ) . $lineBreak )
                  : '';
                my $pretty2 = int( $data->{"MVA_$_"} + 0.5 ) . ' MVA';
                $pretty2 =~ s/([0-9]+)([0-9]{3})/$1,$2/;
                my $style =
                  $data->{"MVA_$_"}
                  ? ( 'setlinewidth('
                      . ( $data->{"MVA_$_"} * $widthFactor )
                      . ')' )
                  : 'dotted';
                (
                    qq%Site_$_ [label="$pretty1$_ sites$lineBreak$pretty2"];%,
                    qq%Net_$_->Site_$_ [style="$style"];%
                );
            }
            else {
                ();
            }
        } qw(132kV EHV)
      ),
      defined $data->{"MW_Dis_Ground"} || defined $data->{"MVA_Dis_Ground"}
      ? (
        map {
            my $pretty3 = "$_ HV/LV";
            if ( $data->{"MVA_Dis_$_"} ) {
                my $pretty1 =
                  $data->{"count_Dis_$_"}
                  ? ( 0.1 * int( 0.5 + 10 * $data->{"count_Dis_$_"} )
                      . $lineBreak )
                  : '';
                my $pretty2 = int( $data->{"MVA_Dis_$_"} + 0.5 ) . ' MVA';
                $pretty2 =~ s/([0-9]+)([0-9]{3})/$1,$2/;
                qq%$_ [label="$pretty1$pretty3$lineBreak$pretty2"];%;
            }
            else {
                ();
            }
        } qw(Ground Pole)
      )
      : (
        q%Dis [label="HV/LV"];%,
        map {
            my $pretty3 = "HV/LV";
            if ( $data->{"MW_$_"} ) {
                my $pretty1 =
                  $data->{"count_$_"}
                  ? ( 0.1 * int( 0.5 + 10 * $data->{"count_$_"} ) . $lineBreak )
                  : '';
                my $pretty2 = int( $data->{"MW_$_"} + 0.5 ) . ' MW';
                $pretty2 =~ s/([0-9]+)([0-9]{3})/$1,$2/;
                qq%$_ [label="$pretty1$pretty3$lineBreak$pretty2"];%;
            }
            elsif ( $data->{"MVA_$_"} ) {
                my $pretty1 =
                  $data->{"count_$_"}
                  ? ( 0.1 * int( 0.5 + 10 * $data->{"count_$_"} ) . $lineBreak )
                  : '';
                my $pretty2 = int( $data->{"MVA_$_"} + 0.5 ) . ' MVA';
                $pretty2 =~ s/([0-9]+)([0-9]{3})/$1,$2/;
                qq%$_ [label="$pretty1$pretty3$lineBreak$pretty2"];%;
            }
            else {
                ();
            }
        } 'Dis',
      ),
      (
        map {
            my ( $from, $to, $name ) = @$_;
            $from .= ':s' unless $from =~ /Net/;
            $to   .= ':n' unless $to =~ /Net/;
            my $flow = $data->{"MW_$name"} || $data->{"MVA_$name"};
            my $style =
              $flow
              ? ( 'setlinewidth('
                  . ( 0.1 * int( 0.5 + 10 * $flow * $widthFactor ) )
                  . ')' )
              : 'dotted';
            my $weight = $flow ? int( 1 + $flow * $widthFactor * 0.1 ) : 1;
            qq%$from->$to [style="$style", weight="$weight"];%;
          }[ GSP_132kV => Net_132kV => GSP_132kV => ],
        [ GSP_EHV   => Net_EHV   => GSP_EHV   => ],
        [ GSP_HV    => Net_HV    => GSP_HV    => ],
        [ Net_132kV => BSP       => BSP       => ],
        [ Net_132kV => Pri_132kV => Pri_132kV => ],
        [ BSP       => Net_EHV   => BSP       => ],
        [ Net_EHV   => Pri_EHV   => Pri_EHV   => ],
        [ Pri_132kV => Net_HV    => Pri_132kV => ],
        [ Pri_EHV   => Net_HV    => Pri_EHV   => ],
        defined $data->{"MW_Dis_Ground"} || defined $data->{"MVA_Dis_Ground"}
        ? (
            [ Net_HV => Ground => Dis_Ground => ],
            [ Net_HV => Pole   => Dis_Pole   => ],
            [ Ground => Net_LV => Dis_Ground => ],
            [ Pole   => Net_LV => Dis_Pole   => ],
          )
        : ( [ Net_HV => Dis => Dis => ], [ Dis => Net_LV => Dis => ], )
      ),
      '}';

    if (undef) {
        require YAML;
        open my $fh, '>', "Data $data->{name}.yml";
        print {$fh} YAML::Dump($data);
    }

    $dotCode;

}

1;
