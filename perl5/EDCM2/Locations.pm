package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2012 Energy Networks Association Limited and contributors. All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted
provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of
conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of
conditions and the following disclaimer in the documentation and/or other materials provided
with the distribution.

THIS SOFTWARE IS PROVIDED BY ENERGY NETWORKS ASSOCIATION LIMITED AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL ENERGY
NETWORKS ASSOCIATION LIMITED OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF
THE POSSIBILITY OF SUCH DAMAGE.

=cut

use warnings;
use strict;
use utf8;

use SpreadsheetModel::Shortcuts ':all';

sub transmissionExit {

    my ( $model, $transmissionExitCharges, $systemPeakLoad ) = @_;

    Arithmetic(
        name       => 'Average transmission exit cost £/kVA/year',
        arithmetic => '=IV1/IV2',
        arguments => { IV1 => $transmissionExitCharges, IV2 => $systemPeakLoad }
    );

}

sub preprocess {

    my ( $model, $locations, $a1d, $r1d, $a1g, $r1g, ) = @_;

    $model->{locationData} = [ Stack( sources => [$locations] ) ];

    push @{ $model->{locationData} }, my $maxkVA = Arithmetic(
        name          => 'Maximum demand scenario kVA',
        defaultFormat => '0soft',
        arithmetic    => $model->{method} =~ /LRIC/i ? '=SQRT(IV2^2+IV1^2)'
        : '=SQRT((IV2+IV5)^2+(IV1+IV6)^2)',
        arguments => {
            IV1 => $r1d,
            IV2 => $a1d,
            $model->{method} =~ /LRIC/i ? ()
            : (
                IV5 => $a1g,
                IV6 => $r1g,
            )
        }
    );

    push @{ $model->{locationData} }, my $rf1 = Arithmetic(
        name       => 'Reactive factor in maximum demand scenario',
        arithmetic => '=IF(IV92=0,0,0-(IV1+IV4)/IV2)',
        arguments  => {
            IV1  => $r1d,
            IV2  => $maxkVA,
            IV92 => $maxkVA,
            IV4  => $r1g,
        }
    ) unless $model->{method} =~ /LRIC/i;

    $maxkVA, $rf1;

}

sub charge1 {

    my (
        $model, $tariffLoc, $locations, $locParent,
        $c1,    $a1d,       $r1d,       $a1g,
        $r1g,   $maxkVA,    $rf1,
    ) = @_;

    if ( $model->{method} eq 'none' ) {

        push @{ $model->{matrixColumns} }, my $invpf1 = Constant(
            name => 'Inverse power factor, maximum demand (kVA/kW)',
            rows => $tariffLoc->{rows},
            data => [ map { 1 } @{ $tariffLoc->{rows}{list} } ],
        );

        return [ undef, undef, undef ], [ undef, $invpf1, undef ], [];

    }

    push @{ $model->{matrixColumns} }, my $locMatchA = Arithmetic(
        name          => 'Location',
        defaultFormat => 'thloc',
        arithmetic    => '=MATCH(IV1,IV5_IV6,0)',
        arguments     => {
            IV1     => $tariffLoc,
            IV5_IV6 => $locations,
        }
    ) if $tariffLoc;

    if ( $model->{method} =~ /LRIC/i ) {

        my @locMatch = ($locMatchA);
        my $last = $model->{linkedLoc} || 8;
        --$last;

        push @{ $model->{matrixColumns} }, $locMatch[$_] = Arithmetic(
            name          => "Linked location $_",
            defaultFormat => 'thloc',
            arithmetic    => '=MATCH(INDEX(IV7_IV8,IV1),IV5_IV6,0)',
            arguments     => {
                IV1     => $locMatch[ $_ - 1 ],
                IV5_IV6 => $locations,
                IV7_IV8 => $locParent,
            }
        ) foreach 1 .. $last;

        my @c1l = map {
            push @{ $model->{matrixColumns} }, my $ca = Arithmetic(
                name       => 'Local charge 1 £/kVA/year at ' . $_->{name},
                arithmetic => $model->{noNegative}
                ? '=IF(ISNUMBER(IV1),MAX(0,INDEX(IV53_IV54,IV52)),0)'
                : '=IF(ISNUMBER(IV1),INDEX(IV53_IV54,IV52),0)',
                arguments => {
                    IV1       => $_,
                    IV52      => $_,
                    IV53_IV54 => $c1->[0],
                }
            );
            $ca;
        } @locMatch;

        my @c1n = map {
            push @{ $model->{matrixColumns} }, my $ca = Arithmetic(
                name       => 'Network charge 1 £/kVA/year at ' . $_->{name},
                arithmetic => $model->{noNegative}
                ? '=IF(ISNUMBER(IV1),MAX(0,INDEX(IV53_IV54,IV52)),0)'
                : '=IF(ISNUMBER(IV1),INDEX(IV53_IV54,IV52),0)',
                arguments => {
                    IV1       => $_,
                    IV52      => $_,
                    IV53_IV54 => $c1->[1],
                }
            );
            $ca;
        } @locMatch;

        my @kVA1 = map {
            push @{ $model->{matrixColumns} }, my $ca = Arithmetic(
                name       => 'Maximum demand run kVA at ' . $_->{name},
                arithmetic => '=IF(ISNUMBER(IV1),INDEX(IV53_IV54,IV52),0)',
                arguments  => {
                    IV1       => $_,
                    IV52      => $_,
                    IV53_IV54 => $maxkVA,
                }
            );
            $ca;
        } @locMatch;

        push @{ $model->{matrixColumns} }, my $c1l = Arithmetic(
            name       => 'Average local charge 1 (£/kVA/year)',
            arithmetic => '=IF('
              . join( '+', map { 'IV' . ( 1 + $_ ) } 0 .. $last )
              . '=0,IF(COUNT('
              . join( ',', map { "IU8$_" } 0 .. $last ) . '),('
              . join( '+', map { "IU4$_" } 0 .. $last )
              . ')/COUNT('
              . join( ',', map { "IU9$_" } 0 .. $last )
              . '),0),('
              . join( '+', map { "IU2$_*IU3$_" } 0 .. $last ) . ')/('
              . join( '+', map { "IU1$_" } 0 .. $last ) . '))',
            arguments => {
                map {
                    'IV' . ( $_ + 1 ) => $kVA1[$_],
                      "IU1$_" => $kVA1[$_],
                      "IU2$_" => $kVA1[$_],
                      "IU3$_" => $c1l[$_],
                      "IU4$_" => $c1l[$_],
                      "IU8$_" => $locMatch[$_],
                      "IU9$_" => $locMatch[$_],
                } 0 .. $last
            }
        );

        push @{ $model->{matrixColumns} }, my $c1n = Arithmetic(
            name       => 'Average network charge 1 (£/kVA/year)',
            arithmetic => '=IF('
              . join( '+', map { 'IV' . ( 1 + $_ ) } 0 .. $last )
              . '=0,IF(COUNT('
              . join( ',', map { "IU8$_" } 0 .. $last ) . '),('
              . join( '+', map { "IU4$_" } 0 .. $last )
              . ')/COUNT('
              . join( ',', map { "IU9$_" } 0 .. $last )
              . '),0),('
              . join( '+', map { "IU2$_*IU3$_" } 0 .. $last ) . ')/('
              . join( '+', map { "IU1$_" } 0 .. $last ) . '))',
            arguments => {
                map {
                    'IV' . ( $_ + 1 ) => $kVA1[$_],
                      "IU1$_" => $kVA1[$_],
                      "IU2$_" => $kVA1[$_],
                      "IU3$_" => $c1n[$_],
                      "IU4$_" => $c1n[$_],
                      "IU8$_" => $locMatch[$_],
                      "IU9$_" => $locMatch[$_],
                } 0 .. $last
            }
        );

        push @{ $model->{matrixColumns} }, my $active1 = Arithmetic(
            name       => 'Total active power in maximum demand scenario (kW)',
            arithmetic => '=0-' . join(
                '-',
                map {
                        'IF(ISNUMBER(IV'
                      . ( $_ + 1 )
                      . "),INDEX(IU1${_}_IU2$_,IU3$_),0)"
                } 0 .. $last
            ),
            arguments => {
                map {
                    'IV' . ( $_ + 1 ) => $locMatch[$_],
                      "IU1${_}_IU2$_" => $a1d,
                      "IU3$_"         => $locMatch[$_],
                } 0 .. $last
            }
        );

        push @{ $model->{matrixColumns} }, my $invpf1 = Arithmetic(
            name       => 'Inverse power factor, maximum demand (kVA/kW)',
            arithmetic => '=IF(IV6=0,1,MAX(1,SQRT(IV9^2+('
              . join( '+',
                map { "IF(ISNUMBER(IU2$_),INDEX(IU4${_}_IU5$_,IU3$_),0)" }
                  0 .. $last )
              . ')^2)/IV1))',
            arguments => {
                IV1 => $active1,
                IV6 => $active1,
                IV9 => $active1,
                map {
                    (
                        "IU2$_"         => $locMatch[$_],
                        "IU3$_"         => $locMatch[$_],
                        "IU4${_}_IU5$_" => $r1d,
                      )
                } 0 .. $last
            }
        );

        return [ $c1l, $c1n, undef ], [ undef, $invpf1, undef ], [];

    }

    push @{ $model->{matrixColumns} }, my $locMatchB = Arithmetic(
        name          => 'Parent location',
        defaultFormat => 'thloc',
        arithmetic    => '=MATCH(INDEX(IV7_IV8,IV1),IV5_IV6,0)',
        arguments     => {
            IV1     => $locMatchA,
            IV5_IV6 => $locations,
            IV7_IV8 => $locParent,
        }
    ) if $tariffLoc;

    push @{ $model->{matrixColumns} }, my $locMatchC = Arithmetic(
        name          => 'Grandparent location',
        defaultFormat => 'thloc',
        arithmetic    => '=MATCH(INDEX(IV7_IV8,IV1),IV5_IV6,0)',
        arguments     => {
            IV1     => $locMatchB,
            IV5_IV6 => $locations,
            IV7_IV8 => $locParent,
        }
    ) if $tariffLoc;

    push @{ $model->{matrixColumns} }, my $ca1 = Arithmetic(
        name       => 'Location charge 1 £/kVA/year',
        arithmetic => $model->{noNegative}
        ? '=IF(ISNUMBER(IV1),MAX(0,INDEX(IV53_IV54,IV52)),0)'
        : '=IF(ISNUMBER(IV1),INDEX(IV53_IV54,IV52),0)',
        arguments => {
            IV1       => $locMatchA,
            IV52      => $locMatchA,
            IV53_IV54 => $c1,
        }
    ) if $tariffLoc;

    push @{ $model->{matrixColumns} }, my $cb1 = Arithmetic(
        name => (
            $model->{method} =~ /LRIC/i ? 'Linked location 1'
            : 'Parent location'
          )
          . ' charge 1 £/kVA/year',
        arithmetic => $model->{noNegative}
        ? '=IF(ISNUMBER(IV1),MAX(0,INDEX(IV53_IV54,IV52)),0)'
        : '=IF(ISNUMBER(IV1),INDEX(IV53_IV54,IV52),0)',
        arguments => {
            IV1       => $locMatchB,
            IV52      => $locMatchB,
            IV53_IV54 => $c1,
        }
    ) if $tariffLoc;

    push @{ $model->{matrixColumns} }, my $cc1 = Arithmetic(
        name       => ('Grandparent location') . ' charge 1 £/kVA/year',
        arithmetic => $model->{noNegative}
        ? '=IF(ISNUMBER(IV1),MAX(0,INDEX(IV53_IV54,IV52)),0)'
        : '=IF(ISNUMBER(IV1),INDEX(IV53_IV54,IV52),0)',
        arguments => {
            IV1       => $locMatchC,
            IV52      => $locMatchC,
            IV53_IV54 => $c1,
        }
    );

    push @{ $model->{matrixColumns} }, my $rfa1 = Arithmetic(
        name => 'Network group reactive factor, maximum demand (kVAr/kVA)',
        arithmetic => '=IF(ISNUMBER(IV1),INDEX(IV53_IV54,IV52),0)',
        arguments  => {
            IV1       => $locMatchA,
            IV52      => $locMatchA,
            IV53_IV54 => $rf1,
        }
    );

    push @{ $model->{matrixColumns} }, my $pfa1 = Arithmetic(
        name       => 'Network group power factor, maximum demand (kW/kVA)',
        arithmetic => '=SQRT(1-IV1^2)',
        arguments  => { IV1 => $rfa1 }
    );

    push @{ $model->{matrixColumns} }, my $rfb1 = Arithmetic(
        name       => 'Parent group reactive factor, maximum demand (kVAr/kVA)',
        arithmetic => '=IF(ISNUMBER(IV1),INDEX(IV53_IV54,IV52),0)',
        arguments  => {
            IV1       => $locMatchB,
            IV52      => $locMatchB,
            IV53_IV54 => $rf1,
        }
    );

    push @{ $model->{matrixColumns} }, my $pfb1 = Arithmetic(
        name       => 'Parent group power factor, maximum demand (kW/kVA)',
        arithmetic => '=SQRT(1-IV1^2)',
        arguments  => { IV1 => $rfb1 }
    );

    push @{ $model->{matrixColumns} }, my $rfc1 = Arithmetic(
        name => 'Grandparent group reactive factor, maximum demand (kVAr/kVA)',
        arithmetic => '=IF(ISNUMBER(IV1),INDEX(IV53_IV54,IV52),0)',
        arguments  => {
            IV1       => $locMatchC,
            IV52      => $locMatchC,
            IV53_IV54 => $rf1,
        }
    );

    push @{ $model->{matrixColumns} }, my $pfc1 = Arithmetic(
        name       => 'Grandparent group power factor, maximum demand (kW/kVA)',
        arithmetic => '=SQRT(1-IV1^2)',
        arguments  => { IV1 => $rfc1 }
    );

    push @{ $model->{matrixColumns} }, my $rft1 = Arithmetic(
        name       => 'Top level reactive factor, maximum demand (kVAr/kVA)',
        arithmetic => '=IF(ISNUMBER(IV3),IV6,IF(ISNUMBER(IV2),IV5,IV4))',
        arguments  => {
            IV1 => $locMatchA,
            IV2 => $locMatchB,
            IV3 => $locMatchC,
            IV4 => $rfa1,
            IV5 => $rfb1,
            IV6 => $rfc1,
        }
    );

    push @{ $model->{matrixColumns} }, my $pft1 = Arithmetic(
        name       => 'Top level power factor, maximum demand (kW/kVA)',
        arithmetic => '=SQRT(1-IV1^2)',
        arguments  => { IV1 => $rft1 }
    );

    [ $ca1, $cb1, $cc1 ], [ $pfa1, $pfb1, $pfc1 ], [ $rfa1, $rfb1, $rfc1 ];

}

1;
