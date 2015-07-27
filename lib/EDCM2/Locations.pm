package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2012 Energy Networks Association Limited and others.
Copyright 2015 Franck Latrémolière, Reckon LLP and others.

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

use SpreadsheetModel::Shortcuts ':all';

sub transmissionExit {

    my ( $model, $transmissionExitCharges, $systemPeakLoad ) = @_;

    Arithmetic(
        name       => 'Average transmission exit cost £/kVA/year',
        arithmetic => '=A1/A2',
        arguments => { A1 => $transmissionExitCharges, A2 => $systemPeakLoad }
    );

}

sub preprocessLocationData {

    my ( $model, $locations, $a1d, $r1d, $a1g, $r1g, ) = @_;

    return $model->{method} =~ /LRIC/i ? [ $a1d, $r1d ] : undef,
      $model->{method} =~ /LRIC/i ? undef : [ $a1d, $r1d, $a1g, $r1g ]
      unless $model->{legacy201};

    my @columns = Stack( sources => [$locations] );

    push @columns,
      my $maxkVA = Arithmetic(
        name          => 'Maximum demand scenario kVA',
        defaultFormat => '0soft',
        arithmetic    => $model->{method} =~ /LRIC/i ? '=SQRT(A2^2+A1^2)'
        : '=SQRT((A2+A5)^2+(A1+A6)^2)',
        arguments => {
            A1 => $r1d,
            A2 => $a1d,
            $model->{method} =~ /LRIC/i ? ()
            : (
                A5 => $a1g,
                A6 => $r1g,
            )
        }
      );

    push @columns,
      my $rf1 = Arithmetic(
        name       => 'Reactive factor in maximum demand scenario',
        arithmetic => '=IF(A92=0,0,0-(A1+A4)/A2)',
        arguments  => {
            A1  => $r1d,
            A2  => $maxkVA,
            A92 => $maxkVA,
            A4  => $r1g,
        }
      ) unless $model->{method} =~ /LRIC/i;

    $model->{locationTables} = [
        Columnset(
            name    => 'Preprocessing of location data',
            columns => \@columns,
        )
      ]
      unless $model->{legacy201};

    $maxkVA, $rf1;

}

sub charge1 {

    my (
        $model, $tariffLoc, $locations, $locParent,
        $c1,    $a1d,       $r1d,       $a1g,
        $r1g,   $maxkVA,    $rf1,
    ) = @_;

    if ( $model->{method} eq 'none' ) {

        my $invpf1 = Constant(
            name => 'Inverse power factor, maximum demand (kVA/kW)',
            rows => $tariffLoc->{rows},
            data => [ map { 1 } @{ $tariffLoc->{rows}{list} } ],
        );

        return [ undef, undef, undef ], [ undef, $invpf1, undef ], [];

    }

    my $locMatchA = Arithmetic(
        name          => 'Location',
        groupName     => 'Locations',
        newBlock      => 1,
        defaultFormat => 'locsoft',
        arithmetic    => '=MATCH(A1,A5_A6,0)',
        arguments     => {
            A1     => $tariffLoc,
            A5_A6 => $locations,
        }
    ) if $tariffLoc;

    if ( $model->{method} =~ /LRIC/i ) {

        my @locMatch = ($locMatchA);
        my $last = $model->{linkedLoc} || 8;
        --$last;

        $locMatch[$_] = Arithmetic(
            name          => "Linked location $_",
            defaultFormat => 'locsoft',
            arithmetic    => '=MATCH(INDEX(A7_A8,A1),A5_A6,0)',
            arguments     => {
                A1     => $locMatch[ $_ - 1 ],
                A5_A6 => $locations,
                A7_A8 => $locParent,
            }
        ) foreach 1 .. $last;

        my @c1l = map {

            my $ca = Arithmetic(
                name      => 'Local charge 1 £/kVA/year at ' . $_->{name},
                groupName => 'Local charge 1',
                $_ == $locMatchA ? ( newBlock => 1 ) : (),
                arithmetic => $model->{noNegative}
                ? '=IF(ISNUMBER(A1),MAX(0,INDEX(A53_A54,A52)),0)'
                : '=IF(ISNUMBER(A1),INDEX(A53_A54,A52),0)',
                arguments => {
                    A1       => $_,
                    A52      => $_,
                    A53_A54 => $c1->[0],
                }
            );
            $ca;
        } @locMatch;

        my @c1n = map {

            my $ca = Arithmetic(
                name      => 'Network charge 1 £/kVA/year at ' . $_->{name},
                groupName => 'Network charge 1',
                $_ == $locMatchA ? ( newBlock => 1 ) : (),
                arithmetic => $model->{noNegative}
                ? '=IF(ISNUMBER(A1),MAX(0,INDEX(A53_A54,A52)),0)'
                : '=IF(ISNUMBER(A1),INDEX(A53_A54,A52),0)',
                arguments => {
                    A1       => $_,
                    A52      => $_,
                    A53_A54 => $c1->[1],
                }
            );
            $ca;
        } @locMatch;

        my @kVA1 = map {

            my $ca =
              $model->{legacy201}
              ? Arithmetic(
                name          => 'Maximum demand run kVA at ' . $_->{name},
                groupName     => 'Maximum demand',
                defaultFormat => '0soft',
                arithmetic    => '=IF(ISNUMBER(A1),INDEX(A53_A54,A52),0)',
                arguments     => {
                    A1       => $_,
                    A52      => $_,
                    A53_A54 => $maxkVA,
                }
              )
              : Arithmetic(
                name          => 'Maximum demand run kVA at ' . $_->{name},
                groupName     => 'Maximum demand',
                defaultFormat => '0soft',
                arithmetic    => '=IF(ISNUMBER(A1),SQRT('
                  . 'INDEX(A53_A54,A52)^2+INDEX(A63_A64,A62)^2),0)',
                arguments => {
                    A1       => $_,
                    A52      => $_,
                    A53_A54 => $maxkVA->[0],
                    A62      => $_,
                    A63_A64 => $maxkVA->[1],
                }
              );
            $ca;
        } @locMatch;

        my $c1l = Arithmetic(
            name       => 'Average local charge 1 (£/kVA/year)',
            newBlock   => 1,
            arithmetic => '=IF('
              . join( '+', map { 'A' . ( $_ ? "5$_" : 1 ) } 0 .. $last )
              . '=0,IF(COUNT('
              . join( ',', map { "A8$_" } 0 .. $last ) . '),('
              . join( '+', map { "A4$_" } 0 .. $last )
              . ')/COUNT('
              . join( ',', map { "A9$_" } 0 .. $last )
              . '),0),('
              . join( '+', map { "A2$_*A3$_" } 0 .. $last ) . ')/('
              . join( '+', map { "A6$_" } 0 .. $last ) . '))',
            arguments => {
                map {
                    'A' . ( $_ ? "5$_" : 1 ) => $kVA1[$_],
                      "A6$_" => $kVA1[$_],
                      "A2$_" => $kVA1[$_],
                      "A3$_" => $c1l[$_],
                      "A4$_" => $c1l[$_],
                      "A8$_" => $locMatch[$_],
                      "A9$_" => $locMatch[$_],
                } 0 .. $last
            }
        );

        my $c1n = Arithmetic(
            name       => 'Average network charge 1 (£/kVA/year)',
            arithmetic => '=IF('
              . join( '+', map { 'A' . ( $_ ? "5$_" : 1 ) } 0 .. $last )
              . '=0,IF(COUNT('
              . join( ',', map { "A8$_" } 0 .. $last ) . '),('
              . join( '+', map { "A4$_" } 0 .. $last )
              . ')/COUNT('
              . join( ',', map { "A9$_" } 0 .. $last )
              . '),0),('
              . join( '+', map { "A2$_*A3$_" } 0 .. $last ) . ')/('
              . join( '+', map { "A6$_" } 0 .. $last ) . '))',
            arguments => {
                map {
                    'A' . ( $_ ? "5$_" : 1 ) => $kVA1[$_],
                      "A6$_" => $kVA1[$_],
                      "A2$_" => $kVA1[$_],
                      "A3$_" => $c1n[$_],
                      "A4$_" => $c1n[$_],
                      "A8$_" => $locMatch[$_],
                      "A9$_" => $locMatch[$_],
                } 0 .. $last
            }
        );

        my $active1 = Arithmetic(
            name => 'Total active power in maximum demand scenario (kW)',
            defaultFormat => '0soft',
            arithmetic    => '=0-' . join(
                '-',
                map {
                        'IF(ISNUMBER(A'
                      . ( $_ ? "5$_" : 1 )
                      . "),INDEX(A6${_}_A2$_,A3$_),0)"
                } 0 .. $last
            ),
            arguments => {
                map {
                    'A' . ( $_ ? "5$_" : 1 ) => $locMatch[$_],
                      "A6${_}_A2$_" => $a1d,
                      "A3$_"         => $locMatch[$_],
                } 0 .. $last
            }
        );

        my $invpf1 = Arithmetic(
            name       => 'Inverse power factor, maximum demand (kVA/kW)',
            arithmetic => '=IF(A6=0,1,MAX(1,SQRT(A9^2+('
              . join( '+',
                map { "IF(ISNUMBER(A2$_),INDEX(A4${_}_A5$_,A3$_),0)" }
                  0 .. $last )
              . ')^2)/A1))',
            arguments => {
                A1 => $active1,
                A6 => $active1,
                A9 => $active1,
                map {
                    (
                        "A2$_"         => $locMatch[$_],
                        "A3$_"         => $locMatch[$_],
                        "A4${_}_A5$_" => $r1d,
                      )
                } 0 .. $last
            }
        );

        return [ $c1l, $c1n, undef ], [ undef, $invpf1, undef ], [];

    }

    my $locMatchB = Arithmetic(
        name          => 'Parent location',
        defaultFormat => 'locsoft',
        arithmetic    => '=MATCH(INDEX(A7_A8,A1),A5_A6,0)',
        arguments     => {
            A1     => $locMatchA,
            A5_A6 => $locations,
            A7_A8 => $locParent,
        }
    ) if $tariffLoc;

    my $locMatchC = Arithmetic(
        name          => 'Grandparent location',
        defaultFormat => 'locsoft',
        arithmetic    => '=MATCH(INDEX(A7_A8,A1),A5_A6,0)',
        arguments     => {
            A1     => $locMatchB,
            A5_A6 => $locations,
            A7_A8 => $locParent,
        }
    ) if $tariffLoc;

    my $ca1 = Arithmetic(
        name       => 'Location charge 1 £/kVA/year',
        arithmetic => $model->{noNegative}
        ? '=IF(ISNUMBER(A1),MAX(0,INDEX(A53_A54,A52)),0)'
        : '=IF(ISNUMBER(A1),INDEX(A53_A54,A52),0)',
        arguments => {
            A1       => $locMatchA,
            A52      => $locMatchA,
            A53_A54 => $c1,
        }
    ) if $tariffLoc;

    my $cb1 = Arithmetic(
        name => (
            $model->{method} =~ /LRIC/i ? 'Linked location 1'
            : 'Parent location'
          )
          . ' charge 1 £/kVA/year',
        arithmetic => $model->{noNegative}
        ? '=IF(ISNUMBER(A1),MAX(0,INDEX(A53_A54,A52)),0)'
        : '=IF(ISNUMBER(A1),INDEX(A53_A54,A52),0)',
        arguments => {
            A1       => $locMatchB,
            A52      => $locMatchB,
            A53_A54 => $c1,
        }
    ) if $tariffLoc;

    my $cc1 = Arithmetic(
        name       => 'Grandparent location charge 1 £/kVA/year',
        arithmetic => $model->{noNegative}
        ? '=IF(ISNUMBER(A1),MAX(0,INDEX(A53_A54,A52)),0)'
        : '=IF(ISNUMBER(A1),INDEX(A53_A54,A52),0)',
        arguments => {
            A1       => $locMatchC,
            A52      => $locMatchC,
            A53_A54 => $c1,
        }
    );

    my ( $rfa1, $rfb1, $rfc1 );

    if ( $model->{legacy201} ) {
        $rfa1 = Arithmetic(
            name => 'Network group reactive factor, maximum demand (kVAr/kVA)',
            arithmetic => '=IF(ISNUMBER(A1),INDEX(A53_A54,A52),0)',
            arguments  => {
                A1       => $locMatchA,
                A52      => $locMatchA,
                A53_A54 => $rf1,
            }
        );
    }
    else {
        my $kVA = Arithmetic(
            name          => 'Network group maximum demand (kVA)',
            defaultFormat => '0soft',
            arithmetic    => '=IF(ISNUMBER(A1),'
              . 'SQRT((INDEX(A53_A54,A52)+INDEX(A73_A74,A72))^2+'
              . '(INDEX(A63_A64,A62)+INDEX(A83_A84,A82))^2)' . ',0)',
            arguments => {
                A1       => $locMatchA,
                A52      => $locMatchA,
                A53_A54 => $rf1->[0],
                A62      => $locMatchA,
                A63_A64 => $rf1->[1],
                A72      => $locMatchA,
                A73_A74 => $rf1->[2],
                A82      => $locMatchA,
                A83_A84 => $rf1->[3],
            },
        );
        $rfa1 = Arithmetic(
            name => 'Network group reactive factor, maximum demand (kVAr/kVA)',
            arithmetic =>
              '=IF(A1,0-(INDEX(A23_A24,A22)+INDEX(A33_A34,A32))/A4,0)',
            arguments => {
                A1       => $kVA,
                A4       => $kVA,
                A22      => $locMatchA,
                A23_A24 => $rf1->[1],
                A32      => $locMatchA,
                A33_A34 => $rf1->[3],
            }
        );
    }

    if ( $model->{legacy201} ) {
        $rfb1 = Arithmetic(
            name => 'Parent group reactive factor, maximum demand (kVAr/kVA)',
            arithmetic => '=IF(ISNUMBER(A1),INDEX(A53_A54,A52),0)',
            arguments  => {
                A1       => $locMatchB,
                A52      => $locMatchB,
                A53_A54 => $rf1,
            }
        );
    }
    else {
        my $kVA = Arithmetic(
            name          => 'Parent group maximum demand (kVA)',
            defaultFormat => '0soft',
            arithmetic    => '=IF(ISNUMBER(A1),'
              . 'SQRT((INDEX(A53_A54,A52)+INDEX(A73_A74,A72))^2+'
              . '(INDEX(A63_A64,A62)+INDEX(A83_A84,A82))^2)' . ',0)',
            arguments => {
                A1       => $locMatchB,
                A52      => $locMatchB,
                A53_A54 => $rf1->[0],
                A62      => $locMatchB,
                A63_A64 => $rf1->[1],
                A72      => $locMatchB,
                A73_A74 => $rf1->[2],
                A82      => $locMatchB,
                A83_A84 => $rf1->[3],
            },
        );
        $rfb1 = Arithmetic(
            name => 'Parent group reactive factor, maximum demand (kVAr/kVA)',
            arithmetic =>
              '=IF(A1,0-(INDEX(A23_A24,A22)+INDEX(A33_A34,A32))/A4,0)',
            arguments => {
                A1       => $kVA,
                A4       => $kVA,
                A22      => $locMatchB,
                A23_A24 => $rf1->[1],
                A32      => $locMatchB,
                A33_A34 => $rf1->[3],
            }
        );
    }

    if ( $model->{legacy201} ) {
        $rfc1 = Arithmetic(
            name =>
              'Grandparent group reactive factor, maximum demand (kVAr/kVA)',
            arithmetic => '=IF(ISNUMBER(A1),INDEX(A53_A54,A52),0)',
            arguments  => {
                A1       => $locMatchC,
                A52      => $locMatchC,
                A53_A54 => $rf1,
            }
        );
    }
    else {
        my $kVA = Arithmetic(
            name          => 'Grandparent group maximum demand (kVA)',
            defaultFormat => '0soft',
            arithmetic    => '=IF(ISNUMBER(A1),'
              . 'SQRT((INDEX(A53_A54,A52)+INDEX(A73_A74,A72))^2+'
              . '(INDEX(A63_A64,A62)+INDEX(A83_A84,A82))^2)' . ',0)',
            arguments => {
                A1       => $locMatchC,
                A52      => $locMatchC,
                A53_A54 => $rf1->[0],
                A62      => $locMatchC,
                A63_A64 => $rf1->[1],
                A72      => $locMatchC,
                A73_A74 => $rf1->[2],
                A82      => $locMatchC,
                A83_A84 => $rf1->[3],
            },
        );
        $rfc1 = Arithmetic(
            name =>
              'Grandparent group reactive factor, maximum demand (kVAr/kVA)',
            arithmetic =>
              '=IF(A1,0-(INDEX(A23_A24,A22)+INDEX(A33_A34,A32))/A4,0)',
            arguments => {
                A1       => $kVA,
                A4       => $kVA,
                A22      => $locMatchC,
                A23_A24 => $rf1->[1],
                A32      => $locMatchC,
                A33_A34 => $rf1->[3],
            }
        );
    }

    my $pfa1 = Arithmetic(
        name       => 'Network group power factor, maximum demand (kW/kVA)',
        arithmetic => '=SQRT(1-A1^2)',
        arguments  => { A1 => $rfa1 }
    );

    my $pfb1 = Arithmetic(
        name       => 'Parent group power factor, maximum demand (kW/kVA)',
        arithmetic => '=SQRT(1-A1^2)',
        arguments  => { A1 => $rfb1 }
    );

    my $pfc1 = Arithmetic(
        name       => 'Grandparent group power factor, maximum demand (kW/kVA)',
        groupName  => 'Network flows',
        arithmetic => '=SQRT(1-A1^2)',
        arguments  => { A1 => $rfc1 }
    );

    my $rft1 = Arithmetic(
        name       => 'Top level reactive factor, maximum demand (kVAr/kVA)',
        arithmetic => '=IF(ISNUMBER(A3),A6,IF(ISNUMBER(A2),A5,A4))',
        arguments  => {
            A1 => $locMatchA,
            A2 => $locMatchB,
            A3 => $locMatchC,
            A4 => $rfa1,
            A5 => $rfb1,
            A6 => $rfc1,
        }
    );

    my $pft1 = Arithmetic(
        name       => 'Top level power factor, maximum demand (kW/kVA)',
        arithmetic => '=SQRT(1-A1^2)',
        arguments  => { A1 => $rft1 }
    );

    [ $ca1, $cb1, $cc1 ], [ $pfa1, $pfb1, $pfc1 ], [ $rfa1, $rfb1, $rfc1 ];

}

1;
