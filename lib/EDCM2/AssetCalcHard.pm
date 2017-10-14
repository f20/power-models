package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2012 Energy Networks Association Limited and others.
Copyright 2013-2017 Franck Latrémolière, Reckon LLP and others.

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

sub notionalAssetCalculatorHardcoded {

    my (
        $model,                $tariffCategoryset, $useProportions,
        $useProportionsCooked, $customerCategory,  $accretion,
        $diversity,            $capUseRate,        $purpleUseRate,
    ) = @_;

    my (
        @assetsCapacity,       @assetsConsumption,
        @assetsCapacityCooked, @assetsConsumptionCooked,
    );

    my $starIV5       = $useProportions       ? '*A5' : '';
    my $starIV5Cooked = $useProportionsCooked ? '*A5' : '';

    if ( !$useTextMatching ) {    # does not support allowInvalid

        push @assetsCapacity,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[1] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic => qq@=IF(A1=1000,A4*A8/(1+A3)$starIV5,0)@,
            arguments  => {
                A1 => $tariffCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[1] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic => qq@=IF(A1=1000,A4*A8/(1+A3)$starIV5Cooked,0)@,
            newBlock   => 1,
            arguments  => {
                A1 => $tariffCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportionsCooked
                ? ( A5 => $useProportionsCooked )
                : (),
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[2] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic => qq@=IF(MOD(A1,1000)=100,A4*A8/(1+A3)$starIV5,0)@,
            arguments  => {
                A1 => $tariffCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[2] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
              qq@=IF(MOD(A1,1000)=100,A4*A8/(1+A3)$starIV5Cooked,0)@,
            arguments => {
                A1 => $tariffCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportionsCooked
                ? ( A5 => $useProportionsCooked )
                : (),
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[3] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic => qq@=IF(MOD(A1,100)=10,A4*A8/(1+A3)$starIV5,0)@,
            arguments  => {
                A1 => $tariffCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[3] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic => qq@=IF(MOD(A1,100)=10,A4*A8/(1+A3)$starIV5Cooked,0)@,
            arguments  => {
                A1 => $tariffCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportionsCooked
                ? ( A5 => $useProportionsCooked )
                : (),
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[4] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
            arithmetic =>
              qq@=IF(AND(MOD(A1,10)>0,MOD(A2,1000)>1),A4*A8/(1+A3)$starIV5,0)@,
            arguments => {
                A1 => $tariffCategory,
                A2 => $tariffCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[4] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
            arithmetic =>
qq@=IF(AND(MOD(A1,10)>0,MOD(A2,1000)>1),A4*A8/(1+A3)$starIV5Cooked,0)@,
            arguments => {
                A1 => $tariffCategory,
                A2 => $tariffCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportionsCooked
                ? ( A5 => $useProportionsCooked )
                : (),
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[5] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[5] ] ),
            arithmetic => qq@=IF(MOD(A1,1000)=1,A4*A8/(1+A3)$starIV5,0)@,
            arguments  => {
                A1 => $tariffCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[5] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[5] ] ),
            arithmetic => qq@=IF(MOD(A1,1000)=1,A4*A8/(1+A3)$starIV5Cooked,0)@,
            arguments  => {
                A1 => $tariffCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportionsCooked
                ? ( A5 => $useProportionsCooked )
                : (),
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[1] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic => qq@=IF(A1>1000,A4*A9$starIV5,0)@,
            arguments  => {
                A1 => $tariffCategory,
                A4 => $accretion,
                A9 => ref $purpleUseRate eq 'ARRAY'
                ? $purpleUseRate->[1]
                : $purpleUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[1] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic => qq@=IF(A1>1000,A4*A9$starIV5Cooked,0)@,
            arguments  => {
                A1 => $tariffCategory,
                A4 => $accretion,
                A9 => ref $purpleUseRate eq 'ARRAY' ? $purpleUseRate->[1]
                : $purpleUseRate,
                $useProportionsCooked ? ( A5 => $useProportionsCooked )
                : (),
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[2] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic => qq@=IF(MOD(A1,1000)>100,A4*A9$starIV5,0)@,
            arguments  => {
                A1 => $tariffCategory,
                A4 => $accretion,
                A9 => ref $purpleUseRate eq 'ARRAY'
                ? $purpleUseRate->[1]
                : $purpleUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[2] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic => qq@=IF(MOD(A1,1000)>100,A4*A9$starIV5Cooked,0)@,
            arguments  => {
                A1 => $tariffCategory,
                A4 => $accretion,
                A9 => ref $purpleUseRate eq 'ARRAY' ? $purpleUseRate->[1]
                : $purpleUseRate,
                $useProportionsCooked ? ( A5 => $useProportionsCooked )
                : (),
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[3] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic => qq@=IF(MOD(A1,100)>10,A4*A9$starIV5,0)@,
            arguments  => {
                A1 => $tariffCategory,
                A4 => $accretion,
                A9 => ref $purpleUseRate eq 'ARRAY'
                ? $purpleUseRate->[1]
                : $purpleUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[3] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic => qq@=IF(MOD(A1,100)>10,A4*A9$starIV5Cooked,0)@,
            arguments  => {
                A1 => $tariffCategory,
                A4 => $accretion,
                A9 => ref $purpleUseRate eq 'ARRAY' ? $purpleUseRate->[1]
                : $purpleUseRate,
                $useProportionsCooked ? ( A5 => $useProportionsCooked )
                : (),
            },
          );

    }

    elsif ( !$model->{allowInvalid} ) {    # legacy, without allowInvalid

        push @assetsCapacity,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[1] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic => qq@=IF(A1="D1000",A4*A8/(1+A3)$starIV5,0)@,
            arguments  => {
                A1 => $customerCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[1] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic => qq@=IF(A1="D1000",A4*A8/(1+A3)$starIV5Cooked,0)@,
            arguments  => {
                A1 => $customerCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportionsCooked
                ? ( A5 => $useProportionsCooked )
                : (),
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[2] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
              qq@=IF(ISNUMBER(SEARCH("D?100",A1)),A4*A8/(1+A3)$starIV5,0)@,
            arguments => {
                A1 => $customerCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[2] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?100",A1)),A4*A8/(1+A3)$starIV5Cooked,0)@,
            arguments => {
                A1 => $customerCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportionsCooked
                ? ( A5 => $useProportionsCooked )
                : (),
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[3] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
              qq@=IF(ISNUMBER(SEARCH("D??10",A1)),A4*A8/(1+A3)$starIV5,0)@,
            arguments => {
                A1 => $customerCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[3] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D??10",A1)),A4*A8/(1+A3)$starIV5Cooked,0)@,
            arguments => {
                A1 => $customerCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportionsCooked
                ? ( A5 => $useProportionsCooked )
                : (),
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[4] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
            arithmetic =>
qq@=IF(OR(A6="D0002",ISNUMBER(SEARCH("D?1?1",A1)),ISNUMBER(SEARCH("D??11",A7))),A4*A8/(1+A3)$starIV5,0)@,
            arguments => {
                A1 => $customerCategory,
                A4 => $accretion,
                A6 => $customerCategory,
                A7 => $customerCategory,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[4] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
            arithmetic =>
qq@=IF(OR(A6="D0002",ISNUMBER(SEARCH("D?1?1",A1)),ISNUMBER(SEARCH("D??11",A7))),A4*A8/(1+A3)$starIV5Cooked,0)@,
            arguments => {
                A1 => $customerCategory,
                A7 => $customerCategory,
                A4 => $accretion,
                A6 => $customerCategory,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportionsCooked
                ? ( A5 => $useProportionsCooked )
                : (),
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[5] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[5] ] ),
            arithmetic =>
              qq@=IF(ISNUMBER(SEARCH("D?001",A1)),A4*A8/(1+A3)$starIV5,0)@,
            arguments => {
                A1 => $customerCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[5] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[5] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?001",A1)),A4*A8/(1+A3)$starIV5Cooked,0)@,
            arguments => {
                A1 => $customerCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportionsCooked
                ? ( A5 => $useProportionsCooked )
                : (),
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[1] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic =>
qq@=IF(A1="D1000",0,IF(ISNUMBER(SEARCH("D1???",A2)),A4*A9$starIV5,0))@,
            arguments => {
                A1 => $customerCategory,
                A2 => $customerCategory,
                A4 => $accretion,
                A9 => ref $purpleUseRate eq 'ARRAY'
                ? $purpleUseRate->[1]
                : $purpleUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[1] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic =>
qq@=IF(A1="D1000",0,IF(ISNUMBER(SEARCH("D1???",A2)),A4*A9$starIV5Cooked,0))@,
            arguments => {
                A1 => $customerCategory,
                A2 => $customerCategory,
                A4 => $accretion,
                A9 => ref $purpleUseRate eq 'ARRAY' ? $purpleUseRate->[1]
                : $purpleUseRate,
                $useProportionsCooked ? ( A5 => $useProportionsCooked )
                : (),
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[2] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?100",A1)),0,IF(ISNUMBER(SEARCH("D?1??",A2)),A4*A9$starIV5,0))@,
            arguments => {
                A1 => $customerCategory,
                A2 => $customerCategory,
                A4 => $accretion,
                A9 => ref $purpleUseRate eq 'ARRAY'
                ? $purpleUseRate->[1]
                : $purpleUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[2] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?100",A1)),0,IF(ISNUMBER(SEARCH("D?1??",A2)),A4*A9$starIV5Cooked,0))@,
            arguments => {
                A1 => $customerCategory,
                A2 => $customerCategory,
                A4 => $accretion,
                A9 => ref $purpleUseRate eq 'ARRAY' ? $purpleUseRate->[1]
                : $purpleUseRate,
                $useProportionsCooked ? ( A5 => $useProportionsCooked )
                : (),
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[3] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D??10",A1)),0,IF(ISNUMBER(SEARCH("D??1?",A2)),A4*A9$starIV5,0))@,
            arguments => {
                A1 => $customerCategory,
                A2 => $customerCategory,
                A4 => $accretion,
                A9 => ref $purpleUseRate eq 'ARRAY'
                ? $purpleUseRate->[1]
                : $purpleUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[3] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D??10",A1)),0,IF(ISNUMBER(SEARCH("D??1?",A2)),A4*A9$starIV5Cooked,0))@,
            arguments => {
                A1 => $customerCategory,
                A2 => $customerCategory,
                A4 => $accretion,
                A9 => ref $purpleUseRate eq 'ARRAY' ? $purpleUseRate->[1]
                : $purpleUseRate,
                $useProportionsCooked ? ( A5 => $useProportionsCooked )
                : (),
            },
          );

    }

    else {    # legacy, if allowInvalid

        push @assetsCapacity,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[1] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic => qq@=IF(A1="D1000",A4*A8/(1+A3)$starIV5,0)@,
            arguments  => {
                A1 => $customerCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[1] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic => qq@=IF(A1="D1000",A4*A8/(1+A3)$starIV5Cooked,0)@,
            arguments  => {
                A1 => $customerCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                A5 => $useProportionsCooked,
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[1] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic =>
qq@=IF(OR(A1="D1000",ISNUMBER(SEARCH("G????",A20))),0,A4*A9$starIV5)@,
            arguments => {
                A1  => $customerCategory,
                A20 => $customerCategory,
                A21 => $customerCategory,
                A4  => $accretion,
                A9  => ref $purpleUseRate eq 'ARRAY'
                ? $purpleUseRate->[1]
                : $purpleUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        my $useProportionsCapped = Arithmetic(
            name       => 'Network use factors (capped only)',
            arithmetic => '=MIN(A1+0,A2+0)',
            arguments  => {
                A1 => $useProportions,
                A2 => $usePropCap,
            }
        );

        push @assetsConsumptionCooked,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[1] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic =>
qq@=IF(OR(A1="D1000",ISNUMBER(SEARCH("G????",A20))),0,IF(ISNUMBER(SEARCH("D1???",A21)),A5,A6)*A4*A9)@,
            arguments => {
                A1  => $customerCategory,
                A20 => $customerCategory,
                A21 => $customerCategory,
                A4  => $accretion,
                A9  => ref $purpleUseRate eq 'ARRAY'
                ? $purpleUseRate->[1]
                : $purpleUseRate,
                A5 => $useProportionsCooked,
                A6 => $useProportionsCapped,
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[2] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
              qq@=IF(ISNUMBER(SEARCH("D?100",A1)),A4*A8/(1+A3)$starIV5,0)@,
            arguments => {
                A1 => $customerCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[2] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?100",A1)),A4*A8/(1+A3)$starIV5Cooked,0)@,
            arguments => {
                A1 => $customerCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                A5 => $useProportionsCooked,
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[2] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",A20)),ISNUMBER(SEARCH("D?100",A1))),0,A4*A9$starIV5)@,
            arguments => {
                A1  => $customerCategory,
                A20 => $customerCategory,
                A21 => $customerCategory,
                A4  => $accretion,
                A9  => ref $purpleUseRate eq 'ARRAY'
                ? $purpleUseRate->[1]
                : $purpleUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[2] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",A20)),ISNUMBER(SEARCH("D?100",A1))),0,IF(ISNUMBER(SEARCH("D?1??",A21)),A5,A6)*A4*A9)@,
            arguments => {
                A1  => $customerCategory,
                A20 => $customerCategory,
                A21 => $customerCategory,
                A4  => $accretion,
                A9  => ref $purpleUseRate eq 'ARRAY'
                ? $purpleUseRate->[1]
                : $purpleUseRate,
                A5 => $useProportionsCooked,
                A6 => $useProportionsCapped,
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[3] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
              qq@=IF(ISNUMBER(SEARCH("D??10",A1)),A4*A8/(1+A3)$starIV5,0)@,
            arguments => {
                A1  => $customerCategory,
                A20 => $customerCategory,
                A21 => $customerCategory,
                A4  => $accretion,
                A3  => $diversity,
                A8  => $capUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[3] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D??10",A1)),A4*A8/(1+A3)$starIV5Cooked,0)@,
            arguments => {
                A1  => $customerCategory,
                A20 => $customerCategory,
                A21 => $customerCategory,
                A4  => $accretion,
                A3  => $diversity,
                A8  => $capUseRate,
                A5  => $useProportionsCooked,
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[3] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",A20)),ISNUMBER(SEARCH("D??10",A1))),0,A4*A9$starIV5)@,
            arguments => {
                A1  => $customerCategory,
                A20 => $customerCategory,
                A21 => $customerCategory,
                A4  => $accretion,
                A9  => ref $purpleUseRate eq 'ARRAY'
                ? $purpleUseRate->[1]
                : $purpleUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[3] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",A20)),ISNUMBER(SEARCH("D??10",A1))),0,IF(ISNUMBER(SEARCH("D??1?",A21)),A5,A6)*A4*A9)@,
            arguments => {
                A1  => $customerCategory,
                A20 => $customerCategory,
                A21 => $customerCategory,
                A4  => $accretion,
                A9  => ref $purpleUseRate eq 'ARRAY'
                ? $purpleUseRate->[1]
                : $purpleUseRate,
                A5 => $useProportionsCooked,
                A6 => $useProportionsCapped,
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[4] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
            arithmetic =>
qq@=IF(OR(A6="D0002",ISNUMBER(SEARCH("D?1?1",A1)),ISNUMBER(SEARCH("D??11",A7))),A4*A8/(1+A3)$starIV5,0)@,
            arguments => {
                A1 => $customerCategory,
                A4 => $accretion,
                A6 => $customerCategory,
                A7 => $customerCategory,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[4] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
            arithmetic =>
qq@=IF(OR(A6="D0002",ISNUMBER(SEARCH("D?1?1",A1)),ISNUMBER(SEARCH("D??11",A7))),A4*A8/(1+A3)$starIV5Cooked,0)@,
            arguments => {
                A1 => $customerCategory,
                A7 => $customerCategory,
                A4 => $accretion,
                A6 => $customerCategory,
                A3 => $diversity,
                A8 => $capUseRate,
                A5 => $useProportionsCooked,
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[4] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",A20)),A22="D0002",ISNUMBER(SEARCH("D?1?1",A1)),ISNUMBER(SEARCH("D??11",A7))),0,A4*A9$starIV5)@,
            arguments => {
                A1  => $customerCategory,
                A22 => $customerCategory,
                A7  => $customerCategory,
                A20 => $customerCategory,
                A21 => $customerCategory,
                A4  => $accretion,
                A9  => ref $purpleUseRate eq 'ARRAY'
                ? $purpleUseRate->[1]
                : $purpleUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[4] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",A20)),A22="D0002",ISNUMBER(SEARCH("D?1?1",A1)),ISNUMBER(SEARCH("D??11",A7))),0,A6*A4*A9)@,
            arguments => {
                A1  => $customerCategory,
                A22 => $customerCategory,
                A7  => $customerCategory,
                A20 => $customerCategory,
                A21 => $customerCategory,
                A4  => $accretion,
                A9  => ref $purpleUseRate eq 'ARRAY'
                ? $purpleUseRate->[1]
                : $purpleUseRate,
                A5 => $useProportionsCooked,
                A6 => $useProportionsCapped,
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[5] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[5] ] ),
            arithmetic =>
              qq@=IF(ISNUMBER(SEARCH("D?001",A1)),A4*A8/(1+A3)$starIV5,0)@,
            arguments => {
                A1 => $customerCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[5] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[5] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?001",A1)),A4*A8/(1+A3)$starIV5Cooked,0)@,
            arguments => {
                A1 => $customerCategory,
                A4 => $accretion,
                A3 => $diversity,
                A8 => $capUseRate,
                A5 => $useProportionsCooked,
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[5] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[5] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",A20)),ISNUMBER(SEARCH("D?001",A1))),0,A4*A9$starIV5)@,
            arguments => {
                A1  => $customerCategory,
                A20 => $customerCategory,
                A21 => $customerCategory,
                A4  => $accretion,
                A9  => ref $purpleUseRate eq 'ARRAY'
                ? $purpleUseRate->[1]
                : $purpleUseRate,
                $useProportions ? ( A5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[5] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[5] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",A20)),ISNUMBER(SEARCH("D?001",A1))),0,A6*A4*A9)@,
            arguments => {
                A1  => $customerCategory,
                A20 => $customerCategory,
                A21 => $customerCategory,
                A4  => $accretion,
                A9  => ref $purpleUseRate eq 'ARRAY'
                ? $purpleUseRate->[1]
                : $purpleUseRate,
                A5 => $useProportionsCooked,
                A6 => $useProportionsCapped,
            },
          );

    }

    Arithmetic(
        name       => 'Total notional capacity assets (£/kVA)',
        groupName  => 'First set of notional capacity assets',
        cols       => 0,
        arithmetic => '=' . join( '+', map { "A$_" } 1 .. @assetsCapacity ),
        arguments  => {
            map { ( "A$_" => $assetsCapacity[ $_ - 1 ] ) } 1 .. @assetsCapacity
        },
      ),
      Arithmetic(
        name       => 'Total notional consumption assets (£/kVA)',
        groupName  => 'First set of notional consumption assets',
        cols       => 0,
        arithmetic => '=' . join( '+', map { "A$_" } 1 .. @assetsConsumption ),
        arguments  => {
            map { ( "A$_" => $assetsConsumption[ $_ - 1 ] ) }
              1 .. @assetsConsumption
        },
      ),
      Arithmetic(
        name       => 'Second set of capacity assets (£/kVA)',
        groupName  => 'Second set of notional capacity assets',
        cols       => 0,
        arithmetic => '='
          . join( '+', map { "A$_" } 1 .. @assetsCapacityCooked ),
        arguments => {
            map { ( "A$_" => $assetsCapacityCooked[ $_ - 1 ] ) }
              1 .. @assetsCapacityCooked
        },
      ),
      Arithmetic(
        name       => 'Second set of consumption assets (£/kVA)',
        groupName  => 'Second set of notional consumption assets',
        cols       => 0,
        arithmetic => '='
          . join( '+', map { "A$_" } 1 .. @assetsConsumptionCooked ),
        arguments => {
            map { ( "A$_" => $assetsConsumptionCooked[ $_ - 1 ] ) }
              1 .. @assetsConsumptionCooked
        },
      );

}

1;
