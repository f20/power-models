package EDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others.

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

sub chargesFcpLric {

    my (
        $model,               $acCoef,             $activeCoincidence,
        $charges1,            $charges2,           $daysInYear,
        $included,            $llfcs,              $reactiveCoincidence,
        $reCoef,              $sFactor,            $tariffDorG,
        $tariffs,             $redHours,           $redHoursGen,
        $totalAgreedCapacity, $creditableCapacity, $redUseRate,
      )
      = @_;

    my $genCredit = Arithmetic(
        name          => 'Generation credit p/kWh',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(IV88,IF(IV1="Generation",-100*IV4*('
          . join( '+',
            map { $charges1->[$_] ? "IU9$_" : () } 0 .. ( $#$charges1 - 1 ) )
          . ')/IV2,0),"")',
        arguments => {
            IV1  => $tariffDorG,
            IV2  => $redHoursGen,
            IV4  => $sFactor,
            IV88 => $included,
            map { $charges1->[$_] ? ( "IU9$_" => $charges1->[$_] ) : () }
              0 .. $#$charges1
        }
    );

    my $genCreditCapacity = Arithmetic(
        name          => 'Generation credit p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic    =>
'=IF(IV88,IF(AND(IV51>0,IV1="Generation"),-100*IV4*IU91/IV2*IV6/IV52,0),"")',
        arguments => {
            IV1    => $tariffDorG,
            IV2    => $daysInYear,
            IV4    => $sFactor,
            IV88   => $included,
            IV51   => $totalAgreedCapacity,
            IV52   => $totalAgreedCapacity,
            IV6    => $creditableCapacity,
            "IU91" => $charges1->[$#$charges1],
        }
    );

    $model->{demandCapacityFcpLric} = my $demandCapacityFcpLric =
      $charges1->[0]
      ? Arithmetic(
        name          => 'Import capacity charge p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(IV8,IF(IV1="Demand",100*IV91/IV2,0),"")',
        arguments     => {
            IV91 => $charges1->[0],
            IV1  => $tariffDorG,
            IV2  => $daysInYear,
            IV8  => $included
        }
      )
      : Arithmetic(
        defaultFormat => '0.00softnz',
        rows          => $tariffs->{rows},
        name          => 'Import capacity charge p/kVA/day',
        arithmetic    => 'IF(IV1,0,"")',
        arguments     => { IV1 => $included }
      );

    $model->{demandConsumptionFcpLric} = my $demandConsumptionFcpLric =
      ( grep { $charges1->[$_] } 1 .. $#$charges1 - 1 )
      ? Arithmetic(
        name          => 'Import demand charge p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(IV88,IF(IV1="Demand",100*(' . join(
            '+',
            map {
                "IU9$_*"
                  . (
                    $reCoef->[$_]
                    ? "MAX(0,IU5$_*IU7$_+IU6$_*IU8$_)"
                    : "IU5$_*IU7$_"
                  )
              } grep { $charges1->[$_] } 1 .. $#$charges1 - 1
          )
          . ')/IV2,0),"")',
        arguments => {
            IV1  => $tariffDorG,
            IV2  => $daysInYear,
            IV88 => $included,
            map {
                (
                    "IU5$_" => 0 ? $redUseRate : $activeCoincidence,
                    "IU9$_" => $charges1->[$_],
                    "IU7$_" => $acCoef->[$_],
                    $reCoef->[$_]
                    ? (
                        "IU6$_" => $reactiveCoincidence,
                        "IU8$_" => $reCoef->[$_]
                      )
                    : ()
                  )
              } grep { $charges1->[$_] } 1 .. $#$charges1 - 1
        }
      )
      : Arithmetic(
        defaultFormat => '0.000softnz',
        rows          => $tariffs->{rows},
        name          => 'Import demand charge before matching p/kVA/day',
        arithmetic    => 'IF(IV1,0,"")',
        arguments     => { IV1 => $included }
      );

    my $unitRateFcpLric = ( grep { $charges1->[$_] } 1 .. $#$charges1 - 1 )
      ? Arithmetic
      (    # this is pretty weird; substantially revised on 25 March 2011
        name          => 'Super red rate p/kWh',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(IV88,IF(IV1="Demand",100*(' . join(
            '+',
            map {
                "IV9$_*"
                  . (
                    $reCoef->[$_]
                    ? "IF(IV2$_=0,IV7$_,MAX(0,IV4$_+IV6$_*IV8$_/IV5$_))"
                    : "IV7$_"
                  )
              } grep { $charges1->[$_] } 1 .. $#$charges1 - 1
          )
          . ')/IV2,0),"")',
        arguments => {
            IV1  => $tariffDorG,
            IV2  => $redHours,
            IV88 => $included,
            map {
                (
                    "IV9$_" => $charges1->[$_],
                    "IV7$_" => $acCoef->[$_],
                    $reCoef->[$_]
                    ? (
                        "IV4$_" => $acCoef->[$_],
                        "IV5$_" => $activeCoincidence,
                        "IV2$_" => $activeCoincidence,
                        "IV6$_" => $reactiveCoincidence,
                        "IV8$_" => $reCoef->[$_]
                      )
                    : ()
                  )
              } grep { $charges1->[$_] } 1 .. $#$charges1 - 1
        }
      )
      : Arithmetic(
        defaultFormat => '0.000softnz',
        rows          => $tariffs->{rows},
        name          => 'Super red rate p/kWh',
        arithmetic    => 'IF(IV1,0,"")',
        arguments     => { IV1 => $included }
      );

    my $capacityChargeBeforeScaling = Arithmetic(
        name          => 'FCP/LRIC capacity charge p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(AND(IV1,IV2="Demand"),IV4+IV5,0)',
        arguments     => {
            IV1 => $included,
            IV2 => $tariffDorG,
            IV4 => $demandCapacityFcpLric,
            IV5 => $demandConsumptionFcpLric
        }
    );

    my ($exportCapacity) =
      grep { $_ } @$charges2
      ? Arithmetic(
        name          => 'Export capacity charge before scaling p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(IV88,IF(IV1="Generation",100*('
          . join( '+',
            ( map { $charges2->[$_] ? "IU9$_" : () } 0 .. $#$charges2 ) )
          . ')/IV2,0),"")',
        arguments => {
            IV1  => $tariffDorG,
            IV2  => $daysInYear,
            IV88 => $included,
            map { $charges2->[$_] ? ( "IU9$_" => $charges2->[$_] ) : () }
              0 .. $#$charges2
        }
      )
      : Arithmetic(
        defaultFormat => '0.000softnz',
        rows          => $tariffs->{rows},
        name          => 'Export capacity charge before matching p/kVA/day',
        arithmetic    => 'IF(IV1,0,"")',
        arguments     => { IV1 => $included }
      );

    $capacityChargeBeforeScaling, $genCredit, $exportCapacity, $unitRateFcpLric,
      $genCreditCapacity;

}

1;
