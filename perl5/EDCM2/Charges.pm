package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2012 Energy Networks Association Limited and others. All rights reserved.

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
        $model,                         $acCoef,
        $activeCoincidence,             $charges1,
        $daysInYear,                    $reactiveCoincidence,
        $reCoef,                        $sFactor,
        $redHours,                      $redHoursGen,
        $demandCapacity,                $chargeableGenerationCapacity,
        $creditableCapacity,            $redUseRate,
        $rateExit,                      $activeCoincidenceUndoctored,
        $reactiveCoincidenceUndoctored, $reactiveCoincidence935
    ) = @_;

    my $genCredit = Arithmetic(
        name          => 'Generation credit (before exempt adjustment) p/kWh',
        defaultFormat => '0.00softnz',
        arithmetic    => '=-100*('
          . join( '+',
            $charges1->[0] ? "IU90*IV1" : (),
            map { $charges1->[$_] ? "IU9$_" : () } 1 .. $#$charges1 )
          . ')/IV2',
        arguments => {
            IV2 => $redHoursGen,
            IV1 => $sFactor,
            map { $charges1->[$_] ? ( "IU9$_" => $charges1->[$_] ) : () }
              0 .. $#$charges1
        }
    );
	
	if ( $model->{DGCONDITION} ) {
	$genCredit = Arithmetic(
        name          => 'Generation credit (before exempt adjustment) p/kWh',
        defaultFormat => '0.00softnz',
        arithmetic    => '=-100*IV1*('
          . join( '+',
            $charges1->[0] ? "IU90" : (),
            map { $charges1->[$_] ? "IU9$_" : () } 1 .. $#$charges1 )
          . ')/IV2',
        arguments => {
            IV2 => $redHoursGen,
            IV1 => $sFactor,
            map { $charges1->[$_] ? ( "IU9$_" => $charges1->[$_] ) : () }
              0 .. $#$charges1
        }
    );
	}
	
    my $genCreditCapacity = Arithmetic(
        name          => 'Generation credit (unrounded) p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(IV1,-100*IU91/IV2*IV6/IV52,0)',
        arguments     => {
            IV2  => $daysInYear,
            IV1  => $chargeableGenerationCapacity,
            IV52 => $chargeableGenerationCapacity,
            IV6  => $creditableCapacity,
            IU91 => $rateExit,
        }
    );

    $model->{demandCapacityFcpLric} = my $demandCapacityFcpLric =
      $charges1->[0]
      ? Arithmetic(
        name          => 'Import capacity charge p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=100*IV1/IV2',
        arguments     => {
            IV1 => $charges1->[0],
            IV2 => $daysInYear,
        }
      )
      : Arithmetic(
        defaultFormat => '0.00softnz',
        rows          => $demandCapacity->{rows},
        name          => 'Import capacity charge p/kVA/day',
        arithmetic    => '=0',
      );

    $model->{demandConsumptionFcpLric} = my $demandConsumptionFcpLric =
      ( grep { $charges1->[$_] } 1 .. $#$charges1 )
      ? Arithmetic(
        name          => 'Import demand charge p/kVA/day',
        defaultFormat => '0.00softnz',
        rows          => $demandCapacity->{rows},
        arithmetic    => '=100*(' . join(
            '+',
            map {
                "IU9$_*"
                  . (
                    $reCoef->[$_]
                    ? "MAX(0,IU5$_*IU7$_+IU6$_*IU8$_)"
                    : "IU5$_*IU7$_"
                  )
            } grep { $charges1->[$_] } 1 .. $#$charges1
          )
          . ')/IV2',
        arguments => {
            IV2 => $daysInYear,
            map {
                (
                    "IU5$_" => $activeCoincidenceUndoctored,
                    "IU9$_" => $charges1->[$_],
                    "IU7$_" => $acCoef->[$_],
                    $reCoef->[$_]
                    ? (
                        "IU6$_" => $reactiveCoincidenceUndoctored,
                        "IU8$_" => $reCoef->[$_]
                      )
                    : ()
                  )
            } grep { $charges1->[$_] } 1 .. $#$charges1
        }
      )
      : Arithmetic(
        defaultFormat => '0.000softnz',
        rows          => $demandCapacity->{rows},
        name          => 'Import demand charge before matching p/kVA/day',
        arithmetic    => '=0',
      );

    # The following is pretty weird; substantially revised on 25 March 2011
    # Weirdness further increased on 1 June 2012.
    my $unitRateFcpLric = ( grep { $charges1->[$_] } 1 .. $#$charges1 )
      ? Arithmetic(
        name          => 'Super red rate p/kWh',
        rows          => $demandCapacity->{rows},
        defaultFormat => '0.00softnz',
        arithmetic    => '=100*(' . join(
            '+',
            map {
                "IV9$_*"
                  . (
                    $reCoef->[$_]
                    ? "IF(IV2$_=0,IV7$_,MAX(0,IV4$_+IV6$_*IV8$_/IV5$_))"
                    : "IV7$_"
                  )
            } grep { $charges1->[$_] } 1 .. $#$charges1
          )
          . ')/IV2',
        arguments => {
            IV2 => $redHours,
            map {
                (
                    "IV9$_" => $charges1->[$_],
                    "IV7$_" => $acCoef->[$_],
                    $reCoef->[$_]
                    ? (
                        "IV4$_" => $acCoef->[$_],
                        "IV5$_" => $activeCoincidenceUndoctored,
                        "IV2$_" => $activeCoincidenceUndoctored,
                        "IV6$_" => $reactiveCoincidenceUndoctored,
                        "IV8$_" => $reCoef->[$_]
                      )
                    : ()
                  )
            } grep { $charges1->[$_] } 1 .. $#$charges1
        }
      )
      : Arithmetic(
        defaultFormat => '0.000softnz',
        rows          => $demandCapacity->{rows},
        name          => 'Super red rate p/kWh',
        arithmetic    => '=0',
      );

    my $capacityChargeBeforeScaling = Arithmetic(
        name          => 'FCP/LRIC capacity charge p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(IV3=0,IV1+IV2,IV11)',
        arguments     => {
            IV1  => $demandCapacityFcpLric,
            IV11 => $demandCapacityFcpLric,
            IV3  => $activeCoincidenceUndoctored,
            IV2  => $demandConsumptionFcpLric
        }
    );

    $capacityChargeBeforeScaling, $genCredit, $unitRateFcpLric,
      $genCreditCapacity, $demandConsumptionFcpLric;

}

1;
