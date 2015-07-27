package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2012 Energy Networks Association Limited and others.
Copyright 2013-2015 Franck Latrémolière, Reckon LLP and others.

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
        $model,                       $acCoef,
        $activeCoincidence,           $charges1,
        $daysInYear,                  $reactiveCoincidence,
        $reCoef,                      $sFactor,
        $purpleHours,                 $purpleHoursGen,
        $demandCapacity,              $chargeableGenerationCapacity,
        $creditableCapacity,          $rateExit,
        $activeCoincidenceUndoctored, $reactiveCoincidenceUndoctored,
        $reactiveCoincidence935
    ) = @_;

    $model->{demandConsumptionFcpLric} = my $demandConsumptionFcpLric =
      (     !$model->{removeDemandCharge1}
          || $model->{removeDemandCharge1} =~ /keepunitrate/i )
      && ( grep { $charges1->[$_] } 1 .. $#$charges1 )
      ? Arithmetic(
        name          => 'Import demand charge p/kVA/day',
        newBlock      => 1,
        defaultFormat => '0.00softnz',
        rows          => $demandCapacity->{rows},
        arithmetic    => '=100*(' . join(
            '+',
            map {
                "A9$_*"
                  . (
                    $reCoef->[$_]
                    ? "MAX(0,A5$_*A7$_+A6$_*A8$_)"
                    : "A5$_*A7$_"
                  )
            } grep { $charges1->[$_] } 1 .. $#$charges1
          )
          . ')/A2',
        arguments => {
            A2 => $daysInYear,
            map {
                (
                    "A5$_" => $activeCoincidenceUndoctored,
                    "A9$_" => $charges1->[$_],
                    "A7$_" => $acCoef->[$_],
                    $reCoef->[$_]
                    ? (
                        "A6$_" => $reactiveCoincidenceUndoctored,
                        "A8$_" => $reCoef->[$_]
                      )
                    : ()
                  )
            } grep { $charges1->[$_] } 1 .. $#$charges1
        }
      )
      : Constant(
        isZero => 1,
        rows   => $demandCapacity->{rows},
        name   => 'Import demand charge before matching p/kVA/day',
        data   => [ map { 0 } @{ $demandCapacity->{rows}{list} } ],
      );

    $model->{demandCapacityFcpLric} = my $demandCapacityFcpLric =
      $charges1->[0]
      && !$model->{removeDemandCharge1} ? Arithmetic(
        name          => 'Import capacity charge p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=100*A1/A2',
        arguments     => {
            A1 => $charges1->[0],
            A2 => $daysInYear,
        }
      )
      : $charges1->[0]
      && $model->{removeDemandCharge1}
      && $model->{removeDemandCharge1} =~ /keepunitrate/i ? Arithmetic(
        name       => 'Deduct unit rate charge 1 from capacity p/kVA/day',
        arithmetic => '=0-A1',
        arguments  => { A1 => $demandConsumptionFcpLric, },
      )
      : Constant(
        isZero => 1,
        rows   => $demandCapacity->{rows},
        name   => 'Import capacity charge p/kVA/day',
        data   => [ map { 0 } @{ $demandCapacity->{rows}{list} } ],
      );

    push @{ $model->{matricesData}[1] },
      Stack( sources => [$demandCapacityFcpLric] )
      if $model->{matricesData};

    # The following is pretty weird; substantially revised on 25 March 2011
    # Weirdness further increased on 1 June 2012.
    my $unitRateFcpLric =
      (     !$model->{removeDemandCharge1}
          || $model->{removeDemandCharge1} =~ /keepunitrate/i )
      && ( grep { $charges1->[$_] } 1 .. $#$charges1 )
      ? Arithmetic(
        name       => "$model->{TimebandName} rate p/kWh",
        rows       => $demandCapacity->{rows},
        arithmetic => '=100*(' . join(
            '+',
            map {
                "A9$_*"
                  . (
                    $reCoef->[$_]
                    ? "IF(A2$_=0,A7$_,MAX(0,A4$_+A6$_*A8$_/A5$_))"
                    : "A7$_"
                  )
            } grep { $charges1->[$_] } 1 .. $#$charges1
          )
          . ')/A2',
        arguments => {
            A2 => $purpleHours,
            map {
                (
                    "A9$_" => $charges1->[$_],
                    "A7$_" => $acCoef->[$_],
                    $reCoef->[$_]
                    ? (
                        "A4$_" => $acCoef->[$_],
                        "A5$_" => $activeCoincidenceUndoctored,
                        "A2$_" => $activeCoincidenceUndoctored,
                        "A6$_" => $reactiveCoincidenceUndoctored,
                        "A8$_" => $reCoef->[$_]
                      )
                    : ()
                  )
            } grep { $charges1->[$_] } 1 .. $#$charges1
        }
      )
      : Constant(
        isZero => 1,
        rows   => $demandCapacity->{rows},
        name   => "$model->{TimebandName} rate p/kWh",
        data   => [ map { 0 } @{ $demandCapacity->{rows}{list} } ],
      );

    $demandCapacityFcpLric = Arithmetic(
        name          => 'FCP capacity charge p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(A3=0,A1+A2,A11)',
        arguments     => {
            A1  => $demandCapacityFcpLric,
            A11 => $demandCapacityFcpLric,
            A3  => $activeCoincidenceUndoctored,
            A2  => $demandConsumptionFcpLric,
        }
    ) if $model->{method} =~ /FCP/i;

    my $genCredit =
      @$charges1
      ? Arithmetic(
        name       => 'Generation credit (before exempt adjustment) p/kWh',
        arithmetic => '=-100*('
          . join( '+',
            $charges1->[0] ? "A90*A1" : (),
            map { $charges1->[$_] ? "A9$_" : () } 1 .. $#$charges1 )
          . ')/A2',
        arguments => {
            A2 => $purpleHoursGen,
            A1 => $sFactor,
            map { $charges1->[$_] ? ( "A9$_" => $charges1->[$_] ) : () }
              0 .. $#$charges1
        }
      )
      : Constant(
        isZero => 1,
        name   => 'Generation credit (before exempt adjustment) p/kWh',
        rows   => $sFactor->{rows},
        data   => [ map { 0 } @{ $sFactor->{rows}{list} } ],
      );

    if ( @$charges1 && $model->{lowerIntermittentCredit} ) {
        $genCredit = Arithmetic(
            name => 'Generation credit (before exempt adjustment) p/kWh',
            defaultFormat => '0.00softnz',
            arithmetic    => '=-100*A1*('
              . join( '+',
                $charges1->[0] ? "A90" : (),
                map { $charges1->[$_] ? "A9$_" : () } 1 .. $#$charges1 )
              . ')/A2',
            arguments => {
                A2 => $purpleHoursGen,
                A1 => $sFactor,
                map { $charges1->[$_] ? ( "A9$_" => $charges1->[$_] ) : () }
                  0 .. $#$charges1
            }
        );
    }

    my $genCreditCapacity = Arithmetic(
        name          => 'Generation credit (unrounded) p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(A1,-100*A91/A2*A6/A52,0)',
        arguments     => {
            A2  => $daysInYear,
            A1  => $chargeableGenerationCapacity,
            A52 => $chargeableGenerationCapacity,
            A6  => $creditableCapacity,
            A91 => $rateExit,
        }
    );

    $demandCapacityFcpLric, $genCredit, $unitRateFcpLric,
      $genCreditCapacity, $demandConsumptionFcpLric;

}

1;
