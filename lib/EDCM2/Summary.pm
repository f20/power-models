package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2013 Energy Networks Association Limited and others.

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

sub revenue {

    my (
        $model,                  $daysInYear,
        $tariffs,                $agreedCapacity,
        $generationCapacity,     $activeUnits,
        $fixedChargeDemand,      $fixedChargeGeneration,
        $importCapacity,         $exportCapacity,
        $importCapacityExceeded, $exportCapacityExceeded,
        $exportCredit,           $genCreditCapacity,
        $importCapacityScaled,   $unitRateFcpLric,
        $activeCoincidence,      $redHours,
    ) = @_;

    my @revenueBits = (

        Arithmetic(
            name          => 'Fixed charges for demand (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => '=0.01*IV9*IV1',
            arguments     => {
                IV1 => $fixedChargeDemand,
                IV9 => $daysInYear,
            }
        ),

        Arithmetic(
            name          => 'Fixed charges for generation (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => '=0.01*IV9*IV1',
            arguments     => {
                IV1 => $fixedChargeGeneration,
                IV9 => $daysInYear,
            }
        ),

        ( grep { $_ } $model->{Thursday31}[0], $model->{Thursday31}[2], ),

        Arithmetic(
            name          => 'Other import capacity charges (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => '=0.01*IV9*IV1*IV23',
            arguments     => {
                IV1  => $importCapacityScaled,
                IV7  => $importCapacityExceeded,
                IV23 => $agreedCapacity,
                IV9  => $daysInYear,
            }
        ),

        Arithmetic(
            name          => 'Import unit charges (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => '=0.01*IV9*IV42*IV1*IV7',
            arguments     => {
                IV9  => $unitRateFcpLric,
                IV42 => $agreedCapacity,
                IV1  => $activeCoincidence,
                IV7  => $redHours,
            }
        ),

        Arithmetic(
            name          => 'Generation credits (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => '=0.01*(IV9*IV42+IV1*IV6*IV7)',
            arguments     => {
                IV9  => $exportCredit,
                IV42 => $activeUnits,
                IV1  => $generationCapacity,
                IV6  => $genCreditCapacity,
                IV7  => $daysInYear,
            }
        ),

    );

    my $revenue = Arithmetic(
        name          => 'Total (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=' . join( '+', map { "IV$_" } 1 .. @revenueBits ),
        arguments =>
          { map { ( "IV$_" => $revenueBits[ $_ - 1 ] ) } 1 .. @revenueBits },
    );

    if (undef) {
        push @{ $model->{summaryTables} },
          Columnset(
            name    => 'Miscellaneous (1)',
            columns => [
                Stack( sources => [$tariffs] ),
                grep { $_ } @revenueBits,
                $revenue,
            ]
          );
    }
    $revenue;
}

sub summary {

    my ( $model, $tariffs, $revenue, $previousIncome, $agreedCapacity,
        $activeCoincidence, $charges1, $charges2, )
      = @_;

    my $r1 = Arithmetic(
        name          => $previousIncome->{name},
        arithmetic    => '=IV1',
        arguments     => { IV1 => $previousIncome },
        defaultFormat => '0copynz',
    );

    my $r2 = Stack( sources => [$revenue] );

    my $change1 = Arithmetic(
        name          => 'Change (£/year)',
        arithmetic    => '=IV1-IV4',
        defaultFormat => '0softpm',
        arguments     => { IV1 => $r2, IV4 => $r1 }
    );

    my $change2 = Arithmetic(
        name          => 'Change (%)',
        arithmetic    => '=IF(IV1<>0,IV3/IV4-1,"")',
        defaultFormat => '%softpm',
        arguments     => { IV1 => $r1, IV3 => $r2, IV4 => $r1 }
    );

    my $tc1 =
      $charges1
      && ( grep { $charges1->[$_] } 0 .. ( @$charges1 ? @$charges1 - 2 : 0 ) )
      ? Arithmetic(
        rows       => $tariffs->{rows},
        name       => 'Total charge 1 £/kVA/year',
        arithmetic => '='
          . (
            join '+',
            map { "IV2$_" }
              grep { $charges1->[$_] } 0 .. ( @$charges1 ? @$charges1 - 2 : 0 )
          ),
        arguments => {
            map { ( "IV2$_" => $charges1->[$_] ) }
            grep { $charges1->[$_] } 0 .. ( @$charges1 ? @$charges1 - 2 : 0 )
        }
      )
      : Constant(
        name => 'Total charge 1 £/kVA/year',
        rows => $tariffs->{rows},
        data => [ map { 0 } @{ $tariffs->{rows}{list} } ]
      );

    if (undef) {
        push @{ $model->{summaryTables} },
          Columnset(
            name => 'Miscellaneous (2)',
            columns =>
              [ Stack( sources => [$tariffs] ), $r1, $r2, $change1, $change2, ]
          );
    }
}

1;
