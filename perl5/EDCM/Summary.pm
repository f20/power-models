package EDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others. All rights reserved.

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
        $llfcs,                  $tariffs,
        $tariffDorG,             $included,
        $agreedCapacity,         $exceededCapacity,
        $activeUnits,            $fixedCharge,
        $importCapacity,         $exportCapacity,
        $importCapacityExceeded, $exportCapacityExceeded,
        $exportCredit,           $genCreditCapacity,
        $importCapacityScaled,   $unitRateFcpLric,
        $activeCoincidence,      $redHours,
      )
      = @_;

    my @revenueBits = (

        Arithmetic(
            name          => 'Fixed charges (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => '=IF(IV4,0.01*IV9*IV1,0)',
            arguments     => {
                IV1 => $fixedCharge,
                IV9 => $daysInYear,
                IV4 => $included
            }
        ),

        $model->{Thursday31}[0],
        $model->{Thursday31}[2],

        Arithmetic(
            name          => 'Other import capacity charges (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => $exceededCapacity
            ? '=IF(AND(IV4,IV3="Demand"),0.01*IV9*(IV1*IV23+IV7*IV24)-IV8,0)'
            : '=IF(IV4,0.01*IV9*IV1*IV23,0)',
            arguments => {
                IV1  => $importCapacityScaled,
                IV7  => $importCapacityExceeded,
                IV23 => $agreedCapacity,
                $exceededCapacity ? ( IV24 => $exceededCapacity ) : (),
                IV9 => $daysInYear,
                IV4 => $included,
                IV3 => $tariffDorG,
                IV8 => $model->{Thursday31}[0],
            }
        ),

        Arithmetic(
            name          => 'Import unit charges (£/year)',
            defaultFormat => '0softnz',
            arithmetic => '=IF(AND(IV1,IV2="Demand"),0.01*IV9*IV42*IV6*IV7,0)',
            arguments  => {
                IV9  => $unitRateFcpLric,
                IV42 => $agreedCapacity,
                IV6  => $activeCoincidence,
                IV7  => $redHours,
                IV1  => $included,
                IV2  => $tariffDorG,
            }
        ),

        Arithmetic(
            name          => 'Other export capacity charges (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => $exceededCapacity
            ? '=IF(AND(IV4,IV3="Generation"),0.01*IV9*(IV1*IV23+IV7*IV24)-IV8,0)'
            : '=IF(IV4,0.01*IV9*IV1*IV23-IV8,0)',
            arguments => {
                IV1  => $exportCapacity,
                IV7  => $exportCapacityExceeded,
                IV23 => $agreedCapacity,
                $exceededCapacity ? ( IV24 => $exceededCapacity ) : (),
                IV9 => $daysInYear,
                IV4 => $included,
                IV3 => $tariffDorG,
                IV8 => $model->{Thursday31}[2],
            }
        ),

        Arithmetic(
            name          => 'Generation credits (£/year)',
            defaultFormat => '0softnz',
            arithmetic    =>
              '=IF(AND(IV1,IV2="Generation"),0.01*(IV9*IV42+IV5*IV6*IV7),0)',
            arguments => {
                IV9  => $exportCredit,
                IV42 => $activeUnits,
                IV1  => $included,
                IV2  => $tariffDorG,
                IV5  => $agreedCapacity,
                IV6  => $genCreditCapacity,
                IV7  => $daysInYear,
            }
        ),

    );

    my $revenue = Arithmetic(
        name          => 'Total (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=' . join( '+', map { "IV$_" } 1 .. @revenueBits ),
        arguments     =>
          { map { ( "IV$_" => $revenueBits[ $_ - 1 ] ) } 1 .. @revenueBits },
    );

    push @{ $model->{summaryTables} },
      Columnset(
        name    => 'Miscellaneous (1)',
        columns => [
            Arithmetic(
                name          => $llfcs->{name},
                arithmetic    => '=IF(IV1,IV2,"")',
                arguments     => { IV1 => $included, IV2 => $llfcs },
                defaultFormat => 'textcopy',
            ),
            Arithmetic(
                name          => $tariffs->{name},
                arithmetic    => '=IF(IV1,IV2,"")',
                arguments     => { IV1 => $included, IV2 => $tariffs },
                defaultFormat => 'textcopy',
            ),
            @revenueBits,
            $revenue,
        ]
      );

    $revenue;

}

sub summary {

    my (
        $model,          $llfcs,          $tariffs,
        $tariffDorG,     $included,       $revenue,
        $previousIncome, $agreedCapacity, $activeCoincidence,
        $charges1,       $charges2,
      )
      = @_;

    my $r1 = Arithmetic(
        name          => $previousIncome->{name},
        arithmetic    => '=IF(IV1,IV2,"")',
        arguments     => { IV1 => $included, IV2 => $previousIncome },
        defaultFormat => '0copynz',
    );

    my $r2 = Stack( sources => [$revenue] );

    my $change1 = Arithmetic(
        name          => 'Change (£/year)',
        arithmetic    => '=IF(IV1,IV3-IV4,"")',
        defaultFormat => '0softpm',
        arguments     => { IV1 => $included, IV3 => $r2, IV4 => $r1 }
    );

    my $change2 = Arithmetic(
        name          => 'Change (%)',
        arithmetic    => '=IF(AND(IV1,IV2<>0),IV3/IV4-1,"")',
        defaultFormat => '%softpm',
        arguments => { IV1 => $included, IV2 => $r1, IV3 => $r2, IV4 => $r1 }
    );

    my $tc1 =
      ( grep { $charges1->[$_] } 0 .. ( @$charges1 ? @$charges1 - 2 : 0 ) )
      ? Arithmetic(
        rows       => $tariffDorG->{rows},
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
        rows => $tariffDorG->{rows},
        data => [ map { 0 } @{ $tariffDorG->{rows}{list} } ]
      );

    my $group = Arithmetic(
        name       => 'Customer group for summary analysis',
        arithmetic =>
'=IF(IV1,IV2&" #"&IF(IV3>15,"15+",IF(IV4<5,"0-5","5-15"))&" "&IF(IV5>15000,">15MVA","<=15MVA")&" "&IF(IV6>.1,">10%","<=10%"),"")',
        arguments => {
            IV1 => $included,
            IV2 => $tariffDorG,
            IV3 => $tc1,
            IV4 => $tc1,
            IV5 => $agreedCapacity,
            IV6 => $activeCoincidence
        }
    );

    push @{ $model->{summaryTables} },
      Columnset(
        name    => 'Miscellaneous (2)',
        columns => [
            Arithmetic(
                name          => $llfcs->{name},
                arithmetic    => '=IF(IV1,IV2,"")',
                arguments     => { IV1 => $included, IV2 => $llfcs },
                defaultFormat => 'textcopy',
            ),
            Arithmetic(
                name          => $tariffs->{name},
                arithmetic    => '=IF(IV1,IV2,"")',
                arguments     => { IV1 => $included, IV2 => $tariffs },
                defaultFormat => 'textcopy',
            ),
            $r1, $r2, $change1, $change2,
        ]
      );

    return unless $model->{summaries} =~ /vol/;

    push @{ $model->{volatilityTables} },
      Columnset(
        name    => 'Data for volatility summary',
        columns => [
            Arithmetic(
                name          => $llfcs->{name},
                arithmetic    => '=IF(IV1,IV2,"")',
                arguments     => { IV1 => $included, IV2 => $llfcs },
                defaultFormat => 'textcopy',
            ),
            Arithmetic(
                name          => $tariffs->{name},
                arithmetic    => '=IF(IV1,IV2,"")',
                arguments     => { IV1 => $included, IV2 => $tariffs },
                defaultFormat => 'textcopy',
            ),
            $tc1, $group,
        ]
      );

    my $classSet = Labelset( list => [ 1 .. 24 ] );

    my $classDorG = Constant(
        rows => $classSet,
        name => 'Demand or generation',
        data => [ split /\n/, <<EOL ] );
Demand
Demand
Demand
Demand
Demand
Demand
Demand
Demand
Demand
Demand
Demand
Demand
Generation
Generation
Generation
Generation
Generation
Generation
Generation
Generation
Generation
Generation
Generation
Generation
EOL

    my $classes = Constant(
        rows => $classSet,
        name => 'Customer group',
        data => [ split /\n/, <<EOL ] );
Demand #0-5 <=15MVA <=10%
Demand #0-5 <=15MVA >10%
Demand #0-5 >15MVA <=10%
Demand #0-5 >15MVA >10%
Demand #5-15 <=15MVA <=10%
Demand #5-15 <=15MVA >10%
Demand #5-15 >15MVA <=10%
Demand #5-15 >15MVA >10%
Demand #15+ <=15MVA <=10%
Demand #15+ <=15MVA >10%
Demand #15+ >15MVA <=10%
Demand #15+ >15MVA >10%
Generation #0-5 <=15MVA <=10%
Generation #0-5 <=15MVA >10%
Generation #0-5 >15MVA <=10%
Generation #0-5 >15MVA >10%
Generation #5-15 <=15MVA <=10%
Generation #5-15 <=15MVA >10%
Generation #5-15 >15MVA <=10%
Generation #5-15 >15MVA >10%
Generation #15+ <=15MVA <=10%
Generation #15+ <=15MVA >10%
Generation #15+ >15MVA <=10%
Generation #15+ >15MVA >10%
EOL

    my $counts = Arithmetic(
        name          => 'Number of tariffs',
        arithmetic    => '=COUNTIF(IV2_IV3,IV1)',
        arguments     => { IV1 => $classes, IV2_IV3 => $group },
        defaultFormat => '0softnz'
    );

    my $totals = Arithmetic(
        name          => 'Total net revenue (£/year)',
        arithmetic    => '=SUMIF(IV2_IV3,IV1,IV8_IV9)',
        defaultFormat => '0softnz',
        arguments     => {
            IV1     => $classes,
            IV2_IV3 => $group,
            IV8_IV9 => $revenue,
        }
    );

    my $averages = Arithmetic(
        name          => 'Average net revenue (£/year)',
        arithmetic    => '=IF(IV5,IV1/IV4,"")',
        defaultFormat => '0softnz',
        arguments     => {
            IV1 => $totals,
            IV4 => $counts,
            IV5 => $counts,
        }
    );

    push @{ $model->{volatilityTables} },
      Columnset(
        name    => 'Volatility summary',
        columns => [ $classDorG, $classes, $counts, $totals, $averages, ]
      );

    push @{ $model->{volatilityTables} },
      Columnset(
        name    => 'Yet another summary',
        columns => [
            Arithmetic(
                name          => 'Total net revenue - demand (£/year)',
                arithmetic    => '=SUMIF(IV2_IV3,"Demand",IV8_IV9)',
                defaultFormat => '0softnz',
                arguments     => {
                    IV2_IV3 => $classDorG,
                    IV8_IV9 => $totals,
                }
            ),
            Arithmetic(
                name          => 'Total net revenue - generation (£/year)',
                arithmetic    => '=SUMIF(IV2_IV3,"Generation",IV8_IV9)',
                defaultFormat => '0softnz',
                arguments     => {
                    IV2_IV3 => $classDorG,
                    IV8_IV9 => $totals,
                }
            ),
            GroupBy(
                name          => 'Total net revenue (£/year)',
                defaultFormat => '0softnz',
                source        => $totals,
            )
        ]
      );

    my $t1 = GroupBy(
        name          => 'Total net revenues under previous method',
        defaultFormat => '0softnz',
        source        => $r1
    );
    my $t2 = GroupBy(
        name          => 'Total net notional forecast revenues in this model',
        defaultFormat => '0softnz',
        source        => $r2
    );

    push @{ $model->{summaryTables} },
      Columnset(
        name    => 'Total net revenues from EDCM users',
        columns => [
            $t1, $t2,
            Arithmetic(
                name          => 'Change',
                arithmetic    => '=IF(IV1,IV3/IV4-1,"")',
                defaultFormat => '%softpm',
                arguments     => { IV1 => $t1, IV3 => $t2, IV4 => $t1 }
            )
        ]
      );

}

1;
