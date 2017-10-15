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

sub preliminaryAdjustments {

    my (
        $model,                            $daysInYear,
        $hoursInPurple,                    $tariffDaysInYearNot,
        $tariffHoursInPurpleNot,           $importCapacity,
        $nonChargeableCapacity,            $activeCoincidence,
        $reactiveCoincidence,              $creditableCapacity,
        $exportCapacityChargeable20052010, $exportCapacityChargeablePost2010,
        $exportCapacityChargeablePre2005,  $exportCapacityExempt,
        $tariffSoleUseMeav,
    ) = @_;

    my $chargeableCapacity = Arithmetic(
        name          => 'Import capacity not subject to DSM (kVA)',
        defaultFormat => '0soft',
        arguments  => { A1 => $importCapacity, A2 => $nonChargeableCapacity, },
        arithmetic => '=A1-A2',
    );

    $importCapacity = Arithmetic(
        name          => 'Maximum import capacity adjusted for part-year (kVA)',
        defaultFormat => '0soft',
        arithmetic    => '=IF(A12="VOID",0,(A1*(1-A2/A3)))',
        arguments     => {
            A1  => $importCapacity,
            A12 => $importCapacity,
            A2  => $tariffDaysInYearNot,
            A3  => $daysInYear,
        },
    );

    $chargeableCapacity = Arithmetic(
        name          => 'Non-DSM import capacity adjusted for part-year (kVA)',
        defaultFormat => '0soft',
        arithmetic    => '=A1-(A11*(1-A2/A3))',
        arguments     => {
            A1  => $importCapacity,
            A11 => $nonChargeableCapacity,
            A2  => $tariffDaysInYearNot,
            A3  => $daysInYear,
        },
    );

    $_ = Arithmetic(
        name          => $_->objectShortName . ' adjusted for part-year',
        defaultFormat => '0soft',
        arithmetic    => '=IF(A12="VOID",0,A1*(1-A2/A3))',
        arguments     => {
            A1  => $_,
            A12 => $_,
            A2  => $tariffDaysInYearNot,
            A3  => $daysInYear,
        }
      )
      foreach $creditableCapacity, $exportCapacityExempt,
      $exportCapacityChargeablePre2005, $exportCapacityChargeable20052010,
      $exportCapacityChargeablePost2010;

    my $exportCapacityChargeable = Arithmetic(
        name      => 'Chargeable export capacity adjusted for part-year (kVA)',
        groupName => 'Export capacities',
        defaultFormat => '0soft',
        arithmetic    => '=A1+A4+A5',
        arguments     => {
            A1 => $exportCapacityChargeablePre2005,
            A4 => $exportCapacityChargeable20052010,
            A5 => $exportCapacityChargeablePost2010,
        },
    );

    $activeCoincidence = Arithmetic(
        name =>
          "$model->{TimebandName} kW divided by kVA adjusted for part-year",
        arithmetic => '=A1*(1-A2/A3)/(1-A4/A5)',
        arguments  => {
            A1 => $activeCoincidence,
            A2 => $tariffHoursInPurpleNot,
            A3 => $hoursInPurple,
            A4 => $tariffDaysInYearNot,
            A5 => $daysInYear,
        }
    );

    $reactiveCoincidence = Arithmetic(
        name =>
          "$model->{TimebandName} kVAr divided by kVA adjusted for part-year",
        arithmetic => '=A1*(1-A2/A3)/(1-A4/A5)',
        arguments  => {
            A1 => $reactiveCoincidence,
            A2 => $tariffHoursInPurpleNot,
            A3 => $hoursInPurple,
            A4 => $tariffDaysInYearNot,
            A5 => $daysInYear,
        }
    ) if $reactiveCoincidence;

    my $demandSoleUseAssetUnscaled = Arithmetic(
        name          => 'Sole use asset MEAV for demand (£)',
        defaultFormat => '0soft',
        arithmetic    => '=IF(A9,A1*A2/(A3+A4+A5),0)',
        arguments     => {
            A1 => $tariffSoleUseMeav,
            A9 => $model->{legacy201}
            ? $tariffSoleUseMeav
            : $importCapacity,
            A2 => $importCapacity,
            A3 => $importCapacity,
            A4 => $exportCapacityExempt,
            A5 => $exportCapacityChargeable,
        }
    );

    my $generationSoleUseAssetUnscaled = Arithmetic(
        name          => 'Sole use asset MEAV for non-exempt generation (£)',
        defaultFormat => '0soft',
        arithmetic    => '=IF(A9,A1*A21/(A3+A4+A5),0)',
        arguments     => {
            A1 => $tariffSoleUseMeav,
            A9 => $model->{legacy201}
            ? $tariffSoleUseMeav
            : $exportCapacityChargeable,
            A3  => $importCapacity,
            A4  => $exportCapacityExempt,
            A5  => $exportCapacityChargeable,
            A21 => $exportCapacityChargeable,
        }
    );

    my $demandSoleUseAsset = Arithmetic(
        name      => 'Demand sole use asset MEAV adjusted for part-year (£)',
        groupName => 'Sole use assets',
        defaultFormat => '0soft',
        arithmetic    => '=A1*(1-A2/A3)',
        arguments     => {
            A1 => $demandSoleUseAssetUnscaled,
            A2 => $tariffDaysInYearNot,
            A3 => $daysInYear,
        },
    );

    my $generationSoleUseAsset = Arithmetic(
        name => 'Generation sole use asset MEAV adjusted for part-year (£)',
        defaultFormat => '0soft',
        arithmetic    => '=A1*(1-A2/A3)',
        arguments     => {
            A1 => $generationSoleUseAssetUnscaled,
            A2 => $tariffDaysInYearNot,
            A3 => $daysInYear,
        }
    );

    $reactiveCoincidence = Arithmetic(
        name       => "$model->{TimebandName} kVAr/agreed kVA (capped)",
        arithmetic => '=MAX(MIN(SQRT(1-MIN(1,A2)^2),'
          . ( $model->{legacy201} ? '' : '0+' )
          . 'A1),0-SQRT(1-MIN(1,A3)^2))',
        arguments => {
            A1 => $reactiveCoincidence,
            A2 => $activeCoincidence,
            A3 => $activeCoincidence,
        }
    );

    $chargeableCapacity,  $exportCapacityChargeable, $importCapacity,
      $activeCoincidence, $reactiveCoincidence,      $creditableCapacity,
      $exportCapacityChargeable20052010, $exportCapacityChargeablePost2010,
      $exportCapacityChargeablePre2005,  $exportCapacityExempt,
      $demandSoleUseAsset,               $generationSoleUseAsset,
      $demandSoleUseAssetUnscaled,       $generationSoleUseAssetUnscaled;

}

1;
