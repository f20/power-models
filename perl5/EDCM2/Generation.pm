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
use SpreadsheetModel::SegmentRoot;

sub gCharge {

    my (
        $model,                            $genPot20p,
        $genPotGP,                         $genPotGL,
        $genPotCdcmCap20052010,            $genPotCdcmCapPost2010,
        $exportCapacityChargeable,         $exportCapacityChargeable20052010,
        $exportCapacityChargeablePost2010, $daysInYear,
    ) = @_;

    $_ = GroupBy(
        name          => $_->objectShortName . ' (total)',
        defaultFormat => '0softnz',
        source        => $_
      )
      foreach $exportCapacityChargeable, $exportCapacityChargeable20052010,
      $exportCapacityChargeablePost2010;

    Arithmetic(
        name          => 'Export capacity charge p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic =>
'=(IV1*(1-IV21/IV23)+(IV3*IV22/(IV24+IV52)+IV4*IV25/(IV26+IV51))/IV231)*100/IV9',
        arguments => {
            IV1   => $genPot20p,
            IV21  => $exportCapacityChargeable20052010,
            IV22  => $exportCapacityChargeablePost2010,
            IV23  => $exportCapacityChargeable,
            IV231 => $exportCapacityChargeable,
            IV24  => $exportCapacityChargeablePost2010,
            IV25  => $exportCapacityChargeable20052010,
            IV26  => $exportCapacityChargeable20052010,
            IV3   => $genPotGP,
            IV4   => $genPotGL,
            IV51  => $genPotCdcmCap20052010,
            IV52  => $genPotCdcmCapPost2010,
            IV9   => $daysInYear,
        }
    );

}

1;
