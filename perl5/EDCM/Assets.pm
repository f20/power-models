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

=head Table numbers used in this file

1105
1122
1131
1132
1133
1134
1135
1136
1137

=cut

use warnings;
use strict;
use utf8;

use SpreadsheetModel::Shortcuts ':all';

sub cdcmAssets {

    my ($model) = @_;

    my $assetLevelset = Labelset(
        name => 'All asset levels',
        list => [ split /\n/, <<EOT ] );
GSP
132kV circuits
132kV/EHV
EHV circuits
EHV/HV
132kV/HV
HV circuits
HV/LV
LV circuits
LV customer
HV customer
EOT

    my $allAssets = Dataset(
        name          => 'Assets in CDCM model (£) (from CDCM table 2705)',
        defaultFormat => '0hard',
        cols          => $assetLevelset,
        data          => [ map { 5e8 } @{ $assetLevelset->{list} } ],
        number        => 1131,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        }
    );

    my $ehvAssets = SumProduct(
        name          => 'EHV assets in CDCM model (£)',
        defaultFormat => '0softnz',
        matrix        => Constant(
            name  => 'EHV asset levels',
            cols  => $assetLevelset,
            byrow => 1,
            data  => [ [qw(1 1 1 1 1 1 0 0 0 0 0)] ]
        ),
        vector => $allAssets
    );

    my $hvLvNetAssets = SumProduct(
        name          => 'HV and LV network assets in CDCM model (£)',
        defaultFormat => '0softnz',
        matrix        => Constant(
            name  => 'HV and LV network asset levels',
            cols  => $assetLevelset,
            byrow => 1,
            data  => [ [qw(0 0 0 0 0 0 1 1 1 0 0)] ]
        ),
        vector => $allAssets
    );

    my $hvLvServAssets = SumProduct(
        name          => 'HV and LV service assets in CDCM model (£)',
        defaultFormat => '0softnz',
        matrix        => Constant(
            name  => 'HV and LV service asset levels',
            cols  => $assetLevelset,
            byrow => 1,
            data  => [ [qw(0 0 0 0 0 0 0 0 0 1 1)] ]
        ),
        vector => $allAssets
    );

    0 and push @{ $model->{tablesA} },
      Columnset(
        name    => 'Total CDCM notional assets',
        columns => [ $ehvAssets, $hvLvNetAssets, $hvLvServAssets ]
      );

    $allAssets, $ehvAssets, $hvLvNetAssets, $hvLvServAssets;

}

sub notionalAssets {

    my (
        $model,          $llfcs,             $tariffs,
        $included,       $activeCoincidence, $reactiveCoincidence,
        $agreedCapacity, $exceededCapacity,  $powerFactorInModel,
        $tariffCategory, $tariffDorG,        $tariffSU,
        $cdcmAssets,     $useProportions,    $ehvAssetLevelset,
        $importForGenerator,
      )
      = @_;

    push @{ $model->{tablesG} }, my $customerCategory = Arithmetic(
        name       => 'Tariff type and category',
        arithmetic => '=IF(IV3,IF(IV1="Demand","D","G")&TEXT(IV2,"0000"),"")',
        arguments  =>
          { IV3 => $included, IV1 => $tariffDorG, IV2 => $tariffCategory }
    );

    my $lossFactors = Dataset(
        name => 'Loss adjustment factor to transmission'
          . ' for each network level (from CDCM table 2004)',
        cols       => $ehvAssetLevelset,
        data       => [qw(1 1.01 1.02 1.03 1.04 1.04)],
        number     => 1135,
        dataset    => $model->{dataset},
        appendTo   => $model->{inputTables},
        validation => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        }
    );

    my $cdcmUse = Dataset(
        name => 'Forecast system simultaneous maximum load (kW)'
          . ' from CDCM users (from CDCM table 2506)',
        defaultFormat => '0hardnz',
        cols          => $ehvAssetLevelset,
        data          => [qw(5e6 5e6 5e6 5e6 5e6 5e6)],
        number        => 1122,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        }
    );

    my $diversity = Dataset(
        name => 'Diversity allowance between level exit '
          . 'and GSP Group (from CDCM table 2611)',
        defaultFormat => '%hard',
        cols          => $ehvAssetLevelset,
        data          => [qw(0.1 0.1 0.3 0.3 0.3 0.7)],
        number        => 1105,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables}
      )
      unless $model->{vedcm} == 39 || $model->{vedcm} == 40;

    my $tariffLossFactor = Arithmetic(
        name       => 'Loss factor to transmission',
        arithmetic =>
'=INDEX(IV8_IV9,IF(ISNUMBER(SEARCH("?0000",IV1)),1,IF(ISNUMBER(SEARCH("?1000",IV20)),2,IF(ISNUMBER(SEARCH("??100",IV21)),3,IF(ISNUMBER(SEARCH("???10",IV22)),4,IF(ISNUMBER(SEARCH("??001",IV23)),6,5))))))',
        arguments => {
            IV1     => $customerCategory,
            IV20    => $customerCategory,
            IV21    => $customerCategory,
            IV22    => $customerCategory,
            IV23    => $customerCategory,
            IV8_IV9 => $lossFactors
          }

    );

    my $redUseRate = Arithmetic(
        name =>
'Peak-time active power consumption adjusted to transmission (kW/kVA)',
        arithmetic => '=IF(AND(IV1,IV6="Demand"),IV2*IV9,0)',
        arguments  => {
            IV1 => $included,
            IV6 => $tariffDorG,
            IV2 => $activeCoincidence,
            IV9 => $tariffLossFactor,
        }
    );

    my $capUseRate = Arithmetic(
        name =>
'Active power equivalent of capacity adjusted to transmission (kW/kVA)',
        arithmetic => '=IF(AND(IV1,IV6="Demand"),IV2*IV9,0)',
        arguments  => {
            IV1 => $included,
            IV6 => $tariffDorG,
            IV2 => $powerFactorInModel,
            IV9 => $tariffLossFactor,
        }
    );

    my (
        @assetsCapacity,       @assetsConsumption,
        @assetsCapacityCooked, @assetsConsumptionCooked,
        $useProportionsCooked, $useProportionsCapped,
    );

    my $usePropCaps = Dataset(
        name     => 'Maximum network use factor',
        data     => [ map { 2 } @{ $ehvAssetLevelset->{list} } ],
        cols     => $ehvAssetLevelset,
        number   => $model->{capTableNumber} || 1133,
        dataset  => $model->{dataset},
        appendTo => $model->{inputTables}
      )
      if $model->{vedcm} == 43 || $model->{vedcm} == 44;

    $useProportionsCooked = $model->{cookedNuf} if $model->{cookedNuf};

    $useProportionsCooked = Arithmetic(
        name       => 'Capped network use factors',
        arithmetic => '=MIN(IV1,IV2)',
        arguments  => {
            IV1 => $useProportions,
            IV2 => $usePropCaps,
        }
      )
      if $model->{vedcm} == 43;

    $useProportionsCooked = Arithmetic(
        name       => 'Cooked network use factors (option 44)',
        arithmetic => '=IF(IV4="D1000",MIN(IV1,IV2),1)',
        arguments  => {
            IV1 => $useProportions,
            IV2 => $usePropCaps,
            IV4 => $customerCategory,
        }
      )
      if $model->{vedcm} == 44;

    $useProportionsCooked = Arithmetic(
        name       => 'Cooked network use factors (option 46)',
        arithmetic => '=(1+IV1)/2',
        arguments  => { IV1 => $useProportions, }
      )
      if $model->{vedcm} == 46;

    $useProportionsCooked = Arithmetic(
        name       => 'Cooked network use factors (bands)',
        arithmetic => '=IF(IV1<2,IF(IV2>0.5,1,IF(IV3>0.25,0.5,0.25)),2)',
        arguments  => {
            IV1 => $useProportions,
            IV2 => $useProportions,
            IV3 => $useProportions,
        }
      )
      if $model->{vedcm} == 54
      || $model->{vedcm} == 56
      || $model->{vedcm} == 58
      || $model->{vedcm} == 60;

    if (   $model->{vedcm} == 47
        || $model->{vedcm} == 53
        || $model->{vedcm} == 55
        || $model->{vedcm} == 57
        || $model->{vedcm} == 59 )
    {

        my $usePropCap = Dataset(
            name     => 'Maximum network use factor',
            data     => [ map { 2 } @{ $ehvAssetLevelset->{list} } ],
            cols     => $ehvAssetLevelset,
            number   => 1133,
            dataset  => $model->{dataset},
            appendTo => $model->{inputTables}
        );

        my $usePropCollar = Dataset(
            name     => 'Minimum network use factor',
            data     => [ map { 0.25 } @{ $ehvAssetLevelset->{list} } ],
            cols     => $ehvAssetLevelset,
            number   => 1134,
            dataset  => $model->{dataset},
            appendTo => $model->{inputTables}
        );

        Columnset(
            name     => 'Network use factor cap and collar',
            number   => 1136,
            dataset  => $model->{dataset},
            appendTo => $model->{inputTables},
            columns  => [ $usePropCollar, $usePropCap ]
          )
          if $model->{vedcm} == 47;

        $useProportions =
          $model->{allowInvalid}
          ? Arithmetic(
            name       => 'Network use factors (first set)',
            arithmetic => '=IF(IV4,0,IV1)',
            arguments  => {
                IV1 => $useProportions,
                IV4 => $importForGenerator,
            }
          )
          : Arithmetic(
            name       => 'Network use factors (first set)',
            arithmetic => '=IF(IV4,IV2,IV1)',
            arguments  => {
                IV1 => $useProportions,
                IV2 => $usePropCollar,
                IV4 => $importForGenerator,
            }
          );

        $useProportionsCapped = Arithmetic(
            name       => 'Network use factors (capped only)',
            arithmetic => '=MIN(IV1,IV2)',
            arguments  => {
                IV1 => $useProportions,
                IV2 => $usePropCap,
            }
          )
          if $model->{allowInvalid};

        $useProportionsCooked = Arithmetic(
            name       => 'Network use factors (second set)',
            arithmetic => '=MAX(IV3,MIN(IV1,IV2))',
            arguments  => {
                IV1 => $useProportions,
                IV2 => $usePropCap,
                IV3 => $usePropCollar,
            }
        );

    }

    my $assetsFixed = Arithmetic(
        name       => 'Sole use assets for demand only (£)',
        arithmetic => '=IF(AND(IV1,IV3="Demand"),IV2,0)',
        arguments => { IV1 => $included, IV2 => $tariffSU, IV3 => $tariffDorG },
        defaultFormat => '0softnz',
    );

    push @{ $model->{tablesD} }, my $accretion = Arithmetic(
        name       => 'Notional asset rate (£/kW)',
        arithmetic => '=IF(IV1,IV2/IV3/IV4,0)',
        arguments  => {
            IV1 => $cdcmUse,
            IV2 => $cdcmAssets,
            IV3 => $cdcmUse,
            IV4 => $lossFactors
        }
    );

    my $accretion132hvHard = Dataset(
        name          => 'Override notional asset rate for 132kV/HV (£/kW)',
        data          => ['#VALUE!'],
        cols          => Labelset(
            list =>
              [ $ehvAssetLevelset->{list}[ $#{ $ehvAssetLevelset->{list} } ] ]
        ),
        number   => 1132,
        dataset  => $model->{dataset},
        appendTo => $model->{inputTables},
    );

    my $accretion132hvcombined = Arithmetic(
        name       => 'Notional asset rate for 132kV/HV (£/kW)',
        arithmetic => '=IF(ISNUMBER(IV1),IV2,IV3)',
        arguments  => {
            IV1 => $accretion132hvHard,
            IV2 => $accretion132hvHard,
            IV3 => $accretion,
        }
    );

    0
      and $accretion132hvcombined = Arithmetic(
        name       => 'Sense checked notional asset rate for 132kV/HV (£/kW)',
        arithmetic => '=IF(IV1>0,IV2,742/0)',
        arguments  => {
            IV1 => $accretion132hvcombined,
            IV2 => $accretion132hvcombined,
        }
      );

    $accretion = Stack(
        name    => 'Notional asset rate adjusted (£/kW)',
        cols    => $ehvAssetLevelset,
        sources => [ $accretion132hvcombined, $accretion ]
    );

    my $starIV5   = $useProportions       ? '*IV5' : '';
    my $starIV5nn = $useProportionsCooked ? '*IV5' : '';

    if ( !$model->{allowInvalid} ) {

        push @assetsCapacity,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[1] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic => qq@=IF(IV1="D1000",IV4*IV8/(1+IV3)$starIV5,0)@,
            arguments  => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV3 => $diversity,
                IV8 => $capUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[1] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic => qq@=IF(IV1="D1000",IV4*IV8/(1+IV3)$starIV5nn,0)@,
            arguments  => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV3 => $diversity,
                IV8 => $capUseRate,
                $useProportionsCooked ? ( IV5 => $useProportionsCooked ) : (),
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[2] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
              qq@=IF(ISNUMBER(SEARCH("D?100",IV1)),IV4*IV8/(1+IV3)$starIV5,0)@,
            arguments => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV3 => $diversity,
                IV8 => $capUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[2] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?100",IV1)),IV4*IV8/(1+IV3)$starIV5nn,0)@,
            arguments => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV3 => $diversity,
                IV8 => $capUseRate,
                $useProportionsCooked ? ( IV5 => $useProportionsCooked ) : (),
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[3] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
              qq@=IF(ISNUMBER(SEARCH("D??10",IV1)),IV4*IV8/(1+IV3)$starIV5,0)@,
            arguments => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV3 => $diversity,
                IV8 => $capUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[3] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D??10",IV1)),IV4*IV8/(1+IV3)$starIV5nn,0)@,
            arguments => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV3 => $diversity,
                IV8 => $capUseRate,
                $useProportionsCooked ? ( IV5 => $useProportionsCooked ) : (),
            },
          );

        if ( $model->{wrong} ) {
            push @assetsCapacity,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[4] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
                arithmetic =>
qq@=IF(OR(IV6="D0002",ISNUMBER(SEARCH("D?1?1",IV1))),IV4*IV8/(1+IV3)$starIV5,0)@,
                arguments => {
                    IV1 => $customerCategory,
                    IV4 => $accretion,
                    IV6 => $customerCategory,
                    IV3 => $diversity,
                    IV8 => $capUseRate,
                    $useProportions ? ( IV5 => $useProportions ) : (),
                },
              );

            push @assetsCapacityCooked,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[4] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
                arithmetic =>
qq@=IF(OR(IV6="D0002",ISNUMBER(SEARCH("D?1?1",IV1))),IV4*IV8/(1+IV3)$starIV5nn,0)@,
                arguments => {
                    IV1 => $customerCategory,
                    IV4 => $accretion,
                    IV6 => $customerCategory,
                    IV3 => $diversity,
                    IV8 => $capUseRate,
                    $useProportionsCooked
                    ? ( IV5 => $useProportionsCooked )
                    : (),
                },
              );

            push @assetsCapacity,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[5] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[5] ] ),
                arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?0?1",IV1)),IV4*IV8/(1+IV3)$starIV5,0)@,
                arguments => {
                    IV1 => $customerCategory,
                    IV4 => $accretion,
                    IV3 => $diversity,
                    IV8 => $capUseRate,
                    $useProportions ? ( IV5 => $useProportions ) : (),
                },
              );

            push @assetsCapacityCooked,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[5] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[5] ] ),
                arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?0?1",IV1)),IV4*IV8/(1+IV3)$starIV5nn,0)@,
                arguments => {
                    IV1 => $customerCategory,
                    IV4 => $accretion,
                    IV3 => $diversity,
                    IV8 => $capUseRate,
                    $useProportionsCooked
                    ? ( IV5 => $useProportionsCooked )
                    : (),
                },
              );

        }
        else {

            push @assetsCapacity,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[4] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
                arithmetic =>
qq@=IF(OR(IV6="D0002",ISNUMBER(SEARCH("D?1?1",IV1)),ISNUMBER(SEARCH("D??11",IV7))),IV4*IV8/(1+IV3)$starIV5,0)@,
                arguments => {
                    IV1 => $customerCategory,
                    IV4 => $accretion,
                    IV6 => $customerCategory,
                    IV7 => $customerCategory,
                    IV3 => $diversity,
                    IV8 => $capUseRate,
                    $useProportions ? ( IV5 => $useProportions ) : (),
                },
              );

            push @assetsCapacityCooked,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[4] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
                arithmetic =>
qq@=IF(OR(IV6="D0002",ISNUMBER(SEARCH("D?1?1",IV1)),ISNUMBER(SEARCH("D??11",IV7))),IV4*IV8/(1+IV3)$starIV5nn,0)@,
                arguments => {
                    IV1 => $customerCategory,
                    IV7 => $customerCategory,
                    IV4 => $accretion,
                    IV6 => $customerCategory,
                    IV3 => $diversity,
                    IV8 => $capUseRate,
                    $useProportionsCooked
                    ? ( IV5 => $useProportionsCooked )
                    : (),
                },
              );

            push @assetsCapacity,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[5] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[5] ] ),
                arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?001",IV1)),IV4*IV8/(1+IV3)$starIV5,0)@,
                arguments => {
                    IV1 => $customerCategory,
                    IV4 => $accretion,
                    IV3 => $diversity,
                    IV8 => $capUseRate,
                    $useProportions ? ( IV5 => $useProportions ) : (),
                },
              );

            push @assetsCapacityCooked,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[5] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[5] ] ),
                arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?001",IV1)),IV4*IV8/(1+IV3)$starIV5nn,0)@,
                arguments => {
                    IV1 => $customerCategory,
                    IV4 => $accretion,
                    IV3 => $diversity,
                    IV8 => $capUseRate,
                    $useProportionsCooked
                    ? ( IV5 => $useProportionsCooked )
                    : (),
                },
              );
        }

        push @assetsConsumption,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[1] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic =>
qq@=IF(IV1="D1000",0,IF(ISNUMBER(SEARCH("D1???",IV2)),IV4*IV9$starIV5,0))@,
            arguments => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV9 => $redUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[1] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic =>
qq@=IF(IV1="D1000",0,IF(ISNUMBER(SEARCH("D1???",IV2)),IV4*IV9$starIV5nn,0))@,
            arguments => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV9 => $redUseRate,
                $useProportionsCooked ? ( IV5 => $useProportionsCooked ) : (),
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[2] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?100",IV1)),0,IF(ISNUMBER(SEARCH("D?1??",IV2)),IV4*IV9$starIV5,0))@,
            arguments => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV9 => $redUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[2] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?100",IV1)),0,IF(ISNUMBER(SEARCH("D?1??",IV2)),IV4*IV9$starIV5nn,0))@,
            arguments => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV9 => $redUseRate,
                $useProportionsCooked ? ( IV5 => $useProportionsCooked ) : (),
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[3] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D??10",IV1)),0,IF(ISNUMBER(SEARCH("D??1?",IV2)),IV4*IV9$starIV5,0))@,
            arguments => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV9 => $redUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[3] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D??10",IV1)),0,IF(ISNUMBER(SEARCH("D??1?",IV2)),IV4*IV9$starIV5nn,0))@,
            arguments => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV9 => $redUseRate,
                $useProportionsCooked ? ( IV5 => $useProportionsCooked ) : (),
            },
          );

    }

    else {    # if allowInvalid

        push @assetsCapacity,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[1] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic => qq@=IF(IV1="D1000",IV4*IV8/(1+IV3)$starIV5,0)@,
            arguments  => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV3 => $diversity,
                IV8 => $capUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[1] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic => qq@=IF(IV1="D1000",IV4*IV8/(1+IV3)$starIV5nn,0)@,
            arguments  => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV3 => $diversity,
                IV8 => $capUseRate,
                IV5 => $useProportionsCooked,
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[1] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic =>
qq@=IF(OR(IV1="D1000",ISNUMBER(SEARCH("G????",IV20))),0,IV4*IV9$starIV5)@,
            arguments => {
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => $redUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[1] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic =>
qq@=IF(OR(IV1="D1000",ISNUMBER(SEARCH("G????",IV20))),0,IF(ISNUMBER(SEARCH("D1???",IV21)),IV5,IV6)*IV4*IV9)@,
            arguments => {
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => $redUseRate,
                IV5  => $useProportionsCooked,
                IV6  => $useProportionsCapped,
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[2] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
              qq@=IF(ISNUMBER(SEARCH("D?100",IV1)),IV4*IV8/(1+IV3)$starIV5,0)@,
            arguments => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV3 => $diversity,
                IV8 => $capUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[2] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?100",IV1)),IV4*IV8/(1+IV3)$starIV5nn,0)@,
            arguments => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV3 => $diversity,
                IV8 => $capUseRate,
                IV5 => $useProportionsCooked,
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[2] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",IV20)),ISNUMBER(SEARCH("D?100",IV1))),0,IV4*IV9$starIV5)@,
            arguments => {
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => $redUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[2] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",IV20)),ISNUMBER(SEARCH("D?100",IV1))),0,IF(ISNUMBER(SEARCH("D?1??",IV21)),IV5,IV6)*IV4*IV9)@,
            arguments => {
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => $redUseRate,
                IV5  => $useProportionsCooked,
                IV6  => $useProportionsCapped,
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[3] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
              qq@=IF(ISNUMBER(SEARCH("D??10",IV1)),IV4*IV8/(1+IV3)$starIV5,0)@,
            arguments => {
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV3  => $diversity,
                IV8  => $capUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[3] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D??10",IV1)),IV4*IV8/(1+IV3)$starIV5nn,0)@,
            arguments => {
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV3  => $diversity,
                IV8  => $capUseRate,
                IV5  => $useProportionsCooked,
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[3] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",IV20)),ISNUMBER(SEARCH("D??10",IV1))),0,IV4*IV9$starIV5)@,
            arguments => {
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => $redUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[3] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",IV20)),ISNUMBER(SEARCH("D??10",IV1))),0,IF(ISNUMBER(SEARCH("D??1?",IV21)),IV5,IV6)*IV4*IV9)@,
            arguments => {
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => $redUseRate,
                IV5  => $useProportionsCooked,
                IV6  => $useProportionsCapped,
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[4] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
            arithmetic =>
qq@=IF(OR(IV6="D0002",ISNUMBER(SEARCH("D?1?1",IV1)),ISNUMBER(SEARCH("D??11",IV7))),IV4*IV8/(1+IV3)$starIV5,0)@,
            arguments => {
                IV1 => $customerCategory,
                IV4 => $accretion,
                IV6 => $customerCategory,
                IV7 => $customerCategory,
                IV3 => $diversity,
                IV8 => $capUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[4] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
            arithmetic =>
qq@=IF(OR(IV6="D0002",ISNUMBER(SEARCH("D?1?1",IV1)),ISNUMBER(SEARCH("D??11",IV7))),IV4*IV8/(1+IV3)$starIV5nn,0)@,
            arguments => {
                IV1 => $customerCategory,
                IV7 => $customerCategory,
                IV4 => $accretion,
                IV6 => $customerCategory,
                IV3 => $diversity,
                IV8 => $capUseRate,
                IV5 => $useProportionsCooked,
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[4] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",IV20)),IV22="D0002",ISNUMBER(SEARCH("D?1?1",IV1)),ISNUMBER(SEARCH("D??11",IV7))),0,IV4*IV9$starIV5)@,
            arguments => {
                IV1  => $customerCategory,
                IV22 => $customerCategory,
                IV7  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => $redUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[4] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",IV20)),IV22="D0002",ISNUMBER(SEARCH("D?1?1",IV1)),ISNUMBER(SEARCH("D??11",IV7))),0,IV6*IV4*IV9)@,
            arguments => {
                IV1  => $customerCategory,
                IV22 => $customerCategory,
                IV7  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => $redUseRate,
                IV5  => $useProportionsCooked,
                IV6  => $useProportionsCapped,
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[5] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[5] ] ),
            arithmetic =>
              qq@=IF(ISNUMBER(SEARCH("D?001",IV1)),IV4*IV8/(1+IV3)$starIV5,0)@,
            arguments => {
                IV1 => $customerCategory,
                IV4 => $accretion,
                IV3 => $diversity,
                IV8 => $capUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsCapacityCooked,
          Arithmetic(
            name       => "Capacity $accretion->{cols}{list}[5] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[5] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?001",IV1)),IV4*IV8/(1+IV3)$starIV5nn,0)@,
            arguments => {
                IV1 => $customerCategory,
                IV4 => $accretion,
                IV3 => $diversity,
                IV8 => $capUseRate,
                IV5 => $useProportionsCooked,
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[5] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[5] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",IV20)),ISNUMBER(SEARCH("D?001",IV1))),0,IV4*IV9$starIV5)@,
            arguments => {
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => $redUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name       => "Consumption $accretion->{cols}{list}[5] (£/kVA)",
            cols       => Labelset( list => [ $accretion->{cols}{list}[5] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",IV20)),ISNUMBER(SEARCH("D?001",IV1))),0,IV6*IV4*IV9)@,
            arguments => {
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => $redUseRate,
                IV5  => $useProportionsCooked,
                IV6  => $useProportionsCapped,
            },
          );

    }

    push @assetsCapacity,
      Arithmetic(
        name       => 'Total notional capacity assets (£/kVA)',
        cols       => 0,
        arithmetic => '=' . join( '+', map { "IV$_" } 1 .. @assetsCapacity ),
        arguments  => {
            map { ( "IV$_" => $assetsCapacity[ $_ - 1 ] ) } 1 .. @assetsCapacity
        },
      );

    push @assetsCapacityCooked,
      Arithmetic(
        name       => 'Second set of capacity assets (£/kVA)',
        cols       => 0,
        arithmetic => '='
          . join( '+', map { "IV$_" } 1 .. @assetsCapacityCooked ),
        arguments => {
            map { ( "IV$_" => $assetsCapacityCooked[ $_ - 1 ] ) }
              1 .. @assetsCapacityCooked
        },
      );

    push @assetsConsumption,
      Arithmetic(
        name       => 'Total notional consumption assets (£/kVA)',
        cols       => 0,
        arithmetic => '=' . join( '+', map { "IV$_" } 1 .. @assetsConsumption ),
        arguments  => {
            map { ( "IV$_" => $assetsConsumption[ $_ - 1 ] ) }
              1 .. @assetsConsumption
        },
      );

    push @assetsConsumptionCooked,
      Arithmetic(
        name       => 'Second set of consumption assets (£/kVA)',
        cols       => 0,
        arithmetic => '='
          . join( '+', map { "IV$_" } 1 .. @assetsConsumptionCooked ),
        arguments => {
            map { ( "IV$_" => $assetsConsumptionCooked[ $_ - 1 ] ) }
              1 .. @assetsConsumptionCooked
        },
      );

    0 and push @{ $model->{matrixTables} },
      GroupBy(
        name          => 'Number of tariffs directly affected by cooking',
        defaultFormat => '0soft',
        source        => Arithmetic(
            name          => 'Affected by capping?',
            defaultFormat => '0softnz',
            arithmetic    => '=IF(AND(IV1=IV2,IV3=IV4),0,1)',
            arguments     => {
                IV1 => $assetsCapacityCooked[$#assetsCapacityCooked],
                IV2 => $assetsCapacity[$#assetsCapacity],
                IV3 => $assetsConsumptionCooked[$#assetsConsumptionCooked],
                IV4 => $assetsConsumption[$#assetsConsumption],
            },
        ),
      );

    my $totalAssetsFixed = GroupBy(
        source        => $assetsFixed,
        name          => 'Total sole use assets for demand (£)',
        defaultFormat => '0softnz'
    );

    my @totalAssetsCapacity = map {
        my $name = $_->objectShortName;
        $name =~ s/\(£\/kVA\)/(£)/;
        SumProduct(
            name          => $name,
            defaultFormat => '0softnz',
            matrix        => $_,
            vector        => $agreedCapacity
        );
    } @assetsCapacity;

    my @totalAssetsConsumption = map {
        my $name = $_->objectShortName;
        $name =~ s/\(£\/kVA\)/(£)/;
        SumProduct(
            name          => $name,
            defaultFormat => '0softnz',
            matrix        => $_,
            vector        => $agreedCapacity
        );
    } @assetsConsumption;

    my ($totalAssetsGenerationSoleUse) = map {
        GroupBy(
            source        => $_,
            name          => $_->objectShortName . ' (aggregate)',
            defaultFormat => $_->{defaultFormat}
        );
      } Arithmetic(
        name       => 'Sole use assets for generation only (£)',
        arithmetic => '=IF(AND(IV1,IV3="Generation"),IV2,0)',
        arguments => { IV1 => $included, IV2 => $tariffSU, IV3 => $tariffDorG },
        defaultFormat => '0softnz',
      );

    0 and push @{ $model->{assetTables} },    #$included,
      Columnset(
        name    => 'EDCM notional assets for each tariff',
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
            $assetsFixed,
            @assetsCapacity,
            @assetsConsumption,
            $totalAssetsGenerationSoleUse->{source},
        ],
      );

    0
      and push @{ $model->{assetTables} },
      Columnset(
        name    => 'Second set of notional assets for each tariff',
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
            @assetsCapacityCooked,
            @assetsConsumptionCooked,
        ],
      )
      if $model->{vedcm} == 42
      || $model->{vedcm} > 52 && $model->{vedcm} < 61
      || $model->{vedcm} == 43
      || $model->{vedcm} == 44
      || $model->{vedcm} == 45
      || $model->{vedcm} == 46
      || $model->{vedcm} == 47
      || $model->{vedcm} == 48
      || $model->{vedcm} == 49
      || $model->{vedcm} == 50
      || $model->{vedcm} == 51
      || $model->{vedcm} == 52;

    push @{ $model->{tablesD} }, my $totalAssets = Arithmetic(
        name          => 'All notional assets in EDCM (£)',
        arithmetic    => '=IV5+IV6+IV7+IV8',
        defaultFormat => '0softnz',
        arguments     => {
            IV5 => $totalAssetsFixed,
            IV6 => $totalAssetsCapacity[$#assetsCapacity],
            IV7 => $totalAssetsConsumption[$#assetsConsumption],
            IV8 => $totalAssetsGenerationSoleUse,
        }
    );

    0 and push @{ $model->{assetTables} },
      Columnset(
        name    => 'Total EDCM notional assets',
        columns => [
            Constant( name => ' ', data => [ [] ] ),
            Constant( name => ' ', data => [ [] ] ),
            $totalAssetsFixed,
            @totalAssetsCapacity,
            @totalAssetsConsumption,
            $totalAssetsGenerationSoleUse,
            $totalAssets
        ]
      );

    my $assetsCapacityDoubleCooked =
      $assetsCapacityCooked[$#assetsCapacityCooked];

    $cdcmUse, $lossFactors, $diversity, $redUseRate, $capUseRate, $assetsFixed,
      $assetsCapacity[$#assetsCapacity],
      $assetsConsumption[$#assetsConsumption], $totalAssetsFixed,
      $totalAssetsCapacity[$#assetsCapacity],
      $totalAssetsConsumption[$#assetsConsumption],
      $totalAssetsGenerationSoleUse, $totalAssets,
      $assetsCapacityCooked[$#assetsCapacityCooked],
      $assetsConsumptionCooked[$#assetsConsumptionCooked],
      $assetsCapacityDoubleCooked,
      $assetsConsumptionCooked[$#assetsConsumptionCooked],;

}

1;
