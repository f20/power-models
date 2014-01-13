package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2012 Energy Networks Association Limited and others.
Copyright 2013 Franck Latrémolière, Reckon LLP and others.

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
        name => 'Assets in CDCM model (£) (from CDCM table 2705 or 2706)',
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
    $model->{transparency}{oli}{1233} = $ehvAssets if $model->{transparency};

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
    $model->{transparency}{oli}{1235} = $hvLvNetAssets
      if $model->{transparency};

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
    $model->{transparency}{oli}{1231} = $hvLvServAssets
      if $model->{transparency};

    $allAssets, $ehvAssets, $hvLvNetAssets, $hvLvServAssets;

}

sub notionalAssets {

    my (
        $model,          $activeCoincidence,  $reactiveCoincidence,
        $agreedCapacity, $powerFactorInModel, $tariffCategory,
        $tariffSUimport, $tariffSUexport,     $cdcmAssets,
        $useProportions, $ehvAssetLevelset,
    ) = @_;

    my $customerCategory;

    # this should only run if $model->{legacy201}
    push @{ $model->{tablesG} },
      $customerCategory = Arithmetic(
        name       => 'Tariff type and category',
        arithmetic => '="D"&TEXT(IV1,"0000")',
        arguments  => { IV1 => $tariffCategory }
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
    );

    my $tariffLossFactor = Arithmetic(
        name => 'Loss factor to transmission',
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
        name => 'Peak-time active power consumption'
          . ' adjusted to transmission (kW/kVA)',
        arithmetic => '=IV1*IV9',
        arguments  => {
            IV1 => $activeCoincidence,
            IV9 => $tariffLossFactor,
        }
    );

    $redUseRate = [
        $redUseRate,
        Arithmetic(
            name => 'Peak-time capacity use adjusted to transmission (kW/kVA)',
            arithmetic => '=SQRT(IV1*IV2+IV3*IV4)*IV8*IV9',
            arguments  => {
                IV1 => $activeCoincidence,
                IV2 => $activeCoincidence,
                IV3 => $reactiveCoincidence,
                IV4 => $reactiveCoincidence,
                IV8 => $powerFactorInModel,
                IV9 => $tariffLossFactor,
            }
        )
      ]
      if $model->{dcp183};

    my $capUseRate = Arithmetic(
        name => 'Active power equivalent of capacity'
          . ' adjusted to transmission (kW/kVA)',
        arithmetic => '=IV1*IV9',
        arguments  => {
            IV9 => $powerFactorInModel,
            IV1 => $tariffLossFactor,
        }
    );

    my (
        @assetsCapacity,       @assetsConsumption,
        @assetsCapacityCooked, @assetsConsumptionCooked,
        $useProportionsCooked, $useProportionsCapped,
    );

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

    $useProportionsCapped = Arithmetic(
        name       => 'Network use factors (capped only)',
        arithmetic => '=MIN(IV1,IV2)',
        arguments  => {
            IV1 => $useProportions,
            IV2 => $usePropCap,
        }
    ) if $model->{allowInvalid};

    $useProportionsCooked = Arithmetic(
        name       => 'Network use factors (second set)',
        arithmetic => '=MAX(IV3+0,MIN(IV1+0,IV2+0))',
        arguments  => {
            IV1 => $useProportions,
            IV2 => $usePropCap,
            IV3 => $usePropCollar,
        }
    );

    push @{ $model->{calc1Tables} },
      my $accretion = Arithmetic(
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
        name => 'Override notional asset rate for 132kV/HV (£/kW)',
        data => ['#VALUE!'],
        cols => Labelset(
            list =>
              [ $ehvAssetLevelset->{list}[ $#{ $ehvAssetLevelset->{list} } ] ]
        ),
        number   => 1132,
        dataset  => $model->{dataset},
        appendTo => $model->{inputTables},
    );

    my $accretion132hvcombined = Arithmetic(
        name => 'Notional asset rate for 132kV/HV (£/kW)',
        $model->{legacy201}
        ? (
            arithmetic => '=IF(ISNUMBER(IV1),IV2,IV3)',
            arguments  => {
                IV1 => $accretion132hvHard,
                IV2 => $accretion132hvHard,
                IV3 => $accretion,
            }
          )
        : (
            arithmetic => '=IF(ISNUMBER(IV1),IF(IV2,IV3,IV4),IV5)',
            arguments  => {
                IV1 => $accretion132hvHard,
                IV2 => $accretion132hvHard,
                IV3 => $accretion132hvHard,
                IV4 => $accretion,
                IV5 => $accretion,
            }
        )
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
            name => "Capacity $accretion->{cols}{list}[2] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
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
            name => "Capacity $accretion->{cols}{list}[2] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
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
            name => "Capacity $accretion->{cols}{list}[3] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
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
            name => "Capacity $accretion->{cols}{list}[3] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
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
            name => "Consumption $accretion->{cols}{list}[1] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic =>
qq@=IF(IV1="D1000",0,IF(ISNUMBER(SEARCH("D1???",IV2)),IV4*IV9$starIV5,0))@,
            arguments => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV9 => ref $redUseRate eq 'ARRAY'
                ? $redUseRate->[1]
                : $redUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[1] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic =>
qq@=IF(IV1="D1000",0,IF(ISNUMBER(SEARCH("D1???",IV2)),IV4*IV9$starIV5nn,0))@,
            arguments => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV9 => ref $redUseRate eq 'ARRAY'
                ? $redUseRate->[1]
                : $redUseRate,
                $useProportionsCooked ? ( IV5 => $useProportionsCooked ) : (),
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[2] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?100",IV1)),0,IF(ISNUMBER(SEARCH("D?1??",IV2)),IV4*IV9$starIV5,0))@,
            arguments => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV9 => ref $redUseRate eq 'ARRAY'
                ? $redUseRate->[1]
                : $redUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[2] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?100",IV1)),0,IF(ISNUMBER(SEARCH("D?1??",IV2)),IV4*IV9$starIV5nn,0))@,
            arguments => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV9 => ref $redUseRate eq 'ARRAY'
                ? $redUseRate->[1]
                : $redUseRate,
                $useProportionsCooked ? ( IV5 => $useProportionsCooked ) : (),
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[3] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D??10",IV1)),0,IF(ISNUMBER(SEARCH("D??1?",IV2)),IV4*IV9$starIV5,0))@,
            arguments => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV9 => ref $redUseRate eq 'ARRAY'
                ? $redUseRate->[1]
                : $redUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[3] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D??10",IV1)),0,IF(ISNUMBER(SEARCH("D??1?",IV2)),IV4*IV9$starIV5nn,0))@,
            arguments => {
                IV1 => $customerCategory,
                IV2 => $customerCategory,
                IV4 => $accretion,
                IV9 => ref $redUseRate eq 'ARRAY'
                ? $redUseRate->[1]
                : $redUseRate,
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
            name => "Consumption $accretion->{cols}{list}[1] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic =>
qq@=IF(OR(IV1="D1000",ISNUMBER(SEARCH("G????",IV20))),0,IV4*IV9$starIV5)@,
            arguments => {
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => ref $redUseRate eq 'ARRAY'
                ? $redUseRate->[1]
                : $redUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[1] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
            arithmetic =>
qq@=IF(OR(IV1="D1000",ISNUMBER(SEARCH("G????",IV20))),0,IF(ISNUMBER(SEARCH("D1???",IV21)),IV5,IV6)*IV4*IV9)@,
            arguments => {
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => ref $redUseRate eq 'ARRAY'
                ? $redUseRate->[1]
                : $redUseRate,
                IV5 => $useProportionsCooked,
                IV6 => $useProportionsCapped,
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[2] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
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
            name => "Capacity $accretion->{cols}{list}[2] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
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
            name => "Consumption $accretion->{cols}{list}[2] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",IV20)),ISNUMBER(SEARCH("D?100",IV1))),0,IV4*IV9$starIV5)@,
            arguments => {
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => ref $redUseRate eq 'ARRAY'
                ? $redUseRate->[1]
                : $redUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[2] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",IV20)),ISNUMBER(SEARCH("D?100",IV1))),0,IF(ISNUMBER(SEARCH("D?1??",IV21)),IV5,IV6)*IV4*IV9)@,
            arguments => {
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => ref $redUseRate eq 'ARRAY'
                ? $redUseRate->[1]
                : $redUseRate,
                IV5 => $useProportionsCooked,
                IV6 => $useProportionsCapped,
            },
          );

        push @assetsCapacity,
          Arithmetic(
            name => "Capacity $accretion->{cols}{list}[3] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
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
            name => "Capacity $accretion->{cols}{list}[3] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
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
            name => "Consumption $accretion->{cols}{list}[3] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",IV20)),ISNUMBER(SEARCH("D??10",IV1))),0,IV4*IV9$starIV5)@,
            arguments => {
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => ref $redUseRate eq 'ARRAY'
                ? $redUseRate->[1]
                : $redUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[3] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",IV20)),ISNUMBER(SEARCH("D??10",IV1))),0,IF(ISNUMBER(SEARCH("D??1?",IV21)),IV5,IV6)*IV4*IV9)@,
            arguments => {
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => ref $redUseRate eq 'ARRAY'
                ? $redUseRate->[1]
                : $redUseRate,
                IV5 => $useProportionsCooked,
                IV6 => $useProportionsCapped,
            },
          );

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
                IV5 => $useProportionsCooked,
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[4] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",IV20)),IV22="D0002",ISNUMBER(SEARCH("D?1?1",IV1)),ISNUMBER(SEARCH("D??11",IV7))),0,IV4*IV9$starIV5)@,
            arguments => {
                IV1  => $customerCategory,
                IV22 => $customerCategory,
                IV7  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => ref $redUseRate eq 'ARRAY'
                ? $redUseRate->[1]
                : $redUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[4] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",IV20)),IV22="D0002",ISNUMBER(SEARCH("D?1?1",IV1)),ISNUMBER(SEARCH("D??11",IV7))),0,IV6*IV4*IV9)@,
            arguments => {
                IV1  => $customerCategory,
                IV22 => $customerCategory,
                IV7  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => ref $redUseRate eq 'ARRAY'
                ? $redUseRate->[1]
                : $redUseRate,
                IV5 => $useProportionsCooked,
                IV6 => $useProportionsCapped,
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
                IV5 => $useProportionsCooked,
            },
          );

        push @assetsConsumption,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[5] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[5] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",IV20)),ISNUMBER(SEARCH("D?001",IV1))),0,IV4*IV9$starIV5)@,
            arguments => {
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => ref $redUseRate eq 'ARRAY'
                ? $redUseRate->[1]
                : $redUseRate,
                $useProportions ? ( IV5 => $useProportions ) : (),
            },
          );

        push @assetsConsumptionCooked,
          Arithmetic(
            name => "Consumption $accretion->{cols}{list}[5] (£/kVA)",
            cols => Labelset( list => [ $accretion->{cols}{list}[5] ] ),
            arithmetic =>
qq@=IF(OR(ISNUMBER(SEARCH("G????",IV20)),ISNUMBER(SEARCH("D?001",IV1))),0,IV6*IV4*IV9)@,
            arguments => {
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV4  => $accretion,
                IV9  => ref $redUseRate eq 'ARRAY'
                ? $redUseRate->[1]
                : $redUseRate,
                IV5 => $useProportionsCooked,
                IV6 => $useProportionsCapped,
            },
          );

    }

    push @assetsCapacity,
      Arithmetic(
        name       => 'Total notional capacity assets (£/kVA)',
        cols       => 0,
        arithmetic => '=' . join( '+', map { "IV$_" } 1 .. @assetsCapacity ),
        arguments  => {
            map { ( "IV$_" => $assetsCapacity[ $_ - 1 ] ) }
              1 .. @assetsCapacity
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

    my $totalAssetsFixed =
      $model->{transparency}
      ? (
        $model->{transparency}{olo}{119301} = Arithmetic(
            name          => 'Total sole use assets for demand (£)',
            defaultFormat => '0softnz',
            arithmetic    => '=IF(IV123,0,IV1)+SUMPRODUCT(IV11_IV12,IV15_IV16)',
            arguments     => {
                IV123     => $model->{transparencyMasterFlag},
                IV1       => $model->{transparency}{ol119301},
                IV11_IV12 => $tariffSUimport,
                IV15_IV16 => $model->{transparency},
            },
        )
      )
      : GroupBy(
        source        => $tariffSUimport,
        name          => 'Total sole use assets for demand (£)',
        defaultFormat => '0softnz'
      );

    my ( $totalAssetsCapacity, $totalAssetsConsumption ) =
      $model->{transparency}
      ? (
        map {
            my $name = $_->[0]->objectShortName;
            $name =~ s/\(£\/kVA\)/(£)/;
            $model->{transparency}{olo}{ $_->[1] } = Arithmetic(
                name          => $name,
                defaultFormat => '0softnz',
                arithmetic =>
                  '=IF(IV123,0,IV1)+SUMPRODUCT(IV11_IV12,IV13_IV14,IV15_IV16)',
                arguments => {
                    IV123     => $model->{transparencyMasterFlag},
                    IV1       => $model->{transparency}{"ol$_->[1]"},
                    IV11_IV12 => $_->[0],
                    IV13_IV14 => $agreedCapacity,
                    IV15_IV16 => $model->{transparency},
                },
            );
          } (
            [ $assetsCapacity[$#assetsCapacity],       119303 ],
            [ $assetsConsumption[$#assetsConsumption], 119304 ],
          )
      )
      : (
        map {
            my $name = $_->objectShortName;
            $name =~ s/\(£\/kVA\)/(£)/;
            SumProduct(
                name          => $name,
                defaultFormat => '0softnz',
                matrix        => $_,
                vector        => $agreedCapacity
            );
          } (
            $assetsCapacity[$#assetsCapacity],
            $assetsConsumption[$#assetsConsumption]
          )
      );

    my $totalAssetsGenerationSoleUse =
      $model->{transparency}
      ? (
        $model->{transparency}{olo}{119302} = Arithmetic(
            name          => 'Total sole use assets for generation (£)',
            defaultFormat => '0softnz',
            arithmetic    => '=IF(IV123,0,IV1)+SUMPRODUCT(IV11_IV12,IV15_IV16)',
            arguments     => {
                IV123     => $model->{transparencyMasterFlag},
                IV1       => $model->{transparency}{ol119302},
                IV11_IV12 => $tariffSUexport,
                IV15_IV16 => $model->{transparency},
            },
        )
      )
      : GroupBy(
        source        => $tariffSUexport,
        name          => $tariffSUexport->objectShortName . ' (aggregate)',
        defaultFormat => $tariffSUexport->{defaultFormat}
      );

    push @{ $model->{calc1Tables} },
      my $totalAssets = Arithmetic(
        name          => 'All notional assets in EDCM (£)',
        arithmetic    => '=IV5+IV6+IV7+IV8',
        defaultFormat => '0softnz',
        arguments     => {
            IV5 => $totalAssetsFixed,
            IV6 => $totalAssetsCapacity,
            IV7 => $totalAssetsConsumption,
            IV8 => $totalAssetsGenerationSoleUse,
        }
      );
    $model->{transparency}{oli}{1229} = $totalAssets if $model->{transparency};

    my $assetsCapacityDoubleCooked =
      $assetsCapacityCooked[$#assetsCapacityCooked];

    $cdcmUse, $lossFactors, $diversity, $redUseRate, $capUseRate,
      $tariffSUimport, $assetsCapacity[$#assetsCapacity],
      $assetsConsumption[$#assetsConsumption], $totalAssetsFixed,
      $totalAssetsCapacity,
      $totalAssetsConsumption,
      $totalAssetsGenerationSoleUse, $totalAssets,
      $assetsCapacityCooked[$#assetsCapacityCooked],
      $assetsConsumptionCooked[$#assetsConsumptionCooked],
      $assetsCapacityDoubleCooked,
      $assetsConsumptionCooked[$#assetsConsumptionCooked],;

}

1;
