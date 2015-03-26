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
        name => 'Assets in CDCM model (£)'
          . ( $model->{transparency} ? '' : ' (from CDCM table 2706)' ),
        defaultFormat => '0hard',
        cols          => $assetLevelset,
        data          => [ '', map { 5e8 } 2 .. @{ $assetLevelset->{list} } ],
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
        groupName     => 'Assets in CDCM model',
        defaultFormat => '0softnz',
        matrix        => Constant(
            name  => 'EHV asset levels',
            cols  => $assetLevelset,
            byrow => 1,
            data  => [ [qw(1 1 1 1 1 1 0 0 0 0 0)] ]
        ),
        vector => $allAssets
    );
    $model->{transparency}{olFYI}{1233} = $ehvAssets if $model->{transparency};

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
    $model->{transparency}{olFYI}{1235} = $hvLvNetAssets
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
    $model->{transparency}{olFYI}{1231} = $hvLvServAssets
      if $model->{transparency};

    $allAssets, $ehvAssets, $hvLvNetAssets, $hvLvServAssets;

}

sub notionalAssets {

    my (
        $model,          $activeCoincidence,  $reactiveCoincidence,
        $agreedCapacity, $powerFactorInModel, $tariffCategory,
        $tariffSUimport, $tariffSUexport,     $cdcmAssets,
        $useProportions, $ehvAssetLevelset,   $cdcmUse,
    ) = @_;

    my $lossFactors = Dataset(
        name => 'Loss adjustment factor to transmission'
          . ' for each network level'
          . ( $model->{transparency} ? '' : ' (from CDCM table 2004)' ),
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

    my $diversity = Dataset(
        name => 'Diversity allowance between level exit '
          . 'and GSP Group'
          . ( $model->{transparency} ? '' : ' (from CDCM table 2611)' ),
        defaultFormat => '%hard',
        cols          => $ehvAssetLevelset,
        data          => [qw(0.1 0.1 0.3 0.3 0.3 0.7)],
        number        => 1105,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables}
    );

    my $useTextMatching =
      $model->{legacy201} || $model->{textCustomerCategories};

    my $customerCategory =
      $useTextMatching
      ? Arithmetic(
        name       => 'Tariff type and category',
        arithmetic => '="D"&TEXT(IV1,"0000")',
        arguments  => { IV1 => $tariffCategory }
      )
      : Arithmetic(
        name          => 'Index of customer category',
        defaultFormat => '0soft',
        arithmetic =>
'=1+(38*MOD(IV10,10)+(19*MOD(IV100,100)+(19*MOD(IV1000,1000)+IV1)/20)/10)/5',
        arguments => {
            IV1    => $tariffCategory,
            IV10   => $tariffCategory,
            IV100  => $tariffCategory,
            IV1000 => $tariffCategory,
        }
      );

    my $tariffCategoryset = Labelset(
        name => 'Customer categories',
        list => [ split /\n/, <<EOL] );
Category 0000
Category 1000
Category 0100
Category 1100
Category 0010
Not used
Category 0110
Category 1110
Category 0001
Category 1001
Category 0101
Category 1101
Category 0011
Not used
Category 0111
Category 1111
Category 0002
EOL

    my $lossFactorMap = Constant(
        name          => 'Mapping of customer category to loss factor',
        defaultFormat => '0con',
        rows          => $tariffCategoryset,
        data          => [
            [
                map { $_ eq '.' ? undef : $_ }
                  qw(1 2 3 3 4 . 4 4 6 6 5 5 5 . 5 5 5)
            ]
        ],
    );

    my $classificationMap = Constant(
        name => 'Treatment of network assets (1: capacity; 2+: consumption)',
        defaultFormat => '0con',
        rows          => $tariffCategoryset,
        cols          => $useProportions->{cols},
        data          => [
            [
                map { $_ eq '.' ? undef : $_ }
                  qw(0 1 0 2 0 . 0 3 0 3 0 3 0 . 0 4 0)
            ],
            [
                map { $_ eq '.' ? undef : $_ }
                  qw(0 0 1 1 0 . 2 2 0 0 2 2 0 . 3 3 0)
            ],
            [
                map { $_ eq '.' ? undef : $_ }
                  qw(0 0 0 0 1 . 1 1 0 0 0 0 2 . 2 2 0)
            ],
            [
                map { $_ eq '.' ? undef : $_ }
                  qw(0 0 0 0 0 . 0 0 0 0 1 1 1 . 1 1 1)
            ],
            [
                map { $_ eq '.' ? undef : $_ }
                  qw(0 0 0 0 0 . 0 0 1 1 0 0 0 . 0 0 0)
            ],
        ],
    );

    push @{ $model->{generalTables} },
      Columnset(
        name    => 'Rules applicable to customer categories',
        columns => [ $lossFactorMap, $classificationMap, ],
      ) if $model->{voltageRulesTransparency};

    my $tariffLossFactor = Arithmetic(
        name       => 'Loss factor to transmission',
        arithmetic => '=INDEX(IV8_IV9,'
          . (
            $useTextMatching
            ? 'IF(ISNUMBER(SEARCH("?0000",IV1)),1,'
              . 'IF(ISNUMBER(SEARCH("?1000",IV20)),2,'
              . 'IF(ISNUMBER(SEARCH("??100",IV21)),3,'
              . 'IF(ISNUMBER(SEARCH("???10",IV22)),4,'
              . 'IF(ISNUMBER(SEARCH("??001",IV23)),6,5)))))'
            : $model->{voltageRulesTransparency} ? 'INDEX(IV5_IV6,IV1)'
            :   'IF(IV1,IF(MOD(IV12,1000),IF(MOD(IV13,100),IF(MOD(IV14,10),IF(MOD(IV15,1000)=1,6,5),4),3),2),1)'
          )
          . ')',
        arguments => {
            $useTextMatching
            ? (
                IV1  => $customerCategory,
                IV20 => $customerCategory,
                IV21 => $customerCategory,
                IV22 => $customerCategory,
                IV23 => $customerCategory,
              )
            : $model->{voltageRulesTransparency} ? (
                IV1     => $customerCategory,
                IV5_IV6 => $lossFactorMap,
              )
            : (
                IV1  => $tariffCategory,
                IV12 => $tariffCategory,
                IV13 => $tariffCategory,
                IV14 => $tariffCategory,
                IV15 => $tariffCategory,
            ),
            IV8_IV9 => $lossFactors,
        },
    );

    my $redUseRate = Arithmetic(
        name => 'Peak-time active power consumption'
          . ' adjusted to transmission (kW/kVA)',
        groupName  => 'Active power consumption',
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
        newBlock   => 1,
        arguments  => {
            IV9 => $powerFactorInModel,
            IV1 => $tariffLossFactor,
        }
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

    push @{ $model->{calc1Tables} },
      my $accretion =
      $useTextMatching
      ? Arithmetic(
        name       => 'Notional asset rate (£/kW)',
        newBlock   => 1,
        arithmetic => '=IF(IV1,IV2/IV3/IV4,0)',
        arguments  => {
            IV1 => $cdcmUse,
            IV2 => $cdcmAssets,
            IV3 => $cdcmUse,
            IV4 => $lossFactors
        },
        location => 'Charging rates',
      )
      : Arithmetic(
        name       => 'Notional asset rate (£/kW)',
        newBlock   => 1,
        arithmetic => '=IV2/IV1/IV4',
        arguments  => {
            IV1 => $cdcmUse,
            IV2 => $cdcmAssets,
            IV4 => $lossFactors
        },
        location => 'Charging rates',
      );

    my $accretion132hvHard = Dataset(
        name => 'Override notional asset rate for 132kV/HV (£/kW)',
        $useTextMatching
        ? ()
        : (
            lines => [
                'This value only affects tariffs if there are'
                  . ' 132kV/HV non-sole-use assets in the EDCM model. '
                  . 'It will not be used if set to zero or blank.',
                'If the forecast system simultaneous maximum load (kW)'
                  . ' from CDCM users at the 132kV/HV network level is zero,'
                  . ' then a non-zero non-blank value must be entered here.',
                'An arbitrary non-zero non-blank value should be entered here'
                  . ' if there are no 132kV/HV assets in the EDCM or in the 500 MW model.',
            ]
        ),
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
        $useTextMatching
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
        ),
        location => 'Charging rates',
    );

    my (
        $assetsCapacity,       $assetsConsumption,
        $assetsCapacityCooked, $assetsConsumptionCooked
    );

    my $useProportionsCooked = sub {
        Arithmetic(
            name       => 'Network use factors (second set)',
            arithmetic => '=MAX(IV3+0,MIN(IV1+0,IV2+0))',
            arguments  => {
                IV1 => $useProportions,
                IV2 => $usePropCap,
                IV3 => $usePropCollar,
            }
        );
    };

    if ( $model->{voltageRulesTransparency} ) {

        $accretion = Stack(
            name      => 'Notional asset rate adjusted (£/kW)',
            groupName => 'Notional asset rate',
            cols      => $useProportions->{cols},
            sources   => [ $accretion132hvcombined, $accretion ],
            location  => 'Charging rates',
        );

        my $machine = sub {
            my ( $name1, $name2, $useProportions, $useRate, $diversity,
                @extras, )
              = @_;

            SumProduct(
                name      => $name1,
                groupName => $name2,
                matrix    => SpreadsheetModel::Custom->new(
                    name      => $name2,
                    groupName => $name2,
                    @extras,
                    custom => [
                            '=IF(INDEX(IV5:IV6,IV4)'
                          . ( $diversity ? '=1' : '>1' )
                          . ',IV1*IV8'
                          . ( $diversity ? '/(1+IV3)' : '' ) . ',0)'
                    ],
                    wsPrepare => sub {
                        my ( $self, $wb, $ws, $format, $formula, $pha, $rowh,
                            $colh )
                          = @_;
                        sub {
                            my ( $x, $y ) = @_;
                            '', $format, $formula->[0],
                              qr/\bIV1\b/ =>
                              Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                                $rowh->{IV1} + $y,
                                $colh->{IV1} + $x
                              ),
                              qr/\bIV4\b/ =>
                              Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                                $rowh->{IV4} + $y,
                                $colh->{IV4}, 0, 1 ),
                              $diversity
                              ? ( # NB: shifted by one to the right because of the GSP entry
                                qr/\bIV3\b/ =>
                                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                                    $rowh->{IV3}, $colh->{IV3} + 1 + $x, 1
                                  )
                              )
                              : (),
                              qr/\bIV8\b/ =>
                              Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                                $rowh->{IV8} + $y,
                                $colh->{IV8}, 0, 1 ),
                              qr/\bIV5\b/ =>
                              Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                                $rowh->{IV5_IV6}, $colh->{IV5_IV6} + $x, 1 ),
                              qr/\bIV6\b/ =>
                              Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                                $rowh->{IV5_IV6} + $classificationMap->lastRow,
                                $colh->{IV5_IV6} + $x,
                                1
                              );
                        };
                    },
                    rows      => $useProportions->{rows},
                    cols      => $useProportions->{cols},
                    arguments => {
                        IV1     => $useProportions,
                        IV5     => $classificationMap,
                        IV5_IV6 => $classificationMap,
                        IV4     => $customerCategory,
                        $diversity ? ( IV3 => $diversity ) : (),
                        IV8 => ref $useRate eq 'ARRAY'
                        ? $useRate->[1]
                        : $useRate,
                    },
                ),
                vector => $accretion,
            );
        };

        $assetsCapacity = $machine->(
            'Capacity assets (£/kVA)',
            'Adjusted network use by capacity',
            $useProportions, $capUseRate, $diversity,
        );
        $assetsConsumption = $machine->(
            'Consumption assets (£/kVA)',
            'Adjusted network use by consumption',
            $useProportions, $redUseRate, undef, newBlock => 1,
        );

        $useProportionsCooked = $useProportionsCooked->();
        $assetsCapacityCooked = $machine->(
            'Second set of capacity assets (£/kVA)',
            'Second set of adjusted network use by capacity',
            $useProportionsCooked,
            $capUseRate,
            $diversity,
        );
        $assetsConsumptionCooked = $machine->(
            'Second set of consumption assets (£/kVA)',
            'Second set of adjusted network use by consumption',
            $useProportionsCooked,
            $redUseRate,
            undef,
            newBlock => 1,
        );

    }

    else {

        $accretion = Stack(
            name      => 'Notional asset rate adjusted (£/kW)',
            groupName => 'Notional asset rate',
            cols      => $ehvAssetLevelset,
            sources   => [ $accretion132hvcombined, $accretion ],
            location  => 'Charging rates',
        );

        $useProportionsCooked = $useProportionsCooked->();

        my (
            @assetsCapacity,       @assetsConsumption,
            @assetsCapacityCooked, @assetsConsumptionCooked,
        );

        my $starIV5       = $useProportions       ? '*IV5' : '';
        my $starIV5Cooked = $useProportionsCooked ? '*IV5' : '';

        if ( !$useTextMatching ) {    # does not support allowInvalid

            push @assetsCapacity,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[1] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
                arithmetic => qq@=IF(IV1=1000,IV4*IV8/(1+IV3)$starIV5,0)@,
                arguments  => {
                    IV1 => $tariffCategory,
                    IV4 => $accretion,
                    IV3 => $diversity,
                    IV8 => $capUseRate,
                    $useProportions ? ( IV5 => $useProportions ) : (),
                },
              );

            push @assetsCapacityCooked,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[1] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
                arithmetic => qq@=IF(IV1=1000,IV4*IV8/(1+IV3)$starIV5Cooked,0)@,
                arguments  => {
                    IV1 => $tariffCategory,
                    IV4 => $accretion,
                    IV3 => $diversity,
                    IV8 => $capUseRate,
                    $useProportionsCooked
                    ? ( IV5 => $useProportionsCooked )
                    : (),
                },
              );

            push @assetsCapacity,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[2] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
                arithmetic =>
                  qq@=IF(MOD(IV1,1000)=100,IV4*IV8/(1+IV3)$starIV5,0)@,
                arguments => {
                    IV1 => $tariffCategory,
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
                  qq@=IF(MOD(IV1,1000)=100,IV4*IV8/(1+IV3)$starIV5Cooked,0)@,
                arguments => {
                    IV1 => $tariffCategory,
                    IV4 => $accretion,
                    IV3 => $diversity,
                    IV8 => $capUseRate,
                    $useProportionsCooked
                    ? ( IV5 => $useProportionsCooked )
                    : (),
                },
              );

            push @assetsCapacity,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[3] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
                arithmetic =>
                  qq@=IF(MOD(IV1,100)=10,IV4*IV8/(1+IV3)$starIV5,0)@,
                arguments => {
                    IV1 => $tariffCategory,
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
                  qq@=IF(MOD(IV1,100)=10,IV4*IV8/(1+IV3)$starIV5Cooked,0)@,
                arguments => {
                    IV1 => $tariffCategory,
                    IV4 => $accretion,
                    IV3 => $diversity,
                    IV8 => $capUseRate,
                    $useProportionsCooked
                    ? ( IV5 => $useProportionsCooked )
                    : (),
                },
              );

            push @assetsCapacity,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[4] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[4] ] ),
                arithmetic =>
qq@=IF(AND(MOD(IV1,10)>0,MOD(IV2,1000)>1),IV4*IV8/(1+IV3)$starIV5,0)@,
                arguments => {
                    IV1 => $tariffCategory,
                    IV2 => $tariffCategory,
                    IV4 => $accretion,
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
qq@=IF(AND(MOD(IV1,10)>0,MOD(IV2,1000)>1),IV4*IV8/(1+IV3)$starIV5Cooked,0)@,
                arguments => {
                    IV1 => $tariffCategory,
                    IV2 => $tariffCategory,
                    IV4 => $accretion,
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
                  qq@=IF(MOD(IV1,1000)=1,IV4*IV8/(1+IV3)$starIV5,0)@,
                arguments => {
                    IV1 => $tariffCategory,
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
                  qq@=IF(MOD(IV1,1000)=1,IV4*IV8/(1+IV3)$starIV5Cooked,0)@,
                arguments => {
                    IV1 => $tariffCategory,
                    IV4 => $accretion,
                    IV3 => $diversity,
                    IV8 => $capUseRate,
                    $useProportionsCooked
                    ? ( IV5 => $useProportionsCooked )
                    : (),
                },
              );

            push @assetsConsumption,
              Arithmetic(
                name => "Consumption $accretion->{cols}{list}[1] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
                arithmetic => qq@=IF(IV1>1000,IV4*IV9$starIV5,0)@,
                arguments  => {
                    IV1 => $tariffCategory,
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
                arithmetic => qq@=IF(IV1>1000,IV4*IV9$starIV5Cooked,0)@,
                arguments  => {
                    IV1 => $tariffCategory,
                    IV4 => $accretion,
                    IV9 => ref $redUseRate eq 'ARRAY' ? $redUseRate->[1]
                    : $redUseRate,
                    $useProportionsCooked ? ( IV5 => $useProportionsCooked )
                    : (),
                },
              );

            push @assetsConsumption,
              Arithmetic(
                name => "Consumption $accretion->{cols}{list}[2] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
                arithmetic => qq@=IF(MOD(IV1,1000)>100,IV4*IV9$starIV5,0)@,
                arguments  => {
                    IV1 => $tariffCategory,
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
                  qq@=IF(MOD(IV1,1000)>100,IV4*IV9$starIV5Cooked,0)@,
                arguments => {
                    IV1 => $tariffCategory,
                    IV4 => $accretion,
                    IV9 => ref $redUseRate eq 'ARRAY' ? $redUseRate->[1]
                    : $redUseRate,
                    $useProportionsCooked ? ( IV5 => $useProportionsCooked )
                    : (),
                },
              );

            push @assetsConsumption,
              Arithmetic(
                name => "Consumption $accretion->{cols}{list}[3] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
                arithmetic => qq@=IF(MOD(IV1,100)>10,IV4*IV9$starIV5,0)@,
                arguments  => {
                    IV1 => $tariffCategory,
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
                arithmetic => qq@=IF(MOD(IV1,100)>10,IV4*IV9$starIV5Cooked,0)@,
                arguments  => {
                    IV1 => $tariffCategory,
                    IV4 => $accretion,
                    IV9 => ref $redUseRate eq 'ARRAY' ? $redUseRate->[1]
                    : $redUseRate,
                    $useProportionsCooked ? ( IV5 => $useProportionsCooked )
                    : (),
                },
              );

        }

        elsif ( !$model->{allowInvalid} ) {    # legacy, without allowInvalid

            push @assetsCapacity,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[1] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
                arithmetic => qq@=IF(IV1="D1000",IV4*IV8/(1+IV3)$starIV5,0)@,
                arguments  => {
                    IV1 => $customerCategory,
                    IV4 => $accretion,
                    IV3 => $diversity,
                    IV8 => $capUseRate,
                    $useProportions ? ( IV5 => $useProportions ) : (),
                },
              );

            push @assetsCapacityCooked,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[1] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
                arithmetic =>
                  qq@=IF(IV1="D1000",IV4*IV8/(1+IV3)$starIV5Cooked,0)@,
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

            push @assetsCapacity,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[2] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
                arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?100",IV1)),IV4*IV8/(1+IV3)$starIV5,0)@,
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
                name => "Capacity $accretion->{cols}{list}[2] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[2] ] ),
                arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D?100",IV1)),IV4*IV8/(1+IV3)$starIV5Cooked,0)@,
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

            push @assetsCapacity,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[3] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
                arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D??10",IV1)),IV4*IV8/(1+IV3)$starIV5,0)@,
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
                name => "Capacity $accretion->{cols}{list}[3] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[3] ] ),
                arithmetic =>
qq@=IF(ISNUMBER(SEARCH("D??10",IV1)),IV4*IV8/(1+IV3)$starIV5Cooked,0)@,
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
qq@=IF(OR(IV6="D0002",ISNUMBER(SEARCH("D?1?1",IV1)),ISNUMBER(SEARCH("D??11",IV7))),IV4*IV8/(1+IV3)$starIV5Cooked,0)@,
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
qq@=IF(ISNUMBER(SEARCH("D?001",IV1)),IV4*IV8/(1+IV3)$starIV5Cooked,0)@,
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
qq@=IF(IV1="D1000",0,IF(ISNUMBER(SEARCH("D1???",IV2)),IV4*IV9$starIV5Cooked,0))@,
                arguments => {
                    IV1 => $customerCategory,
                    IV2 => $customerCategory,
                    IV4 => $accretion,
                    IV9 => ref $redUseRate eq 'ARRAY' ? $redUseRate->[1]
                    : $redUseRate,
                    $useProportionsCooked ? ( IV5 => $useProportionsCooked )
                    : (),
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
qq@=IF(ISNUMBER(SEARCH("D?100",IV1)),0,IF(ISNUMBER(SEARCH("D?1??",IV2)),IV4*IV9$starIV5Cooked,0))@,
                arguments => {
                    IV1 => $customerCategory,
                    IV2 => $customerCategory,
                    IV4 => $accretion,
                    IV9 => ref $redUseRate eq 'ARRAY' ? $redUseRate->[1]
                    : $redUseRate,
                    $useProportionsCooked ? ( IV5 => $useProportionsCooked )
                    : (),
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
qq@=IF(ISNUMBER(SEARCH("D??10",IV1)),0,IF(ISNUMBER(SEARCH("D??1?",IV2)),IV4*IV9$starIV5Cooked,0))@,
                arguments => {
                    IV1 => $customerCategory,
                    IV2 => $customerCategory,
                    IV4 => $accretion,
                    IV9 => ref $redUseRate eq 'ARRAY' ? $redUseRate->[1]
                    : $redUseRate,
                    $useProportionsCooked ? ( IV5 => $useProportionsCooked )
                    : (),
                },
              );

        }

        else {    # legacy, if allowInvalid

            push @assetsCapacity,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[1] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
                arithmetic => qq@=IF(IV1="D1000",IV4*IV8/(1+IV3)$starIV5,0)@,
                arguments  => {
                    IV1 => $customerCategory,
                    IV4 => $accretion,
                    IV3 => $diversity,
                    IV8 => $capUseRate,
                    $useProportions ? ( IV5 => $useProportions ) : (),
                },
              );

            push @assetsCapacityCooked,
              Arithmetic(
                name => "Capacity $accretion->{cols}{list}[1] (£/kVA)",
                cols => Labelset( list => [ $accretion->{cols}{list}[1] ] ),
                arithmetic =>
                  qq@=IF(IV1="D1000",IV4*IV8/(1+IV3)$starIV5Cooked,0)@,
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

            my $useProportionsCapped = Arithmetic(
                name       => 'Network use factors (capped only)',
                arithmetic => '=MIN(IV1+0,IV2+0)',
                arguments  => {
                    IV1 => $useProportions,
                    IV2 => $usePropCap,
                }
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
qq@=IF(ISNUMBER(SEARCH("D?100",IV1)),IV4*IV8/(1+IV3)$starIV5Cooked,0)@,
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
qq@=IF(ISNUMBER(SEARCH("D??10",IV1)),IV4*IV8/(1+IV3)$starIV5Cooked,0)@,
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
qq@=IF(OR(IV6="D0002",ISNUMBER(SEARCH("D?1?1",IV1)),ISNUMBER(SEARCH("D??11",IV7))),IV4*IV8/(1+IV3)$starIV5Cooked,0)@,
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
qq@=IF(ISNUMBER(SEARCH("D?001",IV1)),IV4*IV8/(1+IV3)$starIV5Cooked,0)@,
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

        $assetsCapacity = Arithmetic(
            name       => 'Total notional capacity assets (£/kVA)',
            groupName  => 'First set of notional assets',
            cols       => 0,
            arithmetic => '='
              . join( '+', map { "IV$_" } 1 .. @assetsCapacity ),
            arguments => {
                map { ( "IV$_" => $assetsCapacity[ $_ - 1 ] ) }
                  1 .. @assetsCapacity
            },
        );

        $assetsCapacityCooked = Arithmetic(
            name       => 'Second set of capacity assets (£/kVA)',
            groupName  => 'Second set of notional capacity assets',
            cols       => 0,
            arithmetic => '='
              . join( '+', map { "IV$_" } 1 .. @assetsCapacityCooked ),
            arguments => {
                map { ( "IV$_" => $assetsCapacityCooked[ $_ - 1 ] ) }
                  1 .. @assetsCapacityCooked
            },
        );

        $assetsConsumption = Arithmetic(
            name       => 'Total notional consumption assets (£/kVA)',
            groupName  => 'First set of notional assets',
            cols       => 0,
            arithmetic => '='
              . join( '+', map { "IV$_" } 1 .. @assetsConsumption ),
            arguments => {
                map { ( "IV$_" => $assetsConsumption[ $_ - 1 ] ) }
                  1 .. @assetsConsumption
            },
        );

        $assetsConsumptionCooked = Arithmetic(
            name       => 'Second set of consumption assets (£/kVA)',
            groupName  => 'Second set of notional consumption assets',
            cols       => 0,
            arithmetic => '='
              . join( '+', map { "IV$_" } 1 .. @assetsConsumptionCooked ),
            arguments => {
                map { ( "IV$_" => $assetsConsumptionCooked[ $_ - 1 ] ) }
                  1 .. @assetsConsumptionCooked
            },
        );

    }

    $model->{transparency}{olFYI}{1225} = $accretion
      if $model->{transparency};

    my $totalAssetsFixed =
      $model->{transparencyMasterFlag}
      ? Arithmetic(
        name          => 'Total sole use assets for demand (£)',
        newBlock      => 1,
        defaultFormat => '0softnz',
        arithmetic    => '=IF(IV123,0,IV1)+SUMPRODUCT(IV11_IV12,IV15_IV16)',
        arguments     => {
            IV123     => $model->{transparencyMasterFlag},
            IV1       => $model->{transparency}{ol119301},
            IV11_IV12 => $tariffSUimport,
            IV15_IV16 => $model->{transparency},
        },
      )
      : GroupBy(
        source        => $tariffSUimport,
        name          => 'Total sole use assets for demand (£)',
        newBlock      => 1,
        defaultFormat => '0softnz'
      );

    $model->{transparency}{olTabCol}{119301} = $totalAssetsFixed
      if $model->{transparency};

    my ( $totalAssetsCapacity, $totalAssetsConsumption ) =
      $model->{transparencyMasterFlag}
      ? (
        map {
            my $name = $_->[0]->objectShortName;
            $name =~ s/\(£\/kVA\)/(£)/;
            Arithmetic(
                name          => $name,
                groupName     => 'Notional assets in EDCM model',
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
        } ( [ $assetsCapacity, 119303 ], [ $assetsConsumption, 119304 ], )
      )
      : (
        map {
            my $name = $_->objectShortName;
            $name =~ s/\(£\/kVA\)/(£)/;
            SumProduct(
                name          => $name,
                groupName     => 'Notional assets in EDCM model',
                defaultFormat => '0softnz',
                matrix        => $_,
                vector        => $agreedCapacity
            );
        } ( $assetsCapacity, $assetsConsumption )
      );

    if ( $model->{transparency} ) {
        $model->{transparency}{olTabCol}{119303} = $totalAssetsCapacity;
        $model->{transparency}{olTabCol}{119304} = $totalAssetsConsumption;
    }

    my $totalAssetsGenerationSoleUse =
      $model->{transparencyMasterFlag}
      ? Arithmetic(
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
      : GroupBy(
        source        => $tariffSUexport,
        name          => $tariffSUexport->objectShortName . ' (aggregate)',
        defaultFormat => $tariffSUexport->{defaultFormat}
      );

    $model->{transparency}{olTabCol}{119302} = $totalAssetsGenerationSoleUse
      if $model->{transparency};

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
    $model->{transparency}{olFYI}{1229} = $totalAssets
      if $model->{transparency};

    $lossFactors, $diversity, $accretion, $redUseRate, $capUseRate,
      $tariffSUimport,    $assetsCapacity,
      $assetsConsumption, $totalAssetsFixed,
      $totalAssetsCapacity,
      $totalAssetsConsumption,
      $totalAssetsGenerationSoleUse, $totalAssets,
      $assetsCapacityCooked,
      $assetsConsumptionCooked;

}

1;
