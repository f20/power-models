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
use EDCM2::Ldno;
use EDCM2::Inputs;
use EDCM2::Assets;
use EDCM2::Locations;
use EDCM2::Charges;
use EDCM2::Generation;
use EDCM2::Scaling;
use EDCM2::Sheets;
use EDCM2::DataPreprocess;

sub requiredModulesForRuleset {
    my ( $class, $ruleset ) = @_;
    $ruleset->{transparency}
      && $ruleset->{transparency} =~ /impact/i ? qw(EDCM2::Impact)   : (),
      $ruleset->{customerTemplates}            ? qw(EDCM2::Template) : (),
      $ruleset->{layout} ? qw(SpreadsheetModel::MatrixSheet EDCM2::Layout) : (),
      $ruleset->{checksums} ? qw(SpreadsheetModel::Checksum) : ();
}

sub new {

    my $class = shift;
    my $model = bless {@_}, $class;

    die 'This system will not build an orange '
      . 'EDCM model without a suitable disclaimer.' . "\n--"
      if $model->{colour}
      && $model->{colour} =~ /orange/i
      && !($model->{extraNotice}
        && length( $model->{extraNotice} ) > 299
        && $model->{extraNotice} =~ /DCUSA/ );

    $model->{inputTables} = [];
    $model->{method} ||= 'none';

    # The EDCM timeband is called purple in this code,
    # but its display name defaults to super-red.
    $model->{TimebandName} ||=
      ucfirst( $model->{timebandName} ||= 'super-red' );

    $model->preprocessDataset
      if $model->{dataset} && keys %{ $model->{dataset} };

    $model->{numLocations} ||= $model->{numLocationsDefault};
    $model->{numTariffs}   ||= $model->{numTariffsDefault};

    my $ehvAssetLevelset = Labelset(
        name => 'EHV asset levels',
        list => [ split /\n/, <<EOT ] );
GSP
132kV circuits
132kV/EHV
EHV circuits
EHV/HV
132kV/HV
EOT

    my (
        $tariffs,                          $importCapacity,
        $exportCapacityExempt,             $exportCapacityChargeablePre2005,
        $exportCapacityChargeable20052010, $exportCapacityChargeablePost2010,
        $tariffSoleUseMeav,                $dcp189Input,
        $tariffLoc,                        $tariffCategory,
        $useProportions,                   $activeCoincidence,
        $reactiveCoincidence,              $indirectExposure,
        $nonChargeableCapacity,            $activeUnits,
        $creditableCapacity,               $tariffNetworkSupportFactor,
        $tariffDaysInYearNot,              $tariffHoursInPurpleNot,
        $previousChargeImport,             $previousChargeExport,
        $llfcImport,                       $llfcExport,
        $actualRedDemandRate,
    );
    if ( $model->{dcp189} ) {
        (
            $tariffs,
            $importCapacity,
            $exportCapacityExempt,
            $exportCapacityChargeablePre2005,
            $exportCapacityChargeable20052010,
            $exportCapacityChargeablePost2010,
            $tariffSoleUseMeav,
            $dcp189Input,
            $tariffLoc,
            $tariffCategory,
            $useProportions,
            $activeCoincidence,
            $reactiveCoincidence,
            $indirectExposure,
            $nonChargeableCapacity,
            $activeUnits,
            $creditableCapacity,
            $tariffNetworkSupportFactor,
            $tariffDaysInYearNot,
            $tariffHoursInPurpleNot,
            $previousChargeImport,
            $previousChargeExport,
            $llfcImport,
            $llfcExport,
            $actualRedDemandRate,
        ) = $model->tariffInputs($ehvAssetLevelset);
    }
    else {
        (
            $tariffs,
            $importCapacity,
            $exportCapacityExempt,
            $exportCapacityChargeablePre2005,
            $exportCapacityChargeable20052010,
            $exportCapacityChargeablePost2010,
            $tariffSoleUseMeav,
            $tariffLoc,
            $tariffCategory,
            $useProportions,
            $activeCoincidence,
            $reactiveCoincidence,
            $indirectExposure,
            $nonChargeableCapacity,
            $activeUnits,
            $creditableCapacity,
            $tariffNetworkSupportFactor,
            $tariffDaysInYearNot,
            $tariffHoursInPurpleNot,
            $previousChargeImport,
            $previousChargeExport,
            $llfcImport,
            $llfcExport,
            $actualRedDemandRate,
        ) = $model->tariffInputs($ehvAssetLevelset);
    }

    my ( $locations, $locParent, $c1, $a1d, $r1d, $a1g, $r1g ) =
      $model->loadFlowInputs;

    $model->{transparencyMasterFlag} = Dataset(
        name => 'Is this the master model containing all the tariff data?',
        defaultFormat => 'boolhard',
        singleRowName => 'Enter TRUE or FALSE',
        validation    => {
            validate => 'list',
            value    => [ 'TRUE', 'FALSE' ],
        },
        data     => ['TRUE'],
        dataset  => $model->{dataset},
        appendTo => $model->{inputTables},
        number   => 1190,
      )
      if $model->{transparency}
      && $model->{transparency} !~ /outputonly/i
      && !$model->{legacy201};

    if ( $model->{transparency} ) {

        if ( $model->{transparency} =~ /impact/i ) {

            $model->impactNotes;

            ( $locations, $locParent, $c1, $a1d, $r1d, $a1g, $r1g ) =
              $model->mangleLoadFlowInputs( $locations, $locParent, $c1, $a1d,
                $r1d, $a1g, $r1g );

            (
                $tariffs,
                $importCapacity,
                $exportCapacityExempt,
                $exportCapacityChargeablePre2005,
                $exportCapacityChargeable20052010,
                $exportCapacityChargeablePost2010,
                $tariffSoleUseMeav,
                $tariffLoc,
                $tariffCategory,
                $useProportions,
                $activeCoincidence,
                $reactiveCoincidence,
                $indirectExposure,
                $nonChargeableCapacity,
                $activeUnits,
                $creditableCapacity,
                $tariffNetworkSupportFactor,
                $tariffDaysInYearNot,
                $tariffHoursInPurpleNot,
                $previousChargeImport,
                $previousChargeExport,
                $llfcImport,
                $llfcExport,
                $actualRedDemandRate,
                $model->{transparency},
              )
              = $model->mangleTariffInputs(
                $tariffs,
                $importCapacity,
                $exportCapacityExempt,
                $exportCapacityChargeablePre2005,
                $exportCapacityChargeable20052010,
                $exportCapacityChargeablePost2010,
                $tariffSoleUseMeav,
                $tariffLoc,
                $tariffCategory,
                $useProportions,
                $activeCoincidence,
                $reactiveCoincidence,
                $indirectExposure,
                $nonChargeableCapacity,
                $activeUnits,
                $creditableCapacity,
                $tariffNetworkSupportFactor,
                $tariffDaysInYearNot,
                $tariffHoursInPurpleNot,
                $previousChargeImport,
                $previousChargeExport,
                $llfcImport,
                $llfcExport,
              );

            $model->{transparencyImpact} = 1;

        }
        elsif ( $model->{transparency} =~ /outputonly/i ) {
            $model->{transparency} = {};
        }
        else {
            delete $model->{transparency};
        }

    }

    $model->{transparencyMasterFlag} = Arithmetic(
        name          => 'Is this the master model?',
        defaultFormat => 'boolsoft',
        arithmetic    => '=IF(ISERROR(A5),TRUE,'
          . 'IF(A4="FALSE",FALSE,IF(A3=FALSE,FALSE,TRUE)))',
        arguments => {
            A3 => $model->{transparencyMasterFlag},
            A4 => $model->{transparencyMasterFlag},
            A5 => $model->{transparencyMasterFlag},
        },
    ) if $model->{transparencyMasterFlag};

    $model->{transparency} = Arithmetic(
        name  => 'Weighting of each tariff for reconciliation of totals',
        lines => [
'0 means that the tariff is active and is included in the table 119x aggregates.',
'-1 means that the tariff is included in the table 119x aggregates but should be removed.',
'1 means that the tariff is active and is not included in the table 119x aggregates.',
        ],
        arithmetic => '=IF(OR(A3,'
          . 'NOT(ISERROR(SEARCH("[ADDED]",A2))))' . ',1,'
          . 'IF(ISERROR(SEARCH("[REMOVED]",A1)),0,-1)' . ')',
        arguments => {
            A1 => $tariffs,
            A2 => $tariffs,
            A3 => $model->{transparencyMasterFlag},
        },
    ) if $model->{transparencyMasterFlag} && !defined $model->{transparency};

    if ( $model->{transparencyMasterFlag} ) {

        foreach my $set (
            [
                1191,
                [ 'Baseline total EDCM peak time consumption (kW)', '0hard' ],
                [
                    'Baseline total marginal effect'
                      . ' of indirect cost adder (kVA)',
                    '0hard'
                ],
                [
                    'Baseline total marginal effect of demand adder (kVA)',
                    '0hard'
                ],
                [ 'Baseline revenue from demand charge 1 (£/year)', '0hard' ],
                'EDCM demand aggregates'
            ],
            [
                1192,
                [ 'Baseline total chargeable export capacity (kVA)', '0hard' ],
                [
                    'Baseline total non-exempt 2005-2010 export capacity (kVA)',
                    '0hard'
                ],
                [
                    'Baseline total non-exempt post-2010 export capacity (kVA)',
                    '0hard'
                ],
                [
                    'Baseline net forecast EDCM generation revenue (£/year)',
                    '0hard'
                ],
                'EDCM generation aggregates'
            ],
            [
                1193,
                [ 'Baseline total sole use assets for demand (£)', '0hard' ],
                [
                    'Baseline total sole use assets for generation (£)',
                    '0hard'
                ],
                [ 'Baseline total notional capacity assets (£)',    '0hard' ],
                [ 'Baseline total notional consumption assets (£)', '0hard' ],
                [
                    'Baseline total non sole use'
                      . ' notional assets subject to matching (£)',
                    '0hard'
                ],
                $model->{dcp189} && $model->{dcp189} =~ /preservePot|split/i
                ? [
                    'Baseline total demand sole use assets '
                      . 'qualifying for DCP 189 discount (£)',
                    '0hard'
                  ]
                : (),
                'EDCM notional asset aggregates'
            ],
          )
        {
            my @cols;
            foreach my $col ( 1 .. ( $#$set - 1 ) ) {
                push @cols,
                  $model->{transparency}{"ol$set->[0]0$col"} = Dataset(
                    name          => $set->[$col][0],
                    defaultFormat => $set->[$col][1],
                    data          => [ [''] ],
                  );
            }
            Columnset(
                name          => "Baseline $set->[$#$set]",
                singleRowName => $set->[$#$set],
                number        => $set->[0],
                dataset       => $model->{dataset},
                appendTo      => $model->{inputTables},
                columns       => \@cols,
            );
        }

    }

    my ( $cdcmAssets, $cdcmEhvAssets, $cdcmHvLvShared, $cdcmHvLvService, ) =
      $model->cdcmAssets;

    if ( $model->{ldnoRev} && $model->{ldnoRev} =~ /only/i ) {
        $model->{daysInYear} = Dataset(
            name          => 'Days in year',
            defaultFormat => '0hard',
            data          => [365],
            dataset       => $model->{dataset},
            appendTo      => $model->{inputTables},
            number        => 1111,
            validation    => {
                validate => 'decimal',
                criteria => 'between',
                minimum  => 365,
                maximum  => 366,
            },
        );
        $model->ldnoRev;
        return $model;
    }

    $model->{matricesData} = [ [], [] ]
      if $model->{summaries} && $model->{summaries} =~ /matri/i;

    my (
        $daysInYear,            $chargeDirect,
        $chargeIndirect,        $chargeRates,
        $chargeExit,            $ehvIntensity,
        $allowedRevenue,        $powerFactorInModel,
        $genPot20p,             $genPotGP,
        $genPotGL,              $genPotCdcmCap20052010,
        $genPotCdcmCapPost2010, $hoursInPurple,
    ) = $model->generalInputs;

    $model->ldnoRev if $model->{ldnoRev};

    my $exportEligible = Arithmetic(
        name          => 'Has export charges?',
        defaultFormat => 'boolsoft',
        arithmetic    => '=OR(A1<>"VOID",A2<>"VOID",A3<>"VOID")',
        arguments     => {
            A1 => $exportCapacityChargeablePre2005,
            A2 => $exportCapacityChargeable20052010,
            A3 => $exportCapacityChargeablePost2010,
        }
    );

    my $importEligible = Arithmetic(
        name          => 'Has import charges?',
        defaultFormat => 'boolsoft',
        arithmetic    => '=A1<>"VOID"',
        arguments     => {
            A1 => $importCapacity,
        }
    );

### Marker

    my $importCapacityUnscaled = $importCapacity;
    my $chargeableCapacity     = Arithmetic(
        name          => 'Import capacity not subject to DSM (kVA)',
        defaultFormat => '0soft',
        arguments  => { A1 => $importCapacity, A2 => $nonChargeableCapacity, },
        arithmetic => '=A1-A2',
    );
    my $chargeableCapacity935  = $chargeableCapacity;
    my $activeCoincidence935   = $activeCoincidence;
    my $reactiveCoincidence935 = $reactiveCoincidence;

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

    my $exportCapacityChargeableUnscaled = Arithmetic(
        name          => 'Chargeable export capacity (kVA)',
        defaultFormat => '0soft',
        arithmetic    => '=A1+A4+A5',
        arguments     => {
            A1 => $exportCapacityChargeablePre2005,
            A4 => $exportCapacityChargeable20052010,
            A5 => $exportCapacityChargeablePost2010,
        }
    );

    my $creditableCapacityUnscaled = $creditableCapacity;

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

### Marker

    my $cdcmUse = $model->{cdcmComboTable} ? Stack(
        name => 'Forecast system simultaneous maximum load (kW)'
          . ' from CDCM users',
        defaultFormat => '0copy',
        cols          => $ehvAssetLevelset,
        rows          => 0,
        rowName       => 'System simultaneous maximum load (kW)',
        sources       => [ $model->{cdcmComboTable} ],
      ) : Dataset(
        name => 'Forecast system simultaneous maximum load (kW)'
          . ' from CDCM users'
          . ( $model->{transparency} ? '' : ' (from CDCM table 2506)' ),
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

    my (
        $lossFactors,            $diversity,
        $accretion,              $purpleUseRate,
        $capUseRate,             $assetsFixed,
        $assetsCapacity,         $assetsConsumption,
        $totalAssetsFixed,       $totalAssetsCapacity,
        $totalAssetsConsumption, $totalAssetsGenerationSoleUse,
        $totalEdcmAssets,        $assetsCapacityCooked,
        $assetsConsumptionCooked,
      )
      = $model->notionalAssets(
        $activeCoincidence,      $reactiveCoincidence,
        $importCapacity,         $powerFactorInModel,
        $tariffCategory,         $demandSoleUseAsset,
        $generationSoleUseAsset, $cdcmAssets,
        $useProportions,         $ehvAssetLevelset,
        $cdcmUse,
      );

### Marker

    my $rateDirect = Arithmetic(
        name          => 'Direct cost charging rate',
        groupName     => 'Expenditure charging rates',
        arithmetic    => '=A1/(A2+A3+(A4+A5)/A6)',
        defaultFormat => '%soft',
        arguments     => {
            A1 => $chargeDirect,
            A2 => $totalEdcmAssets,
            A3 => $cdcmEhvAssets,
            A4 => $cdcmHvLvShared,
            A5 => $cdcmHvLvService,
            A6 => $ehvIntensity,
        },
        location => 'Charging rates',
    );
    $model->{transparency}{olFYI}{1245} = $rateDirect if $model->{transparency};

    my $rateRates = Arithmetic(
        name          => 'Network rates charging rate',
        groupName     => 'Expenditure charging rates',
        arithmetic    => '=A1/(A2+A3+A4+A5)',
        defaultFormat => '%soft',
        arguments     => {
            A1 => $chargeRates,
            A2 => $totalEdcmAssets,
            A3 => $cdcmEhvAssets,
            A4 => $cdcmHvLvShared,
            A5 => $cdcmHvLvService,
        },
        location => 'Charging rates',
    );
    $model->{transparency}{olFYI}{1246} = $rateRates if $model->{transparency};

    my $rateIndirect = Arithmetic(
        name          => 'Indirect cost charging rate',
        groupName     => 'Expenditure charging rates',
        arithmetic    => '=A1/(A20+A3+(A4+A5)/A6)',
        defaultFormat => '%soft',
        arguments     => {
            A1  => $chargeIndirect,
            A21 => $totalAssetsCapacity,
            A22 => $totalAssetsConsumption,
            A3  => $cdcmEhvAssets,
            A4  => $cdcmHvLvShared,
            A6  => $ehvIntensity,
            A20 => $totalEdcmAssets,
            A5  => $cdcmHvLvService,
        },
        location => 'Charging rates',
    );
    $model->{transparency}{olFYI}{1250} = $rateIndirect
      if $model->{transparency};

    my $edcmIndirect = Arithmetic(
        name          => 'Indirect costs on EDCM demand (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=A1*(A20-A23)',
        arguments     => {
            A1  => $rateIndirect,
            A21 => $totalAssetsCapacity,
            A22 => $totalAssetsConsumption,
            A20 => $totalEdcmAssets,
            A23 => $totalAssetsGenerationSoleUse,
        },
    );
    $model->{transparency}{olFYI}{1253} = $edcmIndirect
      if $model->{transparency};

    my $edcmDirect = Arithmetic(
        name => 'Direct costs on EDCM demand except'
          . ' through sole use asset charges (£/year)',
        groupName     => 'Expenditure allocated to EDCM demand',
        defaultFormat => '0softnz',
        arithmetic    => '=A1*(A20+A23)',
        arguments     => {
            A1  => $rateDirect,
            A20 => $totalAssetsCapacity,
            A23 => $totalAssetsConsumption,
        },
    );
    $model->{transparency}{olFYI}{1252} = $edcmDirect if $model->{transparency};

    my $edcmRates = Arithmetic(
        name => 'Network rates on EDCM demand except '
          . 'through sole use asset charges (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=A1*(A20+A23)',
        arguments     => {
            A1  => $rateRates,
            A20 => $totalAssetsCapacity,
            A23 => $totalAssetsConsumption,
        },
    );
    $model->{transparency}{olFYI}{1255} = $edcmRates if $model->{transparency};

    my $fixedDcharge =
      !$model->{dcp189} ? Arithmetic(
        name          => 'Demand fixed charge p/day (scaled for part year)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=100/A2*A1*(A6+A88)',
        arguments     => {
            A1  => $demandSoleUseAsset,
            A6  => $rateDirect,
            A88 => $rateRates,
            A2  => $daysInYear,
        }
      )
      : $model->{dcp189} =~ /proportion/i ? Arithmetic(
        name          => 'Demand fixed charge p/day (scaled for part year)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=100/A2*A1*((1-A4)*A6+A88)',
        arguments     => {
            A1  => $demandSoleUseAsset,
            A4  => $dcp189Input,
            A6  => $rateDirect,
            A88 => $rateRates,
            A2  => $daysInYear,
        }
      )
      : Arithmetic(
        name          => 'Demand fixed charge p/day (scaled for part year)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=100/A2*A1*(IF(A4="Y",0,A6)+A88)',
        arguments     => {
            A1  => $demandSoleUseAsset,
            A4  => $dcp189Input,
            A6  => $rateDirect,
            A88 => $rateRates,
            A2  => $daysInYear,
        }
      );

    my $fixedDchargeTrue =
      !$model->{dcp189} ? Arithmetic(
        name          => 'Demand fixed charge p/day',
        groupName     => 'Fixed charges',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(A3,(100/A2*A1*(A6+A88)),0)',
        arguments     => {
            A1  => $demandSoleUseAssetUnscaled,
            A6  => $rateDirect,
            A88 => $rateRates,
            A2  => $daysInYear,
            A3  => $importEligible,
        }
      )
      : $model->{dcp189} =~ /proportion/i ? Arithmetic(
        name          => 'Demand fixed charge p/day',
        groupName     => 'Fixed charges',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(A3,(100/A2*A1*((1-A4)*A6+A88)),0)',
        arguments     => {
            A1  => $demandSoleUseAssetUnscaled,
            A4  => $dcp189Input,
            A6  => $rateDirect,
            A88 => $rateRates,
            A2  => $daysInYear,
            A3  => $importEligible,
        }
      )
      : Arithmetic(
        name          => 'Demand fixed charge p/day',
        groupName     => 'Fixed charges',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(A3,(100/A2*A1*(IF(A4="Y",0,A6)+A88)),0)',
        arguments     => {
            A1  => $demandSoleUseAssetUnscaled,
            A4  => $dcp189Input,
            A6  => $rateDirect,
            A88 => $rateRates,
            A2  => $daysInYear,
            A3  => $importEligible,
        }
      );

    my $fixedGcharge = Arithmetic(
        name          => 'Generation fixed charge p/day (scaled for part year)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=100/A2*A1*(A6+A88)',
        arguments     => {
            A1  => $generationSoleUseAsset,
            A6  => $rateDirect,
            A88 => $rateRates,
            A2  => $daysInYear,
        }
    );

    my $fixedGchargeUnround = Arithmetic(
        name          => 'Export fixed charge (unrounded) p/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=100/A2*A1*(A6+A88)',
        arguments     => {
            A1  => $generationSoleUseAssetUnscaled,
            A6  => $rateDirect,
            A88 => $rateRates,
            A2  => $daysInYear,
        }
    );

    my $fixedGchargeTrue = Arithmetic(
        name          => 'Export fixed charge p/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=ROUND(A1,2)',
        arguments     => { A1 => $fixedGchargeUnround, }
    );

### Marker

    my $totalDcp189DiscountedAssets;

    $totalDcp189DiscountedAssets =
      $model->{transparencyMasterFlag}
      ? (
        $model->{dcp189} =~ /proportion/i
        ? Arithmetic(
            name => 'Total demand sole use assets '
              . 'qualifying for DCP 189 discount (£)',
            defaultFormat => '0softnz',
            arithmetic => '=IF(A123,0,A1)+SUMPRODUCT(A11_A12,A13_A14,A15_A16)',
            arguments  => {
                A123    => $model->{transparencyMasterFlag},
                A1      => $model->{transparency}{ol119306},
                A11_A12 => $demandSoleUseAsset,
                A13_A14 => $dcp189Input,
                A15_A16 => $model->{transparency},
            },
          )
        : Arithmetic(
            name => 'Total demand sole use assets '
              . 'qualifying for DCP 189 discount (£)',
            defaultFormat => '0softnz',
            arithmetic    => '=IF(A123,0,A1)+SUMPRODUCT(A11_A12,A15_A16)',
            arguments     => {
                A123    => $model->{transparencyMasterFlag},
                A1      => $model->{transparency}{ol119306},
                A11_A12 => Arithmetic(
                    name => 'Demand sole use assets '
                      . 'qualifying for DCP 189 discount (£)',
                    defaultFormat => '0softnz',
                    arithmetic    => '=IF(A4="Y",A1,0)',
                    arguments     => {
                        A1 => $demandSoleUseAsset,
                        A4 => $dcp189Input,
                    }
                ),
                A15_A16 => $model->{transparency},
            },
        )
      )
      : $model->{dcp189} =~ /proportion/i ? SumProduct(
        name => 'Total demand sole use assets '
          . 'qualifying for DCP 189 discount (£)',
        defaultFormat => '0softnz',
        matrix        => $dcp189Input,
        vector        => $demandSoleUseAsset,
      )
      : Arithmetic(
        name => 'Total demand sole use assets '
          . 'qualifying for DCP 189 discount (£)',
        defaultFormat => '0softnz',
        arithmetic    => '=SUMIF(A1_A2,"Y",A3_A4)',
        arguments     => {
            A1_A2 => $dcp189Input,
            A3_A4 => $demandSoleUseAsset,
        }
      ) if $model->{dcp189} && $model->{dcp189} =~ /preservePot|split/i;

    $model->{transparency}{olTabCol}{119306} = $totalDcp189DiscountedAssets
      if $model->{transparency} && $totalDcp189DiscountedAssets;

### Marker

    my $cdcmPurpleUse = Stack(
        cols => Labelset( list => [ $cdcmUse->{cols}{list}[0] ] ),
        name    => 'Total CDCM peak time consumption (kW)',
        sources => [$cdcmUse]
    );
    $model->{transparency}{olFYI}{1237} = $cdcmPurpleUse
      if $model->{transparency};

    push @{ $model->{calc3Tables} }, $cdcmHvLvService, $cdcmEhvAssets,
      $cdcmHvLvShared
      if $model->{legacy201};

    my $edcmPurpleUse =
      $model->{transparencyMasterFlag}
      ? Arithmetic(
        name          => 'Total EDCM peak time consumption (kW)',
        defaultFormat => '0softnz',
        arithmetic    => '=IF(A123,0,A1)+SUMPRODUCT(A21_A22,A51_A52,A53_A54)',
        arguments     => {
            A123    => $model->{transparencyMasterFlag},
            A1      => $model->{transparency}{ol119101},
            A21_A22 => $model->{transparency},
            A51_A52 => ref $purpleUseRate eq 'ARRAY'
            ? $purpleUseRate->[0]
            : $purpleUseRate,
            A53_A54 => $importCapacity,
        }
      )
      : SumProduct(
        name   => 'Total EDCM peak time consumption (kW)',
        vector => ref $purpleUseRate eq 'ARRAY'
        ? $purpleUseRate->[0]
        : $purpleUseRate,
        matrix        => $importCapacity,
        defaultFormat => '0softnz'
      );

    $model->{transparency}{olTabCol}{119101} = $edcmPurpleUse
      if $model->{transparency};

    my $overallPurpleUse = Arithmetic(
        name          => 'Estimated total peak-time consumption (kW)',
        defaultFormat => '0softnz',
        arithmetic    => '=A1+A2',
        arguments     => { A1 => $cdcmPurpleUse, A2 => $edcmPurpleUse }
    );
    $model->{transparency}{olFYI}{1238} = $overallPurpleUse
      if $model->{transparency};

    my $rateExit = Arithmetic(
        name       => 'Transmission exit charging rate (£/kW/year)',
        arithmetic => '=A1/A2',
        arguments  => { A1 => $chargeExit, A2 => $overallPurpleUse },
        location   => 'Charging rates',
    );
    $model->{transparency}{olFYI}{1239} = $rateExit if $model->{transparency};

### Marker

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

    $reactiveCoincidence935 = Arithmetic(
        name       => 'Unadjusted but capped red kVAr/agreed kVA',
        arithmetic => '=MAX(MIN(SQRT(1-MIN(1,A2)^2),'
          . ( $model->{legacy201} ? '' : '0+' )
          . 'A1),0-SQRT(1-MIN(1,A3)^2))',
        arguments => {
            A1 => $reactiveCoincidence935,
            A2 => $activeCoincidence935,
            A3 => $activeCoincidence935,
        }
    );

    my ( $charges1, $acCoef, $reCoef ) =
      $model->charge1( $tariffLoc, $locations, $locParent, $c1, $a1d, $r1d,
        $a1g, $r1g,
        $model->preprocessLocationData( $locations, $a1d, $r1d, $a1g, $r1g, ),
      ) if $locations;

    my ( $fcpLricDemandCapacityChargeBig,
        $genCredit, $unitRateFcpLricNonDSM, $genCreditCapacity,
        $demandConsumptionFcpLric, )
      = $model->chargesFcpLric(
        $acCoef,                     $activeCoincidence,
        $charges1,                   $daysInYear,
        $reactiveCoincidence,        $reCoef,
        $tariffNetworkSupportFactor, $hoursInPurple,
        $hoursInPurple,              $importCapacity,
        $exportCapacityChargeable,   $creditableCapacity,
        $rateExit,                   $activeCoincidence935,
        $reactiveCoincidence935,
      );

    $genCredit = Arithmetic(
        name       => 'Generation credit (unrounded) p/kWh',
        groupName  => 'Generation unit rate credit',
        arithmetic => '=IF(A41,(A2*A3/(A4+A5)),0)',
        arguments  => {
            A1  => $exportCapacityChargeable,
            A2  => $genCredit,
            A21 => $genCredit,
            A3  => $exportCapacityChargeable,
            A4  => $exportCapacityChargeable,
            A5  => $exportCapacityExempt,
            A41 => $exportCapacityChargeable,
            A51 => $exportCapacityExempt,
        },
    );

### Marker

    my $gCharge = $model->gCharge(
        $genPot20p,                        $genPotGP,
        $genPotGL,                         $genPotCdcmCap20052010,
        $genPotCdcmCapPost2010,            $exportCapacityChargeable,
        $exportCapacityChargeable20052010, $exportCapacityChargeablePost2010,
        $daysInYear,
    );

    my $exportCapacityCharge = Arithmetic(
        name          => 'Export capacity charge (unrounded) p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(A2,A4,0)',
        arguments     => {
            A1 => $exportCapacityChargeable,
            A2 => $exportEligible,
            A4 => $gCharge,
        },
    );

    my $genCreditRound = Arithmetic(
        name          => "Export $model->{timebandName} unit rate (p/kWh)",
        defaultFormat => '0.000softnz',
        arithmetic    => '=ROUND(A1,3)',
        arguments     => { A1 => $genCredit }
    );

    my $genCreditCapacityRound = Arithmetic(
        name          => 'Generation credit (p/kVA/day)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(A2,ROUND(A1,2),0)',
        arguments     => {
            A1 => $genCreditCapacity,
            A2 => $exportEligible,
        }
    );

    my $exportCapacityChargeRound = Arithmetic(
        name          => 'Export capacity charge (p/kVA/day)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=ROUND(A1,2)',
        arguments     => { A1 => $exportCapacityCharge }
    );

    my $netexportCapacityChargeUnRound = Arithmetic(
        name =>
          'Net export capacity charge (or credit) (unrounded) (p/kVA/day)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(A21,(A1+A2),0)',
        arguments     => {
            A1  => $exportCapacityCharge,
            A2  => $genCreditCapacity,
            A21 => $exportEligible
        }
    );

    my $netexportCapacityChargeRound = Arithmetic(
        name          => 'Export capacity rate (p/kVA/day)',
        groupName     => 'Export capacity rate',
        defaultFormat => '0.00softnz',
        arithmetic    => '=ROUND(A1,2)',
        arguments     => { A1 => $netexportCapacityChargeUnRound, }
    );

    my $generationRevenue =
      $model->{transparencyMasterFlag}
      ? Arithmetic(
        name          => 'Net forecast EDCM generation revenue (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=IF(A123,0,A1)'
          . '+SUMPRODUCT(A21_A22,A51_A52,A53_A54)/100'
          . '+SUMPRODUCT(A31_A32,A71_A72,A73_A74)*A75/100+SUMPRODUCT(A41_A42,A83_A84)*A85/100',
        arguments => {
            A123    => $model->{transparencyMasterFlag},
            A1      => $model->{transparency}{ol119204},
            A21_A22 => $model->{transparency},
            A31_A32 => $model->{transparency},
            A41_A42 => $model->{transparency},
            A51_A52 => $genCreditRound,
            A53_A54 => $activeUnits,
            A71_A72 => $netexportCapacityChargeRound,
            A73_A74 => $exportCapacityChargeable,
            A75     => $daysInYear,
            A83_A84 => $fixedGcharge,
            A85     => $daysInYear,
        }
      )
      : Arithmetic(
        name          => 'Net forecast EDCM generation revenue (£/year)',
        defaultFormat => '0softnz',
        arithmetic =>
'=SUMPRODUCT(A51_A52,A53_A54)/100+SUMPRODUCT(A71_A72,A73_A74)*A75/100+SUM(A83_A84)*A85/100',
        arguments => {
            A51_A52 => $genCreditRound,
            A53_A54 => $activeUnits,
            A71_A72 => $netexportCapacityChargeRound,
            A73_A74 => $exportCapacityChargeable,
            A75     => $daysInYear,
            A83_A84 => $fixedGcharge,
            A85     => $daysInYear,
        }
      );

    $model->{transparency}{olTabCol}{119204} = $generationRevenue
      if $model->{transparency};

### Marker

    my $chargeOther = Arithmetic(
        name => 'Revenue less costs and '
          . (
            !$totalDcp189DiscountedAssets
              || $model->{dcp189} =~ /preservePot/i
            ? 'net forecast EDCM generation revenue'
            : 'adjustments'
          )
          . ' (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=A1-A2-A3-A4-A5'
          . ( $model->{table1101} ? '-A9' : '' )
          . (
            !$totalDcp189DiscountedAssets
              || $model->{dcp189} =~ /preservePot/i ? ''
            : '+A31*A32'
          ),
        arguments => {
            A1 => $allowedRevenue,
            $model->{table1101} ? ( A9 => $chargeExit ) : (),
            A2 => $chargeDirect,
            A3 => $chargeIndirect,
            A4 => $chargeRates,
            A5 => $generationRevenue,
            $totalDcp189DiscountedAssets
            ? (
                A31 => $rateDirect,
                A32 => $totalDcp189DiscountedAssets,
              )
            : (),
        }
    );
    $model->{transparency}{olFYI}{1248} = $chargeOther
      if $model->{transparency};

    my $rateOther = Arithmetic(
        name          => 'Other revenue charging rate',
        groupName     => 'Other revenue charging rate',
        arithmetic    => '=A1/(A21+A22+A3+A4)',
        defaultFormat => '%soft',
        arguments     => {
            A1  => $chargeOther,
            A21 => $totalAssetsCapacity,
            A22 => $totalAssetsConsumption,
            A3  => $cdcmEhvAssets,
            A4  => $cdcmHvLvShared,
        },
        location => 'Charging rates',
    );
    $model->{transparency}{olFYI}{1249} = $rateOther
      if $model->{transparency};

    my $totalRevenue3;

    if ( $model->{legacy201} && !$model->{dcp189} ) {

        my $fixed3contribution = Arithmetic(
            name          => 'Demand fixed pot contribution p/day',
            defaultFormat => '0.00softnz',
            arithmetic    => '=100/A2*A1*(A6+A7+A88)',
            arguments     => {
                A1  => $demandSoleUseAsset,
                A6  => $rateDirect,
                A7  => $rateIndirect,
                A88 => $rateRates,
                A2  => $daysInYear,
            }
        );

        my $capacity3 = Arithmetic(
            name          => 'Capacity pot contribution p/kVA/day',
            defaultFormat => '0.00softnz',
            arithmetic    => '=100/A3*((A1+A53)*(A6+A7+A8+A9)+A41*A42)',
            arguments     => {
                A3  => $daysInYear,
                A1  => $assetsCapacity,
                A53 => $assetsConsumption,
                A41 => $rateExit,
                A42 => ref $purpleUseRate eq 'ARRAY'
                ? $purpleUseRate->[0]
                : $purpleUseRate,
                A6 => $rateDirect,
                A7 => $rateIndirect,
                A8 => $rateRates,
                A9 => $rateOther,
            }
        );

        my $revenue3 = Arithmetic(
            name          => 'Pot contribution £/year',
            defaultFormat => '0softnz',
            arithmetic    => '=A9*0.01*(A1+A2*A3)',
            arguments     => {
                A1 => $fixed3contribution,
                A2 => $capacity3,
                A3 => $importCapacity,
                A9 => $daysInYear,
            }
        );

        $totalRevenue3 = GroupBy(
            name          => 'Pot £/year',
            defaultFormat => '0softnz',
            source        => $revenue3
        );

    }

    else {

        $totalRevenue3 = Arithmetic(
            name          => 'Demand revenue target pot (£/year)',
            newBlock      => 1,
            defaultFormat => '0softnz',
            arithmetic    => '=A5*A6'
              . '+(A11+A12+A13)*(A21+A22+A23)'
              . '+(A14+A15)*A24'
              . ( $totalDcp189DiscountedAssets ? '-A31*A32' : '' ),
            arguments => {
                A5  => $rateExit,
                A6  => $edcmPurpleUse,
                A11 => $totalAssetsFixed,
                A12 => $totalAssetsCapacity,
                A13 => $totalAssetsConsumption,
                A14 => $totalAssetsCapacity,
                A15 => $totalAssetsConsumption,
                A21 => $rateDirect,
                A22 => $rateRates,
                A23 => $rateIndirect,
                A24 => $rateOther,
                $totalDcp189DiscountedAssets
                ? (
                    A31 => $rateDirect,
                    A32 => $totalDcp189DiscountedAssets,
                  )
                : (),
            },
        );

    }

    $model->{transparency}{olFYI}{1201} = $totalRevenue3
      if $model->{transparency};

    push @{ $model->{calc3Tables} }, $totalRevenue3;

### Marker

    my ( $scalingChargeCapacity, $scalingChargeUnits );

    my $capacityChargeT = Arithmetic(
        name          => 'Capacity charge p/kVA/day (exit only)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=100/A2*A41*A1',
        arguments     => {
            A2  => $daysInYear,
            A41 => $rateExit,
            A1  => ref $purpleUseRate eq 'ARRAY'
            ? $purpleUseRate->[0]
            : $purpleUseRate,
        }
    );

    if ( $model->{matricesData} ) {
        push @{ $model->{matricesData}[0] },
          Arithmetic(
            name => "Notional $model->{timebandName} unit rate"
              . ' for transmission exit (p/kWh)',
            rows       => $tariffs->{rows},
            arithmetic => '=100/A2*A41*A9',
            arguments  => {
                A2  => $hoursInPurple,
                A41 => $rateExit,
                A9  => (
                    ref $purpleUseRate eq 'ARRAY'
                    ? $purpleUseRate->[0]
                    : $purpleUseRate
                )->{arguments}{A9},
            },
          );
        $model->{matricesData}[2] = $activeCoincidence;
        $model->{matricesData}[3] = $hoursInPurple;
        $model->{matricesData}[4] = $daysInYear;
    }

    $model->{summaryInformationColumns}[1] = Arithmetic(
        name          => 'Transmission exit charge (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=0.01*A9*A1*A2',
        arguments     => {
            A1 => $importCapacity,
            A2 => $capacityChargeT,
            A9 => $daysInYear,
            A7 => $tariffDaysInYearNot,
        },
    );

    my $importCapacityExceededAdjustment = Arithmetic(
        name =>
          'Adjustment to exceeded import capacity charge for DSM (p/kVA/day)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=IF(A1=0,0,(1-A4/A5)*(A3+'
          . 'IF(A23=0,0,(A2*A21*(A22-A24)/(A9-A91)))))',
        arguments => {
            A3  => $fcpLricDemandCapacityChargeBig,
            A4  => $chargeableCapacity,
            A5  => $importCapacity,
            A1  => $importCapacity,
            A2  => $unitRateFcpLricNonDSM,
            A21 => $activeCoincidence935,
            A23 => $activeCoincidence935,
            A22 => $hoursInPurple,
            A24 => $tariffHoursInPurpleNot,
            A9  => $daysInYear,
            A91 => $tariffDaysInYearNot,
        },
        defaultFormat => '0.00softnz'
    );

    push @{ $model->{calc2Tables} },
      my $unitRateFcpLricDSM = Arithmetic(
        name => "$model->{TimebandName} unit rate adjusted for DSM (p/kWh)",
        arithmetic => '=IF(A6=0,1,A4/A5)*A1',
        arguments  => {
            A1 => $unitRateFcpLricNonDSM,
            A4 => $chargeableCapacity,
            A5 => $importCapacity,
            A6 => $importCapacity,
        }
      );

    push @{ $model->{matricesData}[0] },
      Stack( sources => [$unitRateFcpLricDSM] )
      if $model->{matricesData};

### Marker

    my (
        $importCapacityScaledRound, $purpleRateFcpLricRound,
        $fixedDchargeTrueRound,     $importCapacityScaledSaved,
        $importCapacityExceeded,    $exportCapacityExceeded,
        $importCapacityScaled,      $purpleRateFcpLric,
    );

    my $demandScalingShortfall;

    if ( $model->{legacy201} ) {

        $capacityChargeT = Arithmetic(
            name       => 'Import capacity charge before scaling (p/kVA/day)',
            arithmetic => '=A7+IF(A6=0,1,A4/A5)*A1',
            defaultFormat => '0.00softnz',
            arguments     => {
                A1 => $fcpLricDemandCapacityChargeBig,
                A4 => $chargeableCapacity,
                A5 => $importCapacity,
                A6 => $importCapacity,
                A7 => $capacityChargeT,
            }
        );

        $model->{Thursday32} = [
            Arithmetic(
                name          => 'FCP/LRIC capacity-based charge (£/year)',
                arithmetic    => '=A1*A4*A9/100',
                defaultFormat => '0softnz',
                arguments     => {
                    A1 => $model->{demandCapacityFcpLric},
                    A4 => $chargeableCapacity,
                    A9 => $daysInYear,
                }
            ),
            Arithmetic(
                name          => 'FCP/LRIC unit-based charge (£/year)',
                arithmetic    => '=A1*A4*A9/100',
                defaultFormat => '0softnz',
                arguments     => {
                    A1 => $model->{demandConsumptionFcpLric},
                    A4 => $chargeableCapacity,
                    A9 => $daysInYear,
                }
            ),
        ];
        my $tariffHoursInPurple = Arithmetic(
            name => "Number of $model->{timebandName} hours connected in year",
            defaultFormat => '0.0softnz',
            arithmetic    => '=A2-A1',
            arguments     => {
                A2 => $hoursInPurple,
                A1 => $tariffHoursInPurpleNot,

            }
        );

        $demandScalingShortfall = Arithmetic(
            name          => 'Additional amount to be recovered (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => '=A1'
              . '-(SUM(A21_A22)+SUMPRODUCT(A31_A32,A33_A34)'
              . '+SUMPRODUCT(A41_A42,A43_A44,A35_A36,A51_A52)/A54'
              . ')*A9/100',
            arguments => {
                A1      => $totalRevenue3,
                A31_A32 => $capacityChargeT,
                A33_A34 => $importCapacity,
                A9      => $daysInYear,
                A21_A22 => $fixedDcharge,
                A41_A42 => $unitRateFcpLricDSM,
                A43_A44 => $activeCoincidence935,
                A35_A36 => $importCapacityUnscaled,
                A51_A52 => $tariffHoursInPurple,
                A54     => $daysInYear,
            }
        );

    }
    else {    # not legacy201

        push @{ $model->{calc2Tables} },
          my $capacityChargeT1 = Arithmetic(
            name          => 'Import capacity charge from charge 1 (p/kVA/day)',
            groupName     => 'Charge 1',
            arithmetic    => '=IF(A6=0,1,A4/A5)*A1',
            defaultFormat => '0.00softnz',
            arguments     => {
                A1 => $fcpLricDemandCapacityChargeBig,
                A4 => $chargeableCapacity,
                A5 => $importCapacity,
                A6 => $importCapacity,
            },
          );

        $capacityChargeT = Arithmetic(
            name       => 'Import capacity charge before scaling (p/kVA/day)',
            arithmetic => '=A7+A1',
            defaultFormat => '0.00softnz',
            arguments     => {
                A1 => $capacityChargeT1,
                A7 => $capacityChargeT,
            }
        );

        $model->{Thursday32} = [
            Arithmetic(
                name          => 'FCP/LRIC capacity-based charge (£/year)',
                arithmetic    => '=A1*A4*A9/100',
                defaultFormat => '0softnz',
                arguments     => {
                    A1 => $model->{demandCapacityFcpLric},
                    A4 => $chargeableCapacity,
                    A9 => $daysInYear,
                }
            ),
            Arithmetic(
                name          => 'FCP/LRIC unit-based charge (£/year)',
                arithmetic    => '=A1*A4*A9/100',
                defaultFormat => '0softnz',
                arguments     => {
                    A1 => $model->{demandConsumptionFcpLric},
                    A4 => $chargeableCapacity,
                    A9 => $daysInYear,
                }
            ),
        ];
        my $tariffHoursInPurple = Arithmetic(
            name => "Number of $model->{timebandName} hours connected in year",
            defaultFormat => '0.0softnz',
            arithmetic    => '=A2-A1',
            arguments     => {
                A2 => $hoursInPurple,
                A1 => $tariffHoursInPurpleNot,

            }
        );

        $demandScalingShortfall = Arithmetic(
            name          => 'Additional amount to be recovered (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => '=A1-A2*'
              . (
                $totalDcp189DiscountedAssets ? '(A42-A44)'
                : 'A42'
              )
              . '-A3*A43-A5*A6'
              . ( $model->{removeDemandCharge1} ? '' : '-A9' ),
            arguments => {
                A1  => $totalRevenue3,
                A2  => $rateDirect,
                A3  => $rateRates,
                A42 => $totalAssetsFixed,
                A43 => $totalAssetsFixed,
                $totalDcp189DiscountedAssets
                ? ( A44 => $totalDcp189DiscountedAssets )
                : (),
                A5 => $rateExit,
                A6 => $edcmPurpleUse,
                $model->{removeDemandCharge1} ? ()
                : (
                    A9 => $model->{transparencyMasterFlag} ? Arithmetic(
                        name => 'Revenue from demand charge 1 (£/year)',
                        defaultFormat => '0softnz',
                        arithmetic    => '=IF(A123,0,A1)+('
                          . 'SUMPRODUCT(A64_A65,A31_A32,A33_A34)+'
                          . 'SUMPRODUCT(A66_A67,A41_A42,A43_A44,A35_A36,A51_A52)/A54'
                          . ')*A9/100',
                        arguments => {
                            A123    => $model->{transparencyMasterFlag},
                            A1      => $model->{transparency}{ol119104},
                            A31_A32 => $capacityChargeT1,
                            A33_A34 => $importCapacity,
                            A9      => $daysInYear,
                            A41_A42 => $unitRateFcpLricDSM,
                            A43_A44 => $activeCoincidence935,
                            A35_A36 => $importCapacityUnscaled,
                            A51_A52 => $tariffHoursInPurple,
                            A54     => $daysInYear,
                            A64_A65 => $model->{transparency},
                            A66_A67 => $model->{transparency},
                        },
                      )
                    : Arithmetic(
                        name => 'Revenue from demand charge 1 (£/year)',
                        defaultFormat => '0softnz',
                        arithmetic    => '=('
                          . 'SUMPRODUCT(A31_A32,A33_A34)+'
                          . 'SUMPRODUCT(A41_A42,A43_A44,A35_A36,A51_A52)/A54'
                          . ')*A9/100',
                        arguments => {
                            A31_A32 => $capacityChargeT1,
                            A33_A34 => $importCapacity,
                            A9      => $daysInYear,
                            A41_A42 => $unitRateFcpLricDSM,
                            A43_A44 => $activeCoincidence935,
                            A35_A36 => $importCapacityUnscaled,
                            A51_A52 => $tariffHoursInPurple,
                            A54     => $daysInYear,
                        },
                    ),
                ),
            },
        );

        $model->{transparency}{olFYI}{1254} = $demandScalingShortfall
          if $model->{transparency};
        $model->{transparency}{olTabCol}{119104} =
          $demandScalingShortfall->{arguments}{A9}
          if $model->{transparency}
          && $demandScalingShortfall->{arguments}{A9};

    }

### Marker

    $model->fudge41(
        $activeCoincidence, $importCapacity,
        $edcmIndirect,      $edcmDirect,
        $edcmRates,         $daysInYear,
        \$capacityChargeT,  \$demandScalingShortfall,
        $indirectExposure,  $reactiveCoincidence,
        $powerFactorInModel,
    );

    push @{ $model->{calc4Tables} }, $demandScalingShortfall;

    ( $scalingChargeCapacity, $scalingChargeUnits ) = $model->demandScaling41(
        $importCapacity,       $demandScalingShortfall,
        $daysInYear,           $assetsFixed,
        $assetsCapacityCooked, $assetsConsumptionCooked,
        $capacityChargeT,      $fixedDcharge,
    );

    $model->{summaryInformationColumns}[2] = Arithmetic(
        name          => 'Direct cost allocation (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=A1*MAX(A2,'
          . '0-(A21+IF(A22=0,0,(1-A55/A54)*A31/(A32-A56)*IF(A52=0,1,A51/A53)*A5))'
          . ')*A3*0.01*A7/A9',
        arguments => {
            A1  => $importCapacity,
            A2  => $scalingChargeCapacity,
            A21 => $capacityChargeT,
            A22 => $activeCoincidence935,
            A5  => $demandConsumptionFcpLric,
            A51 => $chargeableCapacity,
            A52 => $importCapacity,
            A53 => $importCapacity,
            A3  => $daysInYear,
            A7  => $edcmDirect,
            A8  => $edcmRates,
            A9  => $demandScalingShortfall,
            A54 => $hoursInPurple,
            A55 => $tariffHoursInPurpleNot,
            A56 => $tariffDaysInYearNot,
            A31 => $daysInYear,
            A32 => $daysInYear,
        },
    );

    $model->{summaryInformationColumns}[4] = Arithmetic(
        name          => 'Network rates allocation (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=A1*MAX(A2,'
          . '0-(A21+IF(A22=0,0,(1-A55/A54)*A31/(A32-A56)*IF(A52=0,1,A51/A53)*A5))'
          . ')*A3*0.01*A8/A9',
        arguments => {
            A1  => $importCapacity,
            A2  => $scalingChargeCapacity,
            A21 => $capacityChargeT,
            A22 => $activeCoincidence935,
            A5  => $demandConsumptionFcpLric,
            A51 => $chargeableCapacity,
            A52 => $importCapacity,
            A53 => $importCapacity,
            A3  => $daysInYear,
            A7  => $edcmDirect,
            A8  => $edcmRates,
            A9  => $demandScalingShortfall,
            A54 => $hoursInPurple,
            A55 => $tariffHoursInPurpleNot,
            A56 => $tariffDaysInYearNot,
            A31 => $daysInYear,
            A32 => $daysInYear,
        },
    );

    $model->{summaryInformationColumns}[7] = Arithmetic(
        name          => 'Demand scaling asset based (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=A1*MAX(A2,'
          . '0-(A21+IF(A22=0,0,(1-A55/A54)*A31/(A32-A56)*IF(A52=0,1,A51/A53)*A5))'
          . ')*A3*0.01*(1-(A8+A7)/A9)',
        arguments => {
            A1  => $importCapacity,
            A2  => $scalingChargeCapacity,
            A21 => $capacityChargeT,
            A22 => $activeCoincidence935,
            A5  => $demandConsumptionFcpLric,
            A51 => $chargeableCapacity,
            A52 => $importCapacity,
            A53 => $importCapacity,
            A3  => $daysInYear,
            A7  => $edcmDirect,
            A8  => $edcmRates,
            A9  => $demandScalingShortfall,
            A54 => $hoursInPurple,
            A55 => $tariffHoursInPurpleNot,
            A56 => $tariffDaysInYearNot,
            A31 => $daysInYear,
            A32 => $daysInYear,
        },
    );

### Marker

    $importCapacityScaled =
      $scalingChargeCapacity
      ? Arithmetic(
        name          => 'Total import capacity charge p/kVA/day',
        defaultFormat => '0.00softnz',
        arithmetic    => '=MAX(0-(A3*A31*A33/A32),A1+A2)',
        arguments     => {
            A1  => $capacityChargeT,
            A3  => $unitRateFcpLricNonDSM,
            A31 => $activeCoincidence,
            A32 => $daysInYear,
            A33 => $hoursInPurple,
            A2  => $scalingChargeCapacity,
        }
      )
      : Stack( sources => [$capacityChargeT] );

    $purpleRateFcpLric = Arithmetic(
        name       => "$model->{TimebandName} rate p/kWh",
        arithmetic => '=IF(A3,IF(A1=0,A9,'
          . 'MAX(0,MIN(A4,A41+(A5/A11*(A7-A71)/(A8-A81))))' . '),0)',
        arguments => {
            A1  => $activeCoincidence,
            A11 => $activeCoincidence935,
            A3  => $importEligible,
            A4  => $unitRateFcpLricDSM,
            A41 => $unitRateFcpLricDSM,
            A9  => $unitRateFcpLricDSM,
            A5  => $importCapacityScaled,
            A51 => $demandConsumptionFcpLric,
            A7  => $daysInYear,
            A71 => $tariffDaysInYearNot,
            A8  => $hoursInPurple,
            A81 => $tariffHoursInPurpleNot,
        }
    ) if $unitRateFcpLricDSM;

    push @{ $model->{calc4Tables} },
      $importCapacityScaledSaved = $importCapacityScaled;

    $importCapacityScaled = Arithmetic(
        name       => 'Import capacity charge p/kVA/day',
        groupName  => 'Demand charges after scaling',
        arithmetic => '=IF(A3,MAX(0,A1),0)',
        arguments  => {
            A1 => $importCapacityScaled,
            A3 => $importEligible,
        },
        defaultFormat => '0.00softnz'
    );

    $importCapacityExceeded = Arithmetic(
        name          => 'Exceeded import capacity charge (p/kVA/day)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=A7+A2',
        defaultFormat => '0.00softnz',
        arguments     => {
            A3 => $fcpLricDemandCapacityChargeBig,
            A2 => $importCapacityExceededAdjustment,
            A4 => $chargeableCapacity,
            A5 => $importCapacity,
            A1 => $importCapacity,
            A7 => $importCapacityScaled,
        },
        defaultFormat => '0.00softnz'
    );

    $model->{summaryInformationColumns}[5] = Arithmetic(
        name          => 'FCP/LRIC charge (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=0.01*(A11*A9*A2+A1*A4*A8*(A6-A61)*(A91/(A92-A71)))',
        arguments     => {
            A1  => $importCapacity,
            A2  => $fcpLricDemandCapacityChargeBig,
            A3  => $capacityChargeT->{arguments}{A1},
            A9  => $daysInYear,
            A4  => $unitRateFcpLricDSM,
            A41 => $activeCoincidence,
            A6  => $hoursInPurple,
            A61 => $tariffHoursInPurpleNot,
            A8  => $activeCoincidence935,
            A91 => $daysInYear,
            A92 => $daysInYear,
            A71 => $tariffDaysInYearNot,
            A11 => $chargeableCapacity,
            A51 => $importCapacity,
            A62 => $importCapacity,

        },
    );

### Marker

    push @{ $model->{tablesG} }, $genCredit, $genCreditCapacity,
      $exportCapacityCharge;

    $fixedDchargeTrueRound = Arithmetic(
        name          => 'Import fixed charge (p/day)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=ROUND(A1,2)',
        arguments     => { A1 => $fixedDchargeTrue, },
    );

    $purpleRateFcpLricRound = Arithmetic(
        name          => "Import $model->{timebandName} unit rate (p/kWh)",
        defaultFormat => '0.000softnz',
        arithmetic    => '=ROUND(A1,3)',
        arguments     => { A1 => $purpleRateFcpLric, },
    );

    $importCapacityScaledRound = Arithmetic(
        name          => 'Import capacity rate (p/kVA/day)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=ROUND(A1,2)',
        arguments     => { A1 => $importCapacityScaled, },
    );

    $exportCapacityExceeded = Arithmetic(
        name          => 'Export exceeded capacity rate (p/kVA/day)',
        defaultFormat => '0.00copynz',
        arithmetic    => '=A1',
        arguments     => { A1 => $exportCapacityChargeRound, },
    );

    my $importCapacityExceededRound = Arithmetic(
        name          => 'Import exceeded capacity rate (p/kVA/day)',
        defaultFormat => '0.00softnz',
        arithmetic    => '=ROUND(A1,2)',
        arguments     => { A1 => $importCapacityExceeded, },
    );

    push @{ $model->{calc4Tables} }, $purpleRateFcpLric,
      $importCapacityScaled,
      $fixedDchargeTrue, $importCapacityExceeded,
      $exportCapacityChargeRound,
      $fixedGchargeTrue;

### Marker

    my @tariffColumns = (
        Stack( sources => [$tariffs] ),
        $purpleRateFcpLricRound,
        $fixedDchargeTrueRound,
        $importCapacityScaledRound,
        $importCapacityExceededRound,
        Stack( sources => [$genCreditRound] ),
        Stack( sources => [$fixedGchargeTrue] ),
        Stack( sources => [$netexportCapacityChargeRound] ),
        $exportCapacityExceeded,
    );

    my $allTariffColumns = $model->{checksums}
      ? [
        @tariffColumns,
        map {
            my $digits = /([0-9])/ ? $1 : 6;
            my $recursive = /table|recursive|model/i ? 1 : 0;
            $model->{"checksum_${recursive}_$digits"} =
              SpreadsheetModel::Checksum->new(
                name => $_,
                /table|recursive|model/i ? ( recursive => 1 ) : (),
                digits  => $digits,
                columns => [ @tariffColumns[ 1 .. 8 ] ],
                factors => [qw(1000 100 100 100 1000 100 100 100)]
              );
          } split /;\s*/,
        $model->{checksums}
      ]
      : \@tariffColumns;

    if ( $model->{layout} ) {

        $model->{tableList} = $model->orderedLayout(
            [ $exportEligible, @{ $tariffColumns[5]{sourceLines} } ],
            $model->{transparencyMasterFlag}
            ? [ $model->{transparencyMasterFlag}, ]
            : (),
            1 || $model->{layout} =~ /multisheet/i
            ? ( [ $gCharge, ], [ $rateExit, ], )
            : [ $gCharge, $rateExit, ],
            [
                @{ $tariffColumns[7]{sourceLines} },
                @{ $tariffColumns[8]{sourceLines} },
            ],
            [ $generationSoleUseAsset, $demandSoleUseAsset, ],
            [ $accretion, ],
            [ $assetsCapacity, ],
            [ $assetsConsumption, ],
            [ $rateDirect, $rateIndirect, $rateRates, ],
            [
                @{ $tariffColumns[6]{sourceLines} },
                @{ $tariffColumns[2]{sourceLines} },
                $fixedGcharge,
            ],
            [ $rateOther, ],
            [ $demandScalingShortfall, ],
            [ $assetsCapacityCooked, ],
            [ $assetsConsumptionCooked, ],
            [
                @{ $tariffColumns[1]{sourceLines} },
                @{ $tariffColumns[3]{sourceLines} },
                @{ $tariffColumns[4]{sourceLines} },
            ]
        );

    }

    if (    $model->{layout}
        and $model->{layout} =~ /no4501/i || $model->{tariff1Row} )
    {

        SpreadsheetModel::MatrixSheet->new(
            $model->{tariff1Row}
            ? (
                dataRow            => $model->{tariff1Row},
                captionDecorations => [qw(algae purple slime)],
              )
            : (),
          )->addDatasetGroup(
            name    => 'Tariff name',
            columns => [ $allTariffColumns->[0] ],
          )->addDatasetGroup(
            name    => 'Import tariff',
            columns => [ @{$allTariffColumns}[ 1 .. 4 ] ],
          )->addDatasetGroup(
            name    => 'Export tariff',
            columns => [ @{$allTariffColumns}[ 5 .. 8 ] ],
          )->addDatasetGroup(
            name    => 'Checksums',
            columns => [ @{$allTariffColumns}[ 9 .. $#$allTariffColumns ] ]
          );

        push @{ $model->{tariffTables} }, @$allTariffColumns;

    }

    else {
        push @{ $model->{tariffTables} },
          Columnset(
            name    => 'EDCM charge',
            columns => $allTariffColumns,
          );
    }

### Marker

    return $model unless $model->{summaries};

    my $format0withLine =
      $model->{vertical}
      ? '0soft'
      : [ base => '0soft', left => 5, left_color => 8 ];

    my @revenueBitsD = (

        Arithmetic(
            name          => 'Capacity charge for demand (£/year)',
            defaultFormat => $format0withLine,
            arithmetic    => '=0.01*A9*A8*A1',
            arguments     => {
                A1 => $importCapacityScaledRound,
                A9 => $daysInYear,
                A7 => $tariffDaysInYearNot,
                A8 => $importCapacity,
            }
        ),

        Arithmetic(
            name => "$model->{TimebandName} charge for demand (£/year)",
            defaultFormat => '0softnz',
            arithmetic    => '=0.01*(A9-A7)*A1*A6*(A91/(A92-A71))*A8',
            arguments     => {
                A1  => $purpleRateFcpLricRound,
                A9  => $hoursInPurple,
                A7  => $tariffHoursInPurpleNot,
                A6  => $importCapacity,
                A8  => $activeCoincidence935,
                A91 => $daysInYear,
                A92 => $daysInYear,
                A71 => $tariffDaysInYearNot
            }
        ),

        Arithmetic(
            name          => 'Fixed charge for demand (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => '=0.01*(A9-A7)*A1',
            arguments     => {
                A1 => $fixedDchargeTrueRound,
                A9 => $daysInYear,
                A7 => $tariffDaysInYearNot,
            }
        ),

    );

    my @revenueBitsG = (

        Arithmetic(
            name => 'Net capacity charge (or credit) for generation (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => '=0.01*A9*A8*A1',
            arguments     => {
                A1 => $netexportCapacityChargeRound,
                A9 => $daysInYear,
                A8 => $exportCapacityChargeable,
            }
        ),

        Arithmetic(
            name          => 'Fixed charge for generation (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => '=0.01*(A9-A7)*A1',
            arguments     => {
                A1 => $fixedGchargeTrue,
                A9 => $daysInYear,
                A7 => $tariffDaysInYearNot,
            }
        ),

        Arithmetic(
            name          => "$model->{TimebandName} credit (£/year)",
            defaultFormat => '0softnz',
            arithmetic    => '=0.01*A1*A6',
            arguments     => {
                A1 => $genCreditRound,
                A6 => $activeUnits,
            }
        ),

    );

    my $rev1d = Stack( sources => [$previousChargeImport] );

    my $rev1g = Stack( sources => [$previousChargeExport] );

    my $rev2d = Arithmetic(
        name          => 'Total for demand (£/year)',
        defaultFormat => $format0withLine,
        arithmetic    => '=' . join( '+', map { "A$_" } 1 .. @revenueBitsD ),
        arguments =>
          { map { ( "A$_" => $revenueBitsD[ $_ - 1 ] ) } 1 .. @revenueBitsD },
    );

    my $rev2g = Arithmetic(
        name          => 'Total for generation (£/year)',
        defaultFormat => $format0withLine,
        arithmetic    => '=' . join( '+', map { "A$_" } 1 .. @revenueBitsG ),
        arguments =>
          { map { ( "A$_" => $revenueBitsG[ $_ - 1 ] ) } 1 .. @revenueBitsG },
    );

    my $change1d = Arithmetic(
        name          => 'Change (demand) (£/year)',
        arithmetic    => '=A1-A4',
        defaultFormat => '0softpm',
        arguments     => { A1 => $rev2d, A4 => $rev1d }
    );

    my $change1g = Arithmetic(
        name          => 'Change (generation) (£/year)',
        arithmetic    => '=A1-A4',
        defaultFormat => '0softpm',
        arguments     => { A1 => $rev2g, A4 => $rev1g }
    );

    my $change2d = Arithmetic(
        name          => 'Change (demand) (%)',
        arithmetic    => '=IF(A1,A3/A4-1,"")',
        defaultFormat => '%softpm',
        arguments     => { A1 => $rev1d, A3 => $rev2d, A4 => $rev1d }
    );

    my $change2g = Arithmetic(
        name          => 'Change (generation) (%)',
        arithmetic    => '=IF(A1,A3/A4-1,"")',
        defaultFormat => '%softpm',
        arguments     => { A1 => $rev1g, A3 => $rev2g, A4 => $rev1g }
    );

    my $soleUseAssetChargeUnround = Arithmetic(
        name          => 'Fixed charge for demand (unrounded) (£/year)',
        defaultFormat => $format0withLine,
        arithmetic    => '=0.01*(A9-A7)*A1',
        arguments     => {
            A1 => $fixedDchargeTrue,
            A9 => $daysInYear,
            A7 => $tariffDaysInYearNot,
        }
    );

    $model->{summaryInformationColumns}[0] = $soleUseAssetChargeUnround;

    0 and my $purpleUnits = Arithmetic(
        name          => "$model->{TimebandName} units (kWh)",
        defaultFormat => '0softnz',
        arithmetic    => '=A1*(A3-A7)*A5',
        arguments     => {
            A3 => $hoursInPurple,
            A7 => $tariffHoursInPurpleNot,
            A1 => $activeCoincidence935,
            A5 => $importCapacityUnscaled,
        }
    );

    my $check = Arithmetic(
        name          => 'Check (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => join(
            '', '=A1',
            map {
                $model->{summaryInformationColumns}[$_]
                  ? ( "-A" . ( 20 + $_ ) )
                  : ()
            } 0 .. $#{ $model->{summaryInformationColumns} }
        ),
        arguments => {
            A1 => $rev2d,
            map {
                $model->{summaryInformationColumns}[$_]
                  ? ( "A" . ( 20 + $_ ),
                    $model->{summaryInformationColumns}[$_] )
                  : ()
            } 0 .. $#{ $model->{summaryInformationColumns} }
        }
    );

    if (   $model->{layout} && $model->{layout} =~ /matrix/i
        || $model->{summaries} =~ /no4601/i )
    {
        my @copyTariffs = map { Stack( sources => [$_] ) } @tariffColumns;
        SpreadsheetModel::MatrixSheet->new(
            $model->{tariff1Row}
            ? (
                dataRow            => $model->{tariff1Row},
                captionDecorations => [qw(algae purple slime)],
              )
            : (),
          )->addDatasetGroup(
            name    => 'Tariff name',
            columns => [ $copyTariffs[0] ],
          )->addDatasetGroup(
            name    => 'Import tariff',
            columns => [ @copyTariffs[ 1 .. 4 ] ],
          )->addDatasetGroup(
            name    => 'Export tariff',
            columns => [ @copyTariffs[ 5 .. 8 ] ],
          )->addDatasetGroup(
            name    => 'Import charges',
            columns => \@revenueBitsD,
          )->addDatasetGroup(
            name    => 'Export charges',
            columns => \@revenueBitsG,
          )->addDatasetGroup(
            name    => 'Change in import charges',
            columns => [ $rev2d, $rev1d, $change1d, $change2d, ],
          )->addDatasetGroup(
            name    => 'Change in export charges',
            columns => [ $rev2g, $rev1g, $change1g, $change2g, ],
          )->addDatasetGroup(
            name => 'Analysis of import charges',
            columns =>
              [ grep { $_ } @{ $model->{summaryInformationColumns} }, $check, ],
          );
        push @{ $model->{revenueTables} }, @copyTariffs, @revenueBitsD,
          @revenueBitsG, $change2d, $change2g, $check;
    }
    elsif ( $model->{vertical} ) {
        push @{ $model->{revenueTables} },
          Columnset(
            name    => 'Import charges',
            columns => [
                Stack( sources => [ $tariffColumns[0] ] ),
                @revenueBitsD, $rev2d, $rev1d, $change1d, $change2d,
            ],
          ),
          Columnset(
            name    => 'Export charges',
            columns => [
                Stack( sources => [ $tariffColumns[0] ] ),
                @revenueBitsG, $rev2g, $rev1g, $change1g, $change2g,
            ],
          ),
          Columnset(
            name    => 'Import charges based on sole-use assets',
            columns => [
                Stack( sources => [ $tariffColumns[0] ] ),
                grep { $_ } $model->{summaryInformationColumns}[0],
            ],
          ),
          Columnset(
            name    => 'Import charges based on non-sole-use notional assets',
            columns => [
                Stack( sources => [ $tariffColumns[0] ] ),
                grep { $_ } @{ $model->{summaryInformationColumns} }[ 2, 4, 7 ],
            ],
          ),
          Columnset(
            name    => 'Import charges based on capacity and consumption',
            columns => [
                Stack( sources => [ $tariffColumns[0] ] ),
                grep { $_ } @{ $model->{summaryInformationColumns} }[ 1, 3, 6 ],
            ],
          ),
          Columnset(
            name    => 'Other elements of import charges',
            columns => [
                Stack( sources => [ $tariffColumns[0] ] ),
                grep { $_ } $model->{summaryInformationColumns}[5],
                $check,
            ],
          );
    }
    else {
        push @{ $model->{revenueTables} },
          Columnset(
            name    => 'Horizontal information',
            columns => [
                ( map { Stack( sources => [$_] ) } @tariffColumns ),
                @revenueBitsD,
                @revenueBitsG,
                $rev2d,
                $rev1d,
                $change1d,
                $change2d,
                $rev2g,
                $rev1g,
                $change1g,
                $change2g,
                grep { $_ } @{ $model->{summaryInformationColumns} },
                $check,
            ],
          );
    }

    my $totalForDemandAllTariffs = GroupBy(
        source        => $rev2d,
        name          => 'Total for demand across all tariffs (£/year)',
        defaultFormat => '0softnz'
    );

    my $totalForGenerationAllTariffs = GroupBy(
        source        => $rev2g,
        name          => 'Total for generation across all tariffs (£/year)',
        defaultFormat => '0softnz'
    );

    if (   $model->{ldnoRevTotal}
        || $model->{summaries} && $model->{summaries} =~ /total/i )
    {

        push @{ $model->{ $model->{layout} ? 'TotalsTables' : 'revenueTables' }
          },
          my $totalAllTariffs = Columnset(
            name    => 'Total for all tariffs (£/year)',
            columns => [
                $model->{layout} ? () : Constant(
                    name          => 'This column is not used',
                    defaultFormat => '0con',
                    data          => [ [''] ]
                ),
                $totalForDemandAllTariffs,
                $totalForGenerationAllTariffs,
                Arithmetic(
                    name          => 'Total for all tariffs (£/year)',
                    defaultFormat => '0softnz',
                    arithmetic    => '=A1+A2',
                    arguments     => {
                        A1 => $totalForDemandAllTariffs,
                        A2 => $totalForGenerationAllTariffs,
                    }
                )
            ]
          );

        push @{ $model->{TotalsTables} },
          Columnset(
            name    => 'Total EDCM revenue (£/year)',
            columns => [
                Arithmetic(
                    name => 'All EDCM tariffs '
                      . 'including discounted CDCM tariffs (£/year)',
                    defaultFormat => '0softnz',
                    arithmetic    => '=A1+A2+A3',
                    arguments     => {
                        A1 => $totalForDemandAllTariffs,
                        A2 => $totalForGenerationAllTariffs,
                        A3 => $model->{ldnoRevTotal},
                    }
                )
            ]
          ) if $model->{ldnoRevTotal};

    }

    push @{ $model->{revenueTables} },
      $model->impactFinancialSummary( $tariffs, \@tariffColumns,
        $actualRedDemandRate, \@revenueBitsD, @revenueBitsG, $rev2g )
      if $model->{transparencyImpact};

    if ( $model->{transparency} ) {
        my %olTabCol;
        while ( my ( $num, $obj ) = each %{ $model->{transparency}{olTabCol} } )
        {
            my $number = int( $num / 100 );
            $olTabCol{$number}[ $num - $number * 100 - 1 ] = $obj;
        }
        $model->{aggregateTables} = [
            (
                map {
                    Columnset(
                        name          => "⇒$_->[0]. $_->[1]",
                        singleRowName => $_->[1],
                        number        => 3600 + $_->[0],
                        columns       => [
                            map { Stack( sources => [$_] ) }
                              @{ $olTabCol{ $_->[0] } }
                        ]
                      )
                  }[ 1191 => 'EDCM demand aggregates' ],
                [ 1192 => 'EDCM generation aggregates' ],
                [ 1193 => 'EDCM notional asset aggregates' ],
            ),
            (
                map {
                    my $obj  = $model->{transparency}{olFYI}{$_};
                    my $name = 'Copy of ' . $obj->{name};
                    $obj->isa('SpreadsheetModel::Columnset')
                      ? Columnset(
                        name    => $name,
                        number  => 3600 + $_,
                        columns => [
                            map { Stack( sources => [$_] ) }
                              @{ $obj->{columns} }
                        ]
                      )
                      : Stack(
                        name    => $name,
                        number  => 3600 + $_,
                        sources => [$obj]
                      );
                  } sort { $a <=> $b }
                  keys %{ $model->{transparency}{olFYI} }
            )
        ];
    }

### Marker

    $model->templates(
        $tariffs,                          $importCapacityUnscaled,
        $exportCapacityExempt,             $exportCapacityChargeablePre2005,
        $exportCapacityChargeable20052010, $exportCapacityChargeablePost2010,
        $tariffSoleUseMeav,                $tariffLoc,
        $tariffCategory,                   $useProportions,
        $activeCoincidence935,             $reactiveCoincidence,
        $indirectExposure,                 $nonChargeableCapacity,
        $activeUnits,                      $creditableCapacity,
        $tariffNetworkSupportFactor,       $tariffDaysInYearNot,
        $tariffHoursInPurpleNot,           $previousChargeImport,
        $previousChargeExport,             $llfcImport,
        $llfcExport,                       \@tariffColumns,
        $daysInYear,                       $hoursInPurple,
    ) if $model->{customerTemplates};

    $model;

}

1;
