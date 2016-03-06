package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others.
Copyright 2013-2016 Franck Latrémolière, Reckon LLP and others.

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

sub routeing {

    my ( $model, $allTariffsByEndUser, $coreLevels, $coreExitLevels, $drmLevels,
        $drmExitLevels, $rerouteingMatrix )
      = @_;

    my $customerLevels = Labelset
      name => 'Customer levels',
      list => [ 'LV customer', 'HV customer' ];

    my $networkLevels = Labelset
      name => 'Network levels (all)',
      list => [ 'GSPs', @{ $drmLevels->{list} }, @{ $customerLevels->{list} } ];

    my $assetLevels = Labelset
      name => 'Asset levels',
      list =>
      [ map { /GSP/i ? $_ : "Assets\t$_" } @{ $networkLevels->{list} } ],
      accepts => [$networkLevels];

    my $operatingLevels = Labelset
      name => 'Operating expenditure levels',
      list => [ map { /GSP/i ? "Transmission\texit" : "Operating\t$_" }
          @{ $networkLevels->{list} } ],
      accepts => [ $networkLevels, $assetLevels ];

    my $chargingLevels = Labelset
      name    => 'Charging levels (all)',
      list    => [ @{ $assetLevels->{list} }, @{ $operatingLevels->{list} } ],
      accepts => [$networkLevels];

    my $assetDrmLevels = Labelset
      name    => 'DRM asset levels',
      list    => [ map { "Assets\t$_" } @{ $drmLevels->{list} } ],
      accepts => [$drmLevels];

    my $operatingDrmLevels = Labelset
      name    => 'DRM operating levels',
      list    => [ map { "Operating\t$_" } @{ $drmLevels->{list} } ],
      accepts => [$assetDrmLevels];

    my $assetDrmExitLevels = Labelset
      name => 'DRM and exit asset levels',
      list =>
      [ map { /GSP/i ? $_ : "Assets\t$_" } @{ $drmExitLevels->{list} } ],
      accepts => [$drmExitLevels];

    my $operatingDrmExitLevels = Labelset
      name => 'DRM and exit operating levels',
      list => [ map { /GSP/i ? "Transmission\texit" : "Operating\t$_" }
          @{ $drmExitLevels->{list} } ],
      accepts => [ $drmExitLevels, $assetDrmExitLevels ];

    my $chargingDrmExitLevels = Labelset
      name => 'Charging levels (DRM and exit)',
      list => [ @{ $assetDrmExitLevels->{list} },
        @{ $operatingDrmExitLevels->{list} } ],
      accepts => [$drmExitLevels];

    my $assetCustomerLevels = Labelset
      name    => 'Asset customer levels',
      list    => [ map { "Assets\t$_" } @{ $customerLevels->{list} } ],
      accepts => [$customerLevels];

    my $operatingCustomerLevels = Labelset
      name    => 'Operating customer levels',
      list    => [ map { "Operating\t$_" } @{ $customerLevels->{list} } ],
      accepts => [ $customerLevels, $assetCustomerLevels ];

    my $customerChargingLevels = Labelset
      name => 'Customer charging levels',
      list => [
        @{ $assetCustomerLevels->{list} },
        @{ $operatingCustomerLevels->{list} }
      ],
      accepts => [$networkLevels];

    my $customerTypesForLosses = $coreLevels;

=head Old classification

Labelset(
        name => 'Supply levels to classify users for loss adjustment factors',
        list => [ split /\n/, <<'EOL' ]
LV
LV substation
HV
HV substation
33kV
132kV
Transmission
EOL
    );

=cut

    my $lineLossFactorsData = Dataset(
        data => [ reverse qw(1.08 1.05 1.04 1.02 1.015 1 1) ],
        name => Label(
            'Loss adjustment factor',
            'Loss adjustment factors to transmission'
        ),
        validation => {
            validate => 'decimal',
            criteria => '>',
            value    => 0
        },
        number   => 1032,
        appendTo => $model->{inputTables},
        dataset  => $model->{dataset},
        lines =>
'Source: losses model or loss adjustment factors at time of system peak.',
        rows => 0,
        cols => $customerTypesForLosses
    );

    0 and $lineLossFactorsData = Stack(
        cols    => $customerTypesForLosses,
        rows    => 0,
        name    => 'All loss adjustment factors to transmission',
        sources => [
            $lineLossFactorsData,
            Constant(
                name => '1 for GSPs',
                rows => 0,
                cols => Labelset(
                    list => [
                        grep { /trans|GSP/i }
                          @{ $customerTypesForLosses->{list} }
                    ]
                ),
                data => [ [1] ]
            )
        ]
    );

    my $customerTypeMatrixForLosses = Constant(
        rows => $allTariffsByEndUser,
        cols => $customerTypesForLosses,

        # hard-coded customer type = network type

        byrow => 1,
        data  => [
            map {
                local $_ = $_;
                s/^LDNO.*?:\s*//;
                    /^LV sub/i   ? [ reverse qw(0 1 0 0 0 0 0) ]
                  : /^LV/i       ? [ reverse qw(1 0 0 0 0 0 0) ]
                  : /^HV sub/i   ? [ reverse qw(0 0 0 1 0 0 0) ]
                  : /^HV/i       ? [ reverse qw(0 0 1 0 0 0 0) ]
                  : /^33kV sub/i ? [ reverse qw(0 0 0 0 0 1 0) ]
                  : /^33/i       ? [ reverse qw(0 0 0 0 1 0 0) ]
                  : /^132/i      ? [ reverse qw(0 0 0 0 0 0 1) ]
                  : []
            } @{ $allTariffsByEndUser->{list} }
        ],
        defaultFormat => '0connz',
        name          => 'Network level for each tariff '
          . '(to get loss factors applicable to capacity)',
    );

    my $lineLossFactorsToGsp = SumProduct(
        name => Label(
            'Loss adjustment factor',
            'Loss adjustment factor to transmission'
        ),
        matrix => $customerTypeMatrixForLosses,
        vector => $lineLossFactorsData,
    );

    $lineLossFactorsToGsp = Arithmetic(
        name => Label(
            'Loss adjustment factor',
            'Loss adjustment factor to transmission'
        ),
        arithmetic => '=IF(A1,A2,1)',
        arguments =>
          { A1 => $lineLossFactorsToGsp, A2 => $lineLossFactorsToGsp }
    ) if $model->{ehv};

    Columnset(
        name    => 'Loss adjustment factors to transmission',
        columns => [
            $customerTypeMatrixForLosses,
            $model->{ehv}
            ? $lineLossFactorsToGsp->{arguments}{A1}
            : (),
            $lineLossFactorsToGsp
        ]
    );

    my $unitsLossAdjustment;

    if ( $model->{boundary} || $model->{portfolio}
        and !$model->{noLossAdj} )
    {

        my $customerTypeMatrixForLossesUnits = Constant(
            rows => $allTariffsByEndUser,
            cols => $customerTypesForLosses,

            # hard-coded customer type = network type

            byrow => 1,
            data  => [
                map {
                    local $_ = $_;
                        /^LDNO LV/i          ? [ reverse qw(0 1 0 0 0 0 0) ]
                      : /^LV sub/i           ? [ reverse qw(0 1 0 0 0 0 0) ]
                      : /^LV/i               ? [ reverse qw(1 0 0 0 0 0 0) ]
                      : /^(LDNO )?HV sub/i   ? [ reverse qw(0 0 0 1 0 0 0) ]
                      : /^(LDNO )?HV/i       ? [ reverse qw(0 0 1 0 0 0 0) ]
                      : /^(LDNO )?33kV sub/i ? [ reverse qw(0 0 0 0 0 1 0) ]
                      : /^(LDNO )?33/i       ? [ reverse qw(0 0 0 0 1 0 0) ]
                      : /^(LDNO )?132/i      ? [ reverse qw(0 0 0 0 0 0 1) ]
                      : []
                } @{ $allTariffsByEndUser->{list} }
            ],
            defaultFormat => '0connz',
            name          => 'Network level for each tariff '
              . '(to get loss factors applicable to units)',
        );

        my $lineLossFactorsToGspUnits = SumProduct(
            name => Label(
                'Loss adjustment factor',
                'Loss adjustment factor to transmission for units'
            ),
            matrix => $customerTypeMatrixForLossesUnits,
            vector => $lineLossFactorsData,
        );

        $unitsLossAdjustment = Arithmetic(
            name => 'Adjustment to be applied to unit rates for losses'
              . ' (embedded network tariffs)',
            defaultFormat => '%softnz',
            arithmetic    => '=A1/A2-1',
            arguments     => {
                A1 => $lineLossFactorsToGspUnits,
                A2 => $lineLossFactorsToGsp
            }
        );

        push @{ $model->{routeing} },
          Columnset(
            name    => 'Additional loss adjustment to calculate unit charges',
            columns => [
                $customerTypeMatrixForLossesUnits, $lineLossFactorsToGspUnits,
                $unitsLossAdjustment
            ]
          );

    }

    my $lineLossFactorsLevel = Stack(
        cols => $coreExitLevels,
        rows => 0,
        name => 'Loss adjustment factor to transmission'
          . ' for each core level',
        sources => [
            $lineLossFactorsData,
            Constant(
                name => '1 for GSP level',
                rows => 0,
                cols => Labelset(
                    list =>
                      [ grep { /trans|GSP/i } @{ $coreExitLevels->{list} } ]
                ),
                data => [ [ map { 1 } 1 .. 20 ] ]
            ),
        ]
    );

    my $lineLossFactorsNetwork = $lineLossFactorsLevel;

    $lineLossFactorsNetwork = Stack(
        cols => $drmExitLevels,
        rows => 0,
        name => 'Loss adjustment factor to transmission'
          . ' for each network level',
        sources => [
            SumProduct(
                name => 'Loss adjustment factor to transmission'
                  . ' for each DRM network level',
                vector => $lineLossFactorsData,
                matrix => Constant(
                    name =>
                      'Mapping of DRM network levels to core network levels',
                    rows          => $drmLevels,
                    cols          => $coreLevels,
                    defaultFormat => '0connz',
                    data          => [
                        [qw(1 0 0 0 0 0 0 0)], [qw(0 1 0 0 0 0 0 0)],
                        [qw(0 0 1 0 0 0 0 0)], [qw(0 0 0 1 1 0 0 0)],
                        [qw(0 0 0 0 0 1 0 0)], [qw(0 0 0 0 0 0 1 0)],
                        [qw(0 0 0 0 0 0 0 1)],
                    ]
                )
            ),
            Constant(
                name => '1 for GSP level',
                rows => 0,
                cols => Labelset(
                    list =>
                      [ grep { /trans|GSP/i } @{ $drmExitLevels->{list} } ]
                ),
                data => [ [ map { 1 } 1 .. 20 ] ]
            ),
        ]
    ) if $rerouteingMatrix;

    0 and $lineLossFactorsNetwork = SumProduct(
        vector => Constant(
            name  => 'Mapping of network levels to customer types for losses',
            lines => 'Loss factors for service cables are neglected.',
            rows  => $coreExitLevels,
            cols  => $customerTypesForLosses,
            byrow => 1,
            data  => [
                map {
                        /HV\s*\/\s*LV/i           ? [qw(0 1 0 0 0 0 0)]
                      : /^LV/i                    ? [qw(1 0 0 0 0 0 0)]
                      : /EHV\s*\/\s*HV/i          ? [qw(0 0 0 1 0 0 0)]
                      : /^HV/i                    ? [qw(0 0 1 0 0 0 0)]
                      : /132(\s*kV)?\s*\/\s*EHV/i ? [qw(0 0 0 0 1 0 0)]
                      : /^EHV/i                   ? [qw(0 0 0 0 1 0 0)]
                      : /^132/i                   ? [qw(0 0 0 0 0 1 0)]
                      : /^(Trans|GSP)/i           ? [qw(0 0 0 0 0 0 1)]
                      : []
                } @{ $coreExitLevels->{list} }
            ],
            defaultFormat => '0connz',
        ),
        matrix => $lineLossFactorsData,
        name => 'Loss adjustment factor to transmission for each network level',
    );

    push @{ $model->{edcmTables} },
      Stack(
        name          => 'Loss adjustment factor to transmission',
        defaultFormat => '0.000hard',
        number        => 1135,
        cols          => Labelset(
            list => [ @{ $lineLossFactorsNetwork->{cols}{list} }[ 0 .. 5 ] ]
        ),
        sources => [$lineLossFactorsNetwork]
      ) if $model->{edcmTables};

    my $lineLossFactorsPure = Arithmetic(
        cols       => $coreExitLevels,
        rows       => $allTariffsByEndUser,
        arithmetic => '=A1/A2',
        arguments =>
          { A1 => $lineLossFactorsToGsp, A2 => $lineLossFactorsNetwork },
        name =>
          'Loss adjustment factors between end user meter reading and each'
          . ' network level (by tariff)',
    );

    my $idnoDataInputTariffs = Labelset(
        name => 'LV LDNO demand portfolio tariffs',
        list => [
            grep { /^LDNO LV/ && !/^LDNO LV B[^:]*[0-9]/i && !/generat/i }
              @{ $allTariffsByEndUser->{list} }
        ]
    );

    undef $idnoDataInputTariffs unless @{ $idnoDataInputTariffs->{list} };

    my $routeingFactors = Constant(
        rows  => $allTariffsByEndUser,
        cols  => $coreExitLevels,
        byrow => 1,
        data  => [
            map {
                my $ar =
                  /EHV-(local|matched)/i
                  ? (
                    /^(LDNO )?LV sub/i
                    ? [ 0, 0, 0, 1, 1, 1, 1, 0, 0, 0 ]
                    : /^LV (additional|related)/i
                    ? [ 0, 0, 0, 1, 1, 1, 1, 1, 0, 0 ]
                    : $model->{boundary} && /^LDNO LV B[^:]*([0-9]+)/i ? [
                        0, 0, 0, 1, 1, 1, 1, ( $1 - 1 ) / $model->{boundary},
                        0, 0
                      ]
                    : /^LDNO LV/i        ? [ 0, 0, 0, 1, 1, 1, 1, undef, 0, 0 ]
                    : /^LV/i             ? [ 0, 0, 0, 1, 1, 1, 1, 1,     1, 0 ]
                    : /^(LDNO )?HV sub/i ? [ 0, 0, 0, 1, 1, 0, 0, 0,     0, 0 ]
                    : /^LDNO HV/i        ? [ 0, 0, 0, 1, 1, 1, 0, 0,     0, 0 ]
                    : /^HV (additional|related)/i
                    ? [ 0, 0, 0, 1, 1, 1, 0, 0, 0, 0 ]
                    : /^HV/i ? [ 0, 0, 0, 1, 1, 1, 0, 0, 0, 1 ]
                    :          die $_
                  )
                  : /^(LDNO )?LV sub.*generat/i
                  ? [ 1, 1, 1, 1, 1, 1, 0, 0, 1, 0 ]
                  : /^(LDNO )?LV.*generat/i ? [ 1, 1, 1, 1, 1, 1, 1, 0, 1, 0 ]
                  : /^(LDNO HV.*: )?HV sub.*generat/i
                  ? [ 1, 1, 1, 1, 0, 0, 0, 0, 0, 1 ]
                  : /^(LDNO HV.*: )?HV.*generat/i
                  ? [ 1, 1, 1, 1, 1, 0, 0, 0, 0, 1 ]
                  : /^(LDNO 33.*: )?33kV sub.*generat/i
                  ? [ 1, 1, 0, 0, 0, 0, 0, 0, 0, 0 ]
                  : /^(LDNO 33.*: )?33.*generat/i
                  ? [ 1, 1, 1, 0, 0, 0, 0, 0, 0, 0 ]
                  : /^(LDNO 132.*: )?132.*generat/i
                  ? [ 1, 0, 0, 0, 0, 0, 0, 0, 0, 0 ]
                  : /^GSP.*generat/i   ? [ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 ]
                  : /^(LDNO )?LV sub/i ? [ 1, 1, 1, 1, 1, 1, 1, 0, 0, 0 ]
                  : /^LV (additional|related)/i
                  ? [ 1, 1, 1, 1, 1, 1, 1, 1, 0, 0 ]
                  : $model->{boundary} && /^LDNO LV B[^:]*([0-9]+)/i
                  ? [ 1, 1, 1, 1, 1, 1, 1, ( $1 - 1 ) / $model->{boundary}, 0,
                    0 ]
                  : /^LDNO LV/i        ? [ 1, 1, 1, 1, 1, 1, 1, undef, 0, 0 ]
                  : /^LV/i             ? [ 1, 1, 1, 1, 1, 1, 1, 1,     1, 0 ]
                  : /^(LDNO )?HV sub/i ? [ 1, 1, 1, 1, 1, 0, 0, 0,     0, 0 ]
                  : /^LDNO HV/i        ? [ 1, 1, 1, 1, 1, 1, 0, 0,     0, 0 ]
                  : /^HV (additional|related)/i
                  ? [ 1, 1, 1, 1, 1, 1, 0, 0, 0, 0 ]
                  : /^HV/i               ? [ 1, 1, 1, 1, 1, 1, 0, 0, 0, 1 ]
                  : /^(LDNO )?33kV sub/i ? [ 1, 1, 1, 0, 0, 0, 0, 0, 0, 0 ]
                  : /^(LDNO )?33/i       ? [ 1, 1, 1, 1, 0, 0, 0, 0, 0, 0 ]
                  : /^(LDNO )?132/i      ? [ 1, 1, 0, 0, 0, 0, 0, 0, 0, 0 ]
                  : /^GSP/i              ? [ 1, 0, 0, 0, 0, 0, 0, 0, 0, 0 ]
                  :                        [];
                if (/([0-9]+)% credit/i) {
                    my $factor = 0.01 * $1;
                    $_ *= $factor foreach @$ar;
                }
                $ar;
            } @{ $allTariffsByEndUser->{list} }
        ],
        defaultFormat => '0.000connz',
        name          => $idnoDataInputTariffs
        ? Label(
            'Default network use factors',
            'Network use factor'
          )
        : Label(
            'Network use factors',
            'Network use factor'
        ),
        lines => <<'EOT'
These network use factors indicate to what extent each network level is used by each tariff. This table reflects the policy that
generators receive credits only in respect of network levels above the voltage of connection. Generators do not receive credits at the
voltage of connection. The factors in this table are before any adjustment for a 132kV/HV network level or for generation-dominated areas.
EOT
    );

    if ( $model->{ldnoSplits} ) {
        my $splits = Dataset(
            name          => 'HV and LV DNO main usage for LDNO tariffs',
            cols          => Labelset( list => [ 'LV circuits', 'HV' ] ),
            data          => [ [0.1], [0.4] ],
            defaultFormat => '%hard',
            number        => 1036,
            dataset       => $model->{dataset},
            appendTo      => $model->{inputTables},
            validation    => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
        );
        $routeingFactors = Stack(
            rows    => $allTariffsByEndUser,
            cols    => $coreExitLevels,
            sources => [
                Arithmetic(
                    name => 'LDNO LV split',
                    rows => Labelset(
                        list => [
                            grep { /^LDNO LV: LV/i && !/^LDNO LV: LV Sub/i; }
                              @{ $allTariffsByEndUser->{list} }
                        ]
                    ),
                    cols       => Labelset( list => ['LV circuits'] ),
                    arithmetic => '=A1',
                    arguments  => { A1           => $splits }
                ),
                Arithmetic(
                    name => 'LDNO HV split',
                    rows => Labelset(
                        list => [
                            grep { /^LDNO HV: [LH]V/i && !/^LDNO HV: HV Sub/i; }
                              @{ $allTariffsByEndUser->{list} }
                        ]
                    ),
                    cols       => Labelset( list => ['HV'] ),
                    arithmetic => '=A1',
                    arguments  => { A1           => $splits }
                ),
                $routeingFactors
            ],
        );
    }

    if ($rerouteingMatrix) {
        $routeingFactors = SumProduct(
            name =>
'Network use factors: interim step in calculations before adjustments',
            vector => $rerouteingMatrix,
            matrix => $routeingFactors
        );
        my $hvSubTariffs = Labelset(
            name => 'HV Sub tariffs',
            list =>
              [ grep { /^(LDNO )?HV Sub/i } @{ $allTariffsByEndUser->{list} } ]
        );
        my $gdTariffs = Labelset(
            name => 'Generation dominated tariffs',
            list => [ grep { /\(GD[PT]\)/i } @{ $allTariffsByEndUser->{list} } ]
        );
        $routeingFactors = Stack(
            name    => 'Network use factors for all tariffs',
            rows    => $allTariffsByEndUser,
            cols    => $drmExitLevels,
            sources => [
                Constant(
                    name => 'Network use factors including 132kV/HV'
                      . ' for generation dominated tariffs',
                    rows  => $gdTariffs,
                    cols  => $drmExitLevels,
                    byrow => 1,
                    data  => [
                        map {
                            /GDT/i
                              ? (
                                /^(LDNO )LV.*generat/i
                                ? [ 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0 ]
                                : [ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1 ]
                              )
                              : (
                                /^(LDNO )?LV.*generat/i
                                ? [ 1, 1, 1, 0, 0, 0, 0, 0, 0, 1, 0 ]
                                : /^(LDNO HV.*: )?HV.*generat/i
                                ? [ 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 1 ]
                                : [ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 ]
                              )
                        } @{ $gdTariffs->{list} }
                    ]
                ),
                Constant(
                    name => 'Network use factors including 132kV/HV'
                      . ' for HV Sub tariffs',
                    rows  => $hvSubTariffs,
                    cols  => $drmExitLevels,
                    byrow => 1,
                    data  => [
                        map {
                            /132/
                              ? (
                                /generat/i
                                ? [ 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1 ]
                                : [ 1, 1, 0, 0, 0, 1, 0, 0, 0, 0, 1 ]
                              )
                              : (
                                /generat/i
                                ? [ 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 1 ]
                                : [ 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 1 ]
                              )
                        } @{ $hvSubTariffs->{list} }
                    ]
                ),
                $routeingFactors
            ]
        );
    }

    if ( !$model->{ldnoSplits} && $idnoDataInputTariffs ) {   # old way, LV only

        $routeingFactors = Stack(
            rows    => $allTariffsByEndUser,
            cols    => $drmExitLevels,
            name    => Label( 'Network use factors', 'Network use factor' ),
            sources => [
                Dataset(
                    name => 'Use of DNO\'s LV network by LV-connected'
                      . ' embedded networks',
                    validation => {
                        validate      => 'decimal',
                        criteria      => 'between',
                        minimum       => 0,
                        maximum       => 4,
                        error_message => 'Must be'
                          . ' a non-negative percentage value.',
                    },
                    number        => 1035,
                    appendTo      => $model->{inputTables},
                    dataset       => $model->{dataset},
                    rows          => $idnoDataInputTariffs,
                    defaultFormat => '%hard',
                    data          => [
                        map {
                                /band 1/i ? 0
                              : /band 2/i ? 0.25
                              : /band 3/i ? 0.5
                              : /band 4/i ? 0.75
                              : 0.15
                        } @{ $idnoDataInputTariffs->{list} }
                    ],
                    cols => Labelset( list => ['LV circuits'], data => [0] )
                ),
                $routeingFactors
            ],
            lines => <<'EOT'
These network use factors indicate to what extent each network level is used by each tariff.
This table reflects the policy that generators receive credits only in respect of network levels above the voltage of connection. Generators do not receive credits at the voltage of connection.
The extent to which embedded networks use the DNO's network at the boundary voltage is a data input.
EOT
        );

    }

    my $lineLossFactors = Arithmetic
      name => 'Loss adjustment factors between end user meter reading and each'
      . ' network level, scaled by network use',
      defaultFormat => '0.000softnz',
      arithmetic    => '=IF(A4="",A5,A1*A2/A3)',
      arguments     => {
        A1 => $routeingFactors,
        A2 => $lineLossFactorsToGsp,
        A3 => $lineLossFactorsNetwork,
        A4 => $lineLossFactorsNetwork,
        A5 => $routeingFactors,
      };

    push @{ $model->{routeing} },
      $lineLossFactorsToGsp,
      $lineLossFactorsNetwork, $routeingFactors, $lineLossFactors;

    $assetCustomerLevels, $assetDrmLevels, $assetLevels,
      $chargingDrmExitLevels,
      $chargingLevels, $customerChargingLevels, $customerLevels,
      $networkLevels,
      $operatingCustomerLevels, $operatingDrmExitLevels, $operatingDrmLevels,
      $operatingLevels,         $routeingFactors,        $lineLossFactorsToGsp,
      $lineLossFactorsLevel,    $lineLossFactorsNetwork, $lineLossFactors,
      $unitsLossAdjustment;

}

1;
