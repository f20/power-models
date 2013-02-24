package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others.
Copyright 2011-2012 Franck Latrémolière, Reckon LLP and others.

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

sub diversity {

    my (
        $model,                  $demandEndUsers,
        $demandTariffsByEndUser, $unitsInYear,
        $loadFactors,            $daysInYear,
        $lineLossFactors,        $diversityAllowances,
        $componentMap,           $volumeData,
        $powerFactorInModel,     $forecastSml,
        $drmExitLevels,          $coreExitLevels,
        $rerouteing13211
    ) = @_;

    push @{ $model->{optionLines} }, !$model->{standing}
      || $model->{standing} =~ /sub/i
      ? 'Standing charges factors: '
      . '100/0/0 LV NHH, '
      . '100/100/20 network, '
      . '100/100/0 substation'
      : $model->{standing} =~ /nhh/i
      ? 'Standing charges factors: 100/0/0 for LV NHH, 100/100/20 for others'
      : $model->{standing} =~ /mig12/i
      ? 'Standing charges factors: 100/0/0 for PC1-4, 100/100/20 network, 100/100/0 substation'
      : $model->{standing} =~ /pc1-?4/i
      ? 'Standing charges factors: 100/0/0 for PC1-4, 100/100/20 for others'
      : $model->{standing} =~ /low/i
      ? 'Standing charges factors: 100/0/0 for everyone'
      : $model->{standing} =~ /reduc/i
      ? 'Standing charges factors: 100/100/20 for everyone'
      : $model->{standing} =~ /edf/i
      ? 'Standing charges factors: 100/50/0 for everyone'
      : $model->{standing} =~ /g3/i
      ? 'Standing charges factors: 100/100/0 for everyone'
      : 'No standing charges factors';

    push @{ $model->{optionLines} }, 'All costs go to capacity for EHV demand'
      if $model->{ehv}
      && $model->{ehv} =~ /cap|33/i;

    my $standingFactors = Constant(
        rows  => $demandEndUsers,
        cols  => $coreExitLevels,
        byrow => 1,
        data  => !$model->{standing} || $model->{standing} =~ /sub/i
        ? [
            map {
                /unmeter|generat/i
                  ? [ map { 0 } 1 .. 8 ]
                  : /LV sub/i ? (
                    $componentMap->{$_}{'Capacity charge p/kVA/day'}
                    ? [qw(0 0 0 0 0 1 1 0)]
                    : [qw(0 0 0 0 0 0 1 0)]
                  )
                  : /LV/i ? (
                    $componentMap->{$_}{'Capacity charge p/kVA/day'}
                    ? [qw(0 0 0 0 0 .2 1 1)]
                    : [qw(0 0 0 0 0 0 0 1)]
                  )
                  : /HV sub/i   ? [qw(0 0 0 1 1 0 0 0)]
                  : /HV/i       ? [qw(0 0 0 .2 1 1 0 0)]
                  : /33kV sub/i ? [
                    $model->{ehv} && $model->{ehv} =~ /cap|33/i
                    ? qw(1 1 1 0 0 0 0 0)
                    : qw(0 1 1 0 0 0 0 0)
                  ]
                  : /33/ ? [
                    $model->{ehv} && $model->{ehv} =~ /cap|33/i
                    ? qw(1 1 1 1 0 0 0 0)
                    : qw(0 .2 1 1 0 0 0 0)
                  ]
                  : /132/  ? [qw(1 1 0 0 0 0 0 0)]
                  : /GSP/i ? [qw(1 0 0 0 0 0 0 0)]
                  : [ map { 0 } 1 .. 8 ]
            } @{ $demandEndUsers->{list} }
          ]
        : $model->{standing} =~ /nhh/i ? [
            map {
                /unmeter|generat/i
                  ? [ map { 0 } 1 .. 8 ]
                  : /LV sub/i ? (
                    $componentMap->{$_}{'Capacity charge p/kVA/day'}
                    ? [qw(0 0 0 0 0 .2 1 0)]
                    : [qw(0 0 0 0 0 0 1 0)]
                  )
                  : /LV/i ? (
                    $componentMap->{$_}{'Capacity charge p/kVA/day'}
                    ? [qw(0 0 0 0 0 .2 1 1)]
                    : [qw(0 0 0 0 0 0 0 1)]
                  )
                  : /HV sub/i   ? [qw(0 0 0 .2 1 0 0 0)]
                  : /HV/i       ? [qw(0 0 0 .2 1 1 0 0)]
                  : /33kV sub/i ? [
                    $model->{ehv} && $model->{ehv} =~ /cap|33/i
                    ? qw(1 1 1 0 0 0 0 0)
                    : qw(0 .2 1 0 0 0 0 0)
                  ]
                  : /33/ ? [
                    $model->{ehv} && $model->{ehv} =~ /cap|33/i
                    ? qw(1 1 1 1 0 0 0 0)
                    : qw(0 .2 1 1 0 0 0 0)
                  ]
                  : /132/  ? [qw(1 1 0 0 0 0 0 0)]
                  : /GSP/i ? [qw(1 0 0 0 0 0 0 0)]
                  : [ map { 0 } 1 .. 8 ]
            } @{ $demandEndUsers->{list} }
          ]
        : $model->{standing} =~ /mig12/i ? [
            map {
                    /unmeter|generat/i ? [ map { 0 } 1 .. 8 ]
                  : /LV Sub/i          ? [qw(0 0 0 0 0 1 1 0)]
                  : /LV/i              ? (
                    $componentMap->{$_}{'Capacity charge p/kVA/day'}
                      || /(profile|pc).*[5-8]|medium/i
                    ? [qw(0 0 0 0 0 .2 1 1)]
                    : [qw(0 0 0 0 0 0 0 1)]
                  )
                  : /HV sub/i   ? [qw(0 0 0 1 1 0 0 0)]
                  : /HV/i       ? [qw(0 0 0 .2 1 1 0 0)]
                  : /33kV sub/i ? [
                    $model->{ehv} && $model->{ehv} =~ /cap|33/i
                    ? qw(1 1 1 0 0 0 0 0)
                    : qw(0 1 1 0 0 0 0 0)
                  ]
                  : /33/ ? [
                    $model->{ehv} && $model->{ehv} =~ /cap|33/i
                    ? qw(1 1 1 1 0 0 0 0)
                    : qw(0 .2 1 1 0 0 0 0)
                  ]
                  : /132/  ? [qw(1 1 0 0 0 0 0 0)]
                  : /GSP/i ? [qw(1 0 0 0 0 0 0 0)]
                  : [ map { 0 } 1 .. 8 ]
            } @{ $demandEndUsers->{list} }
          ]
        : $model->{standing} =~ /pc1-?4/i ? [
            map {
                /unmeter|generat/i
                  ? [ map { 0 } 1 .. 8 ]
                  : /LV/i ? (
                    $componentMap->{$_}{'Capacity charge p/kVA/day'}
                      || /(profile|pc).*[5-8]|medium/i
                    ? [qw(0 0 0 0 0 .2 1 1)]
                    : [qw(0 0 0 0 0 0 0 1)]
                  )
                  : /HV/i       ? [qw(0 0 0 .2 1 1 0 0)]
                  : /33kV sub/i ? [
                    $model->{ehv} && $model->{ehv} =~ /cap|33/i
                    ? qw(1 1 1 0 0 0 0 0)
                    : qw(0 0.2 1 0 0 0 0 0)
                  ]
                  : /33/ ? [
                    $model->{ehv} && $model->{ehv} =~ /cap|33/i
                    ? qw(1 1 1 1 0 0 0 0)
                    : qw(0 .2 1 1 0 0 0 0)
                  ]
                  : /132/  ? [qw(1 1 0 0 0 0 0 0)]
                  : /GSP/i ? [qw(1 0 0 0 0 0 0 0)]
                  : [ map { 0 } 1 .. 8 ]
            } @{ $demandEndUsers->{list} }
          ]
        : $model->{standing} =~ /low/i ? [
            map {
                    /unmeter|generat/i ? [ map { 0 } 1 .. 8 ]
                  : /LV sub/i          ? [qw(0 0 0 0 0 0 1 0)]
                  : /LV/i              ? [qw(0 0 0 0 0 0 0 1)]
                  : /HV sub/i          ? [qw(0 0 0 0 1 0 0 0)]
                  : /HV/i              ? [qw(0 0 0 0 0 1 0 0)]
                  : /33kV sub/i        ? [
                    $model->{ehv} && $model->{ehv} =~ /cap|33/i
                    ? qw(1 1 1 0 0 0 0 0)
                    : qw(0 0 1 0 0 0 0 0)
                  ]
                  : /33/ ? [
                    $model->{ehv} && $model->{ehv} =~ /cap|33/i
                    ? qw(1 1 1 1 0 0 0 0)
                    : qw(0 0 0 1 0 0 0 0)
                  ]
                  : /132/ ? [
                    $model->{ehv} && $model->{ehv} =~ /cap|33/i
                    ? qw(1 1 0 0 0 0 0 0)
                    : qw(0 1 0 0 0 0 0 0)
                  ]
                  : /GSP/i ? [qw(1 0 0 0 0 0 0 0)]
                  : [ map { 0 } 1 .. 8 ]
            } @{ $demandEndUsers->{list} }
          ]
        : $model->{standing} =~ /reduc/i ? [
            map {
                    /unmeter|generat/i ? [ map { 0 } 1 .. 8 ]
                  : /do not do this LV sub/i ? [qw(0 0 0 0 .2 1 1 0)]
                  : /LV/i                    ? [qw(0 0 0 0 0 .2 1 1)]
                  : /do not do this HV sub/i ? [qw(0 0 .2 1 1 0 0 0)]
                  : /HV/i                    ? [qw(0 0 0 .2 1 1 0 0)]
                  : /33kV sub/i              ? [
                    $model->{ehv} && $model->{ehv} =~ /cap|33/i
                    ? qw(1 1 1 0 0 0 0 0)
                    : qw(0 .2 1 0 0 0 0 0)
                  ]
                  : /33/ ? [
                    $model->{ehv} && $model->{ehv} =~ /cap|33/i
                    ? qw(1 1 1 1 0 0 0 0)
                    : qw(0 .2 1 1 0 0 0 0)
                  ]
                  : /132/  ? [qw(1 1 0 0 0 0 0 0)]
                  : /GSP/i ? [qw(1 0 0 0 0 0 0 0)]
                  : [ map { 0 } 1 .. 8 ]
            } @{ $demandEndUsers->{list} }
          ]
        : $model->{standing} =~ /edf/i ? [
            map {
                    /unmeter|generat/i ? [ map { 0 } 1 .. 8 ]
                  : /do not do this LV sub/i ? [qw(0 0 0 0 0 .5 1 0)]
                  : /LV/i                    ? [qw(0 0 0 0 0 0 .5 1)]
                  : /do not do this HV sub/i ? [qw(0 0 0 .5 1 0 0 0)]
                  : /HV/i                    ? [qw(0 0 0 0 .5 1 0 0)]
                  : /33kV sub/i              ? [
                    $model->{ehv} && $model->{ehv} =~ /cap|33/i
                    ? qw(1 1 1 0 0 0 0 0)
                    : qw(0 0 1 0 0 0 0 0)
                  ]
                  : /33/ ? [
                    $model->{ehv} && $model->{ehv} =~ /cap|33/i
                    ? qw(1 1 1 1 0 0 0 0)
                    : qw(0 0 .5 1 0 0 0 0)
                  ]
                  : /132/ ? [
                    $model->{ehv} && $model->{ehv} =~ /cap|33/i
                    ? qw(1 1 0 0 0 0 0 0)
                    : qw(.5 1 0 0 0 0 0 0)
                  ]
                  : /GSP/i ? [qw(1 0 0 0 0 0 0 0)]
                  : [ map { 0 } 1 .. 8 ]
            } @{ $demandEndUsers->{list} }
          ]
        : $model->{standing} =~ /g3/i ? [
            map {
                    /unmeter|generat/i ? [ map { 0 } 1 .. 8 ]
                  : /do not do this LV sub/i ? [qw(0 0 0 0 0 1 1 0)]
                  : /LV/i                    ? [qw(0 0 0 0 0 0 1 1)]
                  : /do not do this HV sub/i ? [qw(0 0 0 1 1 0 0 0)]
                  : /HV/i                    ? [qw(0 0 0 0 1 1 0 0)]
                  : /33kV sub/i              ? [
                    $model->{ehv} && $model->{ehv} =~ /cap|33/i
                    ? qw(1 1 1 0 0 0 0 0)
                    : qw(0 0 1 0 0 0 0 0)
                  ]
                  : /33/ ? [
                    $model->{ehv} && $model->{ehv} =~ /cap|33/i
                    ? qw(1 1 1 1 0 0 0 0)
                    : qw(0 0 1 1 0 0 0 0)
                  ]
                  : /132/  ? [qw(1 1 0 0 0 0 0 0)]
                  : /GSP/i ? [qw(1 0 0 0 0 0 0 0)]
                  : [ map { 0 } 1 .. 8 ]
            } @{ $demandEndUsers->{list} }
          ]
        : [
            map {
                [ map { 0 } 1 .. 8 ]
            } @{ $demandEndUsers->{list} }
        ],
        name  => 'Standing charges factors',
        lines => <<'EOL',
The standing charges factor for each end user and network level is the extent to which the costs of that network level
are applied to tariffs relating to that user using capacity charges/credits rather than unit charges/credits.
EOL
    );

    if ($rerouteing13211) {

        if ( $model->{standing} =~ /132/ ) {

            push @{ $model->{optionLines} },
              'Put some 132kV costs into HV capacity charges';

            my $scf13211 = Constant(
                name => 'Standing charges factors for 132kV/HV',
                rows => $demandEndUsers,
                cols => Labelset( list => [ $drmExitLevels->{list}[5] ] ),
                data => [
                    [
                        map {
                                /unmeter|generat/i ? 0
                              : /HV sub/i          ? 0
                              : /HV/i              ? 1
                              :                      0;
                        } @{ $demandEndUsers->{list} }
                    ]
                ]
            );

            my $scf132 = Arithmetic(
                name       => 'Adjusted standing charges factors for 132kV',
                cols       => Labelset( list => [ $drmExitLevels->{list}[1] ] ),
                rows       => $demandEndUsers,
                arithmetic => '=IV4+0.2*IV1*IV2',
                arguments  => {
                    IV1 => $rerouteing13211,
                    IV2 => $scf13211,
                    IV4 => $standingFactors
                }
            );

            Columnset(
                name    => 'Pre-processing of data for standing charge factors',
                columns => [ $standingFactors, $scf13211, $scf132 ]
            );

            my $scf132sub;

            if (
                my @tariffs132sub =
                grep { /^hv sub.*132/i } @{ $demandEndUsers->{list} }
              )
            {
                $scf132sub = Constant(
                    name =>
                      'Standing charge factors for 132kV/HV substation tariffs',
                    rows => Labelset(
                        name => '132kV/HV substation tariffs',
                        list => \@tariffs132sub
                    ),
                    cols => $drmExitLevels,
                    data => [
                        [ map { 0 } @tariffs132sub ],
                        !$model->{standing} || $model->{standing} =~ /sub/i
                        ? [ map { 1 } @tariffs132sub ]
                        : $model->{standing} =~ /nhh/
                        ? [ map { 0.2 } @tariffs132sub ]
                        : [ map { 0 } @tariffs132sub ],
                        [ map   { 0 } @tariffs132sub ],
                        [ map   { 0 } @tariffs132sub ],
                        [ map   { 0 } @tariffs132sub ],
                        [ map   { 1 } @tariffs132sub ],
                        [ map   { 0 } @tariffs132sub ],
                        [ map   { 0 } @tariffs132sub ],
                        [ map   { 0 } @tariffs132sub ],
                    ]
                );
            }

            $standingFactors = Stack(
                name    => 'Standing charges factors adapted to use 132kV/HV',
                rows    => $demandEndUsers,
                cols    => $drmExitLevels,
                sources => [
                    $scf132sub ? $scf132sub : (), $scf13211,
                    $scf132, $standingFactors,
                ]
            );
        }

        else
        { # bodge 132kV/HV in without allowing any 132kV costs into HV standing charges

            $standingFactors->{cols} = $drmExitLevels;

            $standingFactors->{data} =
              [ map { [ @{$_}[ 0 .. 3, 4, 4 .. 7 ] ]; }
                  @{ $standingFactors->{data} } ];

        }

    }

    push @{ $model->{forecastAml} }, $standingFactors;

    push @{ $model->{optionLines} },
      'Ignore standing charges in operating expenditure allocation'
      if $model->{opAllocSml};

    push @{ $model->{optionLines} },
      'Do not calculate diversity allowance for LV circuits'
      if $model->{useLvAml};

    my $lvCircuitLevel =
      Labelset(
        list => [ grep { /lv circuit/i } @{ $drmExitLevels->{list} } ] );

    my ( $diversityDemandTariffs, $diversityLevels ) =
      !$model->{opAllocSml}
      ? (
        Labelset(
            name => 'Demand tariffs, except related MPAN',
            $demandTariffsByEndUser->{groups} ? 'groups' : 'list' => [
                grep { !/(additional|related) mpan/i } @{
                         $demandTariffsByEndUser->{groups}
                      || $demandTariffsByEndUser->{list}
                }
            ]
        ),
        $drmExitLevels
      )
      : (
        Labelset(
            name => 'LV network demand tariffs, except related MPAN',
            $demandTariffsByEndUser->{groups} ? 'groups' : 'list' => [
                grep { /LV/i && !/LV sub/i; } @{
                         $demandTariffsByEndUser->{groups}
                      || $demandTariffsByEndUser->{list}
                }
            ]
        ),
        $lvCircuitLevel
      );

    my $forecastAmlCapacity;

    unless ( $model->{opAllocSml} && $model->{useLvAml} ) {
        my $demandTariffsCapacity = Labelset(
            name => 'Tariffs with agreed capacity',
            $diversityDemandTariffs->{groups} ? 'groups' : 'list' => [
                grep { $componentMap->{$_}{'Capacity charge p/kVA/day'} } @{
                         $diversityDemandTariffs->{groups}
                      || $diversityDemandTariffs->{list}
                }
            ]
        );

        push @{ $model->{forecastAml} }, $forecastAmlCapacity = Arithmetic(
            name => Label(
                    'Capacity-based contributions to chargeable aggregate '
                  . 'maximum load by network level (kW)'
            ),
            arithmetic => '=IV1*IV2*IV4*IV5',
            cols       => $diversityLevels,
            rows       => $demandTariffsCapacity,
            arguments  => {
                IV1 => $volumeData->{'Capacity charge p/kVA/day'},
                IV2 => $powerFactorInModel,
                IV4 => $standingFactors,
                IV5 => $lineLossFactors
            },
            defaultFormat => '0softnz',
        ) if $volumeData->{'Capacity charge p/kVA/day'};

        if ( $model->{spareCap} ) {
            if ( $model->{spareCap} =~ /first/i ) {
                $model->{spareCap} = Arithmetic(
                    name => 'Deemed spare capacity scaling factor (kVA/kW)',
                    rows =>
                      Labelset( list => [ $demandTariffsCapacity->{list}[0] ] ),
                    arithmetic => '=IV1/IV2*IV3*24*IV4/1000',
                    arguments  => {
                        IV1 => $volumeData->{'Capacity charge p/kVA/day'},
                        IV2 => $unitsInYear,
                        IV3 => $loadFactors,
                        IV4 => $daysInYear,
                    },
                );
            }
            else {
                my $demandTariffsSpareCap =
                    $model->{spareCap} =~ /all/i
                  ? $demandTariffsCapacity
                  : Labelset(
                    list => [
                        grep { /^LV/i && !/^LV Sub/i; }
                          @{ $demandTariffsCapacity->{list} }
                    ]
                  );
                my $cap = Stack(
                    rows    => $demandTariffsSpareCap,
                    sources => [ $volumeData->{'Capacity charge p/kVA/day'} ],
                );
                my $md = Arithmetic(
                    name       => 'Deemed maximum demand (kW)',
                    rows       => $demandTariffsSpareCap,
                    arithmetic => '=IV2*1000/IV3/24/IV4',
                    arguments  => {
                        IV2 => $unitsInYear,
                        IV3 => $loadFactors,
                        IV4 => $daysInYear,
                    },
                );
                Columnset(
                    name    => 'Apparent spare capacity in HH tariffs',
                    columns => [ $cap, $md ]
                );
                $model->{spareCap} = Arithmetic(
                    name => 'Deemed spare capacity scaling factor (kVA/kW)',
                    arithmetic => '=SUM(IV1_IV2)/SUM(IV3_IV4)',
                    arguments  => {
                        IV1_IV2 => $cap,
                        IV3_IV4 => $md,
                    },
                );
            }
        }
    }

    push @{ $model->{forecastAml} }, my $forecastAmlUnits = Arithmetic(
        name => Label(
                'Unit-based contributions to '
              . 'chargeable aggregate '
              . 'maximum load (kW)'
        ),
        arithmetic => '=IV1/IV2'
          . ( $model->{spareCap} ? '*IV81*IV82' : '' )
          . '*IV4*IV5/(24*IV9)*1000',
        cols      => $diversityLevels,
        rows      => $demandTariffsByEndUser,
        arguments => {
            IV1 => $unitsInYear,
            IV2 => $loadFactors,
            IV4 => $standingFactors,
            IV9 => $daysInYear,
            IV5 => $lineLossFactors,
            $model->{spareCap}
            ? (
                IV81 => $model->{spareCap},
                IV82 => $powerFactorInModel,
              )
            : (),
        },
        defaultFormat => '0softnz',
    );

    unless ( $model->{opAllocSml} && $model->{useLvAml} ) {

        push @{ $model->{forecastAml} }, my $chargeableAml = GroupBy(
            name   => 'Forecast chargeable aggregate maximum load (kW)',
            rows   => 0,
            cols   => $diversityLevels,
            source => Stack(
                name => 'Contributions to aggregate '
                  . 'maximum load by network level (kW)',
                rows    => $diversityDemandTariffs,
                cols    => $diversityLevels,
                sources => [
                    $forecastAmlCapacity ? $forecastAmlCapacity : (),
                    $forecastAmlUnits
                ]
            ),
            defaultFormat => '0softnz',
        );

        my $chargeableSml = GroupBy(
            name =>
              'Forecast simultaneous load replaced by standing charge (kW)',
            source => Arithmetic(
                name => 'Forecast simultaneous load subject '
                  . 'to standing charge factors (kW)',
                rows       => $demandTariffsByEndUser,
                cols       => $diversityLevels,
                arithmetic => '=IV1*IV2',
                arguments  => {
                    IV2 => $standingFactors,
                    IV1 => $forecastSml->{source},
                }
            ),
            cols          => $diversityLevels,
            rows          => 0,
            defaultFormat => '0softnz',
        );

=head Do not do this

            $diversityAllowances = Arithmetic(
                name          => 'Calculated diversity allowances',
                arithmetic    => '=IF(IV3>0,IV1/IV2-1,IV4)',
                defaultFormat => '%softnz',
                arguments     => {
                    IV1 => $chargeableAml,
                    IV2 => $chargeableSml,
                    IV3 => $chargeableSml,
                    IV4 => $diversityAllowances
                }
            );

=cut

        if ($rerouteing13211) {

            my $rerouteingMap = Constant(
                name => 'Network level mapping for diversity allowances',
                defaultFormat => '%connz',
                rows          => $coreExitLevels,
                cols          => $drmExitLevels,
                data          => [
                    [qw(1 0 0 0 0 0 0 0)], [qw(0 1 0 0 0 0 0 0)],
                    [qw(0 0 1 0 0 0 0 0)], [qw(0 0 0 1 0 0 0 0)],
                    [qw(0 0 0 0 1 0 0 0)], [qw(0 0 1 0 0 0 0 0)],
                    [qw(0 0 0 0 0 1 0 0)], [qw(0 0 0 0 0 0 1 0)],
                    [qw(0 0 0 0 0 0 0 1)],
                ]
            );

            $diversityAllowances = SumProduct(
                name          => 'Diversity allowances including 132kV/HV',
                vector        => $rerouteingMap,
                matrix        => $diversityAllowances,
                defaultFormat => '%softnz'
            );

        }

        push @{ $model->{forecastAml} }, $diversityAllowances = Stack(
            name => 'Diversity allowances (including calculated LV value)',
            defaultFormat => '%copynz',
            cols          => $drmExitLevels,
            rows          => 0,
            sources       => [
                Arithmetic(
                    name          => 'Calculated LV diversity allowance',
                    arithmetic    => '=IV1/IV2-1',
                    cols          => $lvCircuitLevel,
                    defaultFormat => '%softnz',
                    arguments =>
                      { IV1 => $chargeableAml, IV2 => $chargeableSml }
                ),
                $diversityAllowances
            ]
        ) unless $model->{useLvAml};

        push @{ $model->{forecastAml} }, $forecastSml = Arithmetic(
            name => 'Forecast simultaneous maximum load (kW)'
              . ' adjusted for standing charges',
            arithmetic    => '=IV3-IV2+IV1/(1+IV4)',
            defaultFormat => '0softnz',
            arguments     => {
                IV1 => $chargeableAml,
                IV2 => $chargeableSml,
                IV3 => $forecastSml,
                IV4 => $diversityAllowances
            }
        ) unless $model->{opAllocSml};

    }

    $standingFactors, $forecastSml, $forecastAmlUnits, $diversityAllowances;

}

1;
