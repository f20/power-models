package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others.
Copyright 2011-2014 Franck Latrémolière, Reckon LLP and others.

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
        $demandTariffsByEndUser, $standingForFixedTariffsByEndUser,
        $unitsInYear,            $loadFactors,
        $daysInYear,             $lineLossFactors,
        $diversityAllowances,    $componentMap,
        $volumeData,             $powerFactorInModel,
        $forecastSml,            $drmExitLevels,
        $coreExitLevels,         $rerouteing13211,
    ) = @_;

    push @{ $model->{optionLines} }, !$model->{standing}
      || $model->{standing} =~ /^sub/i
      ? 'Standing charges factors: '
      . '100/0/0 LV NHH, '
      . '100/100/20 network, '
      . '100/100/0 substation'
      : $model->{standing} =~ /nhh/i
      ? 'Standing charges factors: 100/0/0 for LV NHH, 100/100/20 for others'
      : $model->{standing} =~ /mig12/i
      ? 'Standing charges factors: 100/0/0 for PC1-4, 100/100/20 network, 100/100/0 substation'
      : $model->{standing} =~ /dndsub/i
      ? 'Standing charges factors: 100/0/0 for domestic, 100/100/20 non domestic, 100/100/0 substation'
      : $model->{standing} =~ /pc1-?4/i
      ? 'Standing charges factors: 100/0/0 for PC1-4, 100/100/20 for others'
      : $model->{standing} =~ /low/i
      ? 'Standing charges factors: 100/0/0 for everyone'
      : $model->{standing} =~ /reducsub/i
      ? 'Standing charges factors: 100/100/20 for network, 100/100/0 for substation'
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
        data  => !$model->{standing} || $model->{standing} =~ /^sub/i
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
                  : /HV sub/i ? [qw(0 0 0 1 1 0 0 0)]
                  : /HV/i     ? (
                    !$model->{standing}
                      || $model->{standing} !~ /par74/i
                      || $componentMap->{$_}{'Capacity charge p/kVA/day'}
                    ? [qw(0 0 0 .2 1 1 0 0)]
                    : [qw(0 0 0 0 0 1 0 0)]
                  )
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
                  : /132/ ? [
                    $model->{ehv} && $model->{ehv} =~ /cap/i
                    ? qw(1 1 0 0 0 0 0 0)
                    : qw(0 1 0 0 0 0 0 0)
                  ]
                  : /GSP/i ? [
                    $model->{ehv} && $model->{ehv} =~ /cap/i
                    ? qw(1 0 0 0 0 0 0 0)
                    : qw(0 0 0 0 0 0 0 0)
                  ]
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
        : $model->{standing} =~ /dndsub/i ? [
            map {
                    /unmeter|generat/i ? [ map { 0 } 1 .. 8 ]
                  : /LV Sub/i          ? [qw(0 0 0 0 0 1 1 0)]
                  : /LV/i              ? (
                    /non.?dom/i || !/domestic/i
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
        : $model->{standing} =~ /reducsub/i ? [
            map {
                    /unmeter|generat/i ? [ map { 0 } 1 .. 8 ]
                  : /LV sub/i          ? [qw(0 0 0 0 0 1 1 0)]
                  : /LV/i              ? [qw(0 0 0 0 0 .2 1 1)]
                  : /HV sub/i          ? [qw(0 0 0 1 1 0 0 0)]
                  : /HV/i              ? [qw(0 0 0 .2 1 1 0 0)]
                  : /33kV sub/i        ? [
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
                arithmetic => '=A4+0.2*A1*A2',
                arguments  => {
                    A1 => $rerouteing13211,
                    A2 => $scf13211,
                    A4 => $standingFactors
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
                        !$model->{standing} || $model->{standing} =~ /^sub/i
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

    my ( $diversityDemandTariffs, $diversityLevels ) = !$model->{opAllocSml}
      ? (
        Labelset(
            name => 'Demand tariffs with capacity or fixed charges',
            $demandTariffsByEndUser->{groups} ? 'groups' : 'list' => [
                grep {
                         $componentMap->{$_}{'Fixed charge p/MPAN/day'}
                      || $componentMap->{$_}{'Capacity charge p/kVA/day'}
                  } @{
                    $demandTariffsByEndUser->{groups}
                      || $demandTariffsByEndUser->{list}
                  }
            ]
        ),
        $drmExitLevels
      )
      : (
        Labelset(
            name => 'LV network demand tariffs',
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

        push @{ $model->{forecastAml} },
          $forecastAmlCapacity =
             $model->{unauth}
          && $model->{unauth} =~ /day/i && $model->{unauth} =~ /otex/i
          ? Arithmetic(
            name => Label(
                    'Capacity-based contributions to chargeable aggregate '
                  . 'maximum load by network level (kW)'
            ),
            arithmetic => '=(A1+A6)*A2*A4*A5',
            cols       => $diversityLevels,
            rows       => $demandTariffsCapacity,
            arguments  => {
                A1 => $volumeData->{'Capacity charge p/kVA/day'},
                A6 => $volumeData->{'Exceeded capacity charge p/kVA/day'},
                A2 => $powerFactorInModel,
                A4 => $standingFactors,
                A5 => $lineLossFactors,
            },
            defaultFormat => '0softnz',
          )
          : Arithmetic(
            name => Label(
                    'Capacity-based contributions to chargeable aggregate '
                  . 'maximum load by network level (kW)'
            ),
            arithmetic => '=A1*A2*A4*A5',
            cols       => $diversityLevels,
            rows       => $demandTariffsCapacity,
            arguments  => {
                A1 => $volumeData->{'Capacity charge p/kVA/day'},
                A2 => $powerFactorInModel,
                A4 => $standingFactors,
                A5 => $lineLossFactors,
            },
            defaultFormat => '0softnz',
          ) if $volumeData->{'Capacity charge p/kVA/day'};

        # Historical feature, largely superseded by impliedLoadFactors
        if ( $model->{aggCapFactor} ) {
            if ( $model->{aggCapFactor} =~ /first/i ) {
                $model->{aggCapFactor} = Arithmetic(
                    name =>
                      'Deemed supercustomer capacity scaling factor (kVA/kW)',
                    rows =>
                      Labelset( list => [ $demandTariffsCapacity->{list}[0] ] ),
                    arithmetic => '=A1/A2*A3*24*A4/1000',
                    arguments  => {
                        A1 => $volumeData->{'Capacity charge p/kVA/day'},
                        A2 => $unitsInYear,
                        A3 => $loadFactors,
                        A4 => $daysInYear,
                    },
                );
            }
            else {
                my $demandTariffsCapFactorSource =
                    $model->{aggCapFactor} =~ /all/i
                  ? $demandTariffsCapacity
                  : Labelset(
                    list => [
                        grep { /^LV/i && !/^LV Sub/i; }
                          @{ $demandTariffsCapacity->{list} }
                    ]
                  );
                my $cap = Stack(
                    rows    => $demandTariffsCapFactorSource,
                    sources => [ $volumeData->{'Capacity charge p/kVA/day'} ],
                );
                my $md = Arithmetic(
                    name          => 'Deemed maximum demand (kW)',
                    rows          => $demandTariffsCapFactorSource,
                    defaultFormat => '0soft',
                    arithmetic    => '=A2*1000/A3/24/A4',
                    arguments     => {
                        A2 => $unitsInYear,
                        A3 => $demandTariffsByEndUser->{groups}
                        ? Arithmetic(
                            name       => 'Load factor (reshaped)',
                            arithmetic => '=A1',
                            rows       => $demandTariffsByEndUser,
                            arguments  => { A1 => $loadFactors },
                          )
                        : $loadFactors,
                        A4 => $daysInYear,
                    },
                );
                Columnset(
                    name => 'Diversified maximum demand'
                      . ' and aggregate site-specific capacity',
                    columns => [ $cap, $md ]
                );
                $model->{aggCapFactor} = Arithmetic(
                    name =>
                      'Deemed supercustomer capacity scaling factor (kVA/kW)',
                    arithmetic => '=SUM(A1_A2)/SUM(A3_A4)',
                    arguments  => {
                        A1_A2 => $cap,
                        A3_A4 => $md,
                    },
                );
            }
        }
    }

    push @{ $model->{forecastAml} },
      my $forecastAmlUnits = Arithmetic(
        name => Label(
                'Unit-based contributions to '
              . 'chargeable aggregate '
              . 'maximum load (kW)'
        ),
        arithmetic => '=A2/A1'
          . ( $model->{aggCapFactor} ? '*A81*A82' : '' )
          . '*A4*A5/(24*A9)*1000',
        rows      => $standingForFixedTariffsByEndUser,
        cols      => $diversityLevels,
        arguments => {
            A2 => $unitsInYear,
            A1 => $loadFactors,
            A4 => $standingFactors,
            A9 => $daysInYear,
            A5 => $lineLossFactors,
            $model->{aggCapFactor}
            ? (
                A81 => $model->{aggCapFactor},
                A82 => $powerFactorInModel,
              )
            : (),
        },
        defaultFormat => '0softnz',
      );

    unless ( $model->{opAllocSml} && $model->{useLvAml} ) {

        push @{ $model->{forecastAml} },
          my $chargeableAml = GroupBy(
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
                ],
                defaultFormat => '0copynz',
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
                arithmetic => '=A1*A2',
                arguments  => {
                    A2 => $standingFactors,
                    A1 => $forecastSml->{source},
                }
            ),
            cols          => $diversityLevels,
            rows          => 0,
            defaultFormat => '0softnz',
        );

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

        push @{ $model->{forecastAml} },
          $diversityAllowances = Stack(
            name => 'Diversity allowances (including calculated LV value)',
            defaultFormat => '%copynz',
            cols          => $drmExitLevels,
            rows          => 0,
            sources       => [
                Arithmetic(
                    name          => 'Calculated LV diversity allowance',
                    arithmetic    => '=A1/A2-1',
                    cols          => $lvCircuitLevel,
                    defaultFormat => '%softnz',
                    arguments => { A1 => $chargeableAml, A2 => $chargeableSml }
                ),
                $diversityAllowances
            ]
          ) unless $model->{useLvAml};

        push @{ $model->{edcmTables} },
          Stack(
            name => 'Diversity allowance between level exit and GSP Group',
            defaultFormat => '0.000hard',
            number        => 1105,
            rows          => 0,
            cols          => Labelset(
                list => [
                    @{
                        (
                                 $diversityAllowances->{cols}
                              || $diversityAllowances->{rows}
                        )->{list}
                    }[ 0 .. 5 ]
                ]
            ),
            sources => [$diversityAllowances]
          ) if $model->{edcmTables};

        push @{ $model->{forecastAml} },
          $forecastSml = Arithmetic(
            name => 'Forecast simultaneous maximum load (kW)'
              . ' adjusted for standing charges',
            arithmetic    => '=A3-A2+A1/(1+A4)',
            defaultFormat => '0softnz',
            arguments     => {
                A1 => $chargeableAml,
                A2 => $chargeableSml,
                A3 => $forecastSml,
                A4 => $diversityAllowances
            }
          ) unless $model->{opAllocSml};

    }

    $standingFactors, $forecastSml, $forecastAmlUnits, $diversityAllowances;

}

sub impliedLoadFactors {

    my (
        $model,              $allEndUsers,
        $demandEndUsers,     $standingForFixedEndUsers,
        $componentMap,       $volumesByEndUser,
        $unitsByEndUser,     $daysInYear,
        $powerFactorInModel, $loadFactors,
    ) = @_;

    my $tariffGroupset = Labelset( list =>
          [ 'LV Network tariffs', 'LV Sub tariffs', 'HV Network tariffs', ] );

    my $endUsersToUse =
      $model->{impliedLoadFactors} =~ /cf|corr/i
      ? Labelset(
        name => 'End users with maximum import capacity charges',
        list => [
            grep { $componentMap->{$_}{'Capacity charge p/kVA/day'} }
              @{ $demandEndUsers->{list} }
        ],
      )
      : $allEndUsers;

    my $mapping1 = Constant(
        name => 'Mapping of users with capacity charges to tariff groups'
          . ' for deemed load factors',
        defaultFormat => '0connz',
        rows          => $endUsersToUse,
        cols          => $tariffGroupset,
        data          => [
            map {
                /gener/i || !$componentMap->{$_}{'Capacity charge p/kVA/day'}
                  ? [ 0, 0, 0 ]
                  : /^(> )?HV Sub/i ? [ 0, 0, 0 ]
                  : /^(> )?HV/i     ? [ 0, 0, 1 ]
                  : /^(> )?LV Sub/i ? [ 0, 1, 0 ]
                  : /^(> )?LV/i     ? [ 1, 0, 0 ]
                  :                   [ 0, 0, 0 ];
            } @{ $endUsersToUse->{list} }
        ],
        byrow => 1,
    );

    my $mapping2 = Constant(
        name => 'Mapping of users with fixed charges based'
          . ' on standing charges factors to tariff groups'
          . ' for deemed load factors',
        defaultFormat => '0connz',
        rows          => $standingForFixedEndUsers,
        cols          => $tariffGroupset,
        data          => [
            map {
                    /^(> )?HV Sub/i ? [ 0, 0, 0 ]
                  : /^(> )?HV/i     ? [ 0, 0, 1 ]
                  : /^(> )?LV Sub/i ? [ 0, 1, 0 ]
                  : /^(> )?LV/i     ? [ 1, 0, 0 ]
                  : [ 0, 0, 0 ]
            } @{ $standingForFixedEndUsers->{list} }
        ],
        byrow => 1,
    );

    $model->{impliedLoadFactors} =~ /cf|corr/i
      ?

      Arithmetic(
        name       => 'Implied average load factor',
        arithmetic => '=A1*A2',
        arguments  => {
            A1 => SumProduct(
                name   => 'Supercustomer load factor correction factor',
                matrix => $mapping2,
                vector => Arithmetic(
                    name => 'Average load factor correction factor'
                      . ' for tariff group',
                    arithmetic => '=A1/A3/A4',
                    arguments  => {
                        A1 => SumProduct(
                            defaultFormat => '0softnz',
                            name => 'Relevant diversified maximum demand'
                              . ' in tariff group (kW)',
                            matrix => $mapping1,
                            vector => Arithmetic(
                                name => 'Diversified maximum demand (kW)',
                                defaultFormat => '0softnz',
                                rows          => $endUsersToUse,
                                arithmetic    => '=A1/24/A2/A3*1000',
                                arguments     => {
                                    A1 => $unitsByEndUser,
                                    A2 => $daysInYear,
                                    A3 => $loadFactors,
                                },
                            ),
                        ),
                        A3 => SumProduct(
                            defaultFormat => '0softnz',
                            name          => 'Aggregate maximum import capacity'
                              . ' in tariff group (kVA)',
                            matrix => $mapping1,
                            vector => Stack(
                                rows    => $endUsersToUse,
                                sources => [
                                    $volumesByEndUser->{
                                        'Capacity charge p/kVA/day'}
                                ],
                            )
                        ),
                        A4 => $powerFactorInModel,
                    }
                ),
            ),
            A2 => $loadFactors,
        },
      )

      : SumProduct(
        name   => 'Implied average load factor',
        matrix => $mapping2,
        vector => Arithmetic(
            name       => 'Implied average load factor for each tariff group',
            arithmetic => '=A1/24/A2/A3/A4*1000',
            arguments  => {
                A1 => SumProduct(
                    defaultFormat => '0softnz',
                    name   => 'Relevant consumption in tariff group (MWh)',
                    matrix => $mapping1,
                    vector => $unitsByEndUser
                ),
                A2 => $daysInYear,
                A3 => SumProduct(
                    defaultFormat => '0softnz',
                    name =>
                      'Aggregate maximum import capacity in tariff group (kVA)',
                    matrix => $mapping1,
                    vector => $volumesByEndUser->{'Capacity charge p/kVA/day'}
                ),
                A4 => $powerFactorInModel,
            }
        ),
      );

}

sub impliedLoadFactorsInputBased {

    my (
        $model,              $allEndUsers,
        $demandEndUsers,     $standingForFixedEndUsers,
        $componentMap,       $volumesByEndUser,
        $unitsByEndUser,     $daysInYear,
        $powerFactorInModel, $loadFactors,
    ) = @_;

    my $inputLoadFactors = Dataset(
        name       => 'Load factor',
        rows       => $standingForFixedEndUsers,
        validation => {
            validate      => 'decimal',
            criteria      => 'between',
            minimum       => 0,
            maximum       => 1,
            input_title   => 'Load factor:',
            input_message => 'Load factor',
            error_message => 'The load factor'
              . ' must be between 0% and 100%.',
        },
        data => [ map { 0.5; } @{ $demandEndUsers->{list} } ],
    );

    my $inputCapacityFactors = Dataset(
        name       => 'Spare capacity multiplier',
        rows       => $standingForFixedEndUsers,
        validation => {
            validate      => 'decimal',
            criteria      => '>',
            value         => 0,
            input_title   => 'Multiplier:',
            input_message => 'Spare capacity multiplier',
            error_message => 'The spare capacity multiplier must be positive.',
        },
        data => [ map { 2; } @{ $demandEndUsers->{list} } ],
    );

    Columnset(
        name     => 'Capacity assumptions for tariffs without capacity charges',
        columns  => [ $inputLoadFactors, $inputCapacityFactors ],
        number   => 1042,
        appendTo => $model->{inputTables},
        dataset  => $model->{dataset},
    );

    my $inputOnly = $model->{impliedLoadFactors} =~ /inputonly/i;

    Arithmetic(
        name => 'Deemed site-specific load factor for fixed charge calculation',
        arguments => {
            A1 => $inputLoadFactors,
            A4 => $inputCapacityFactors,
            $inputOnly ? ()
            : (
                A3 => $inputLoadFactors,
                A2 => $inputCapacityFactors,
                A9 => $model->impliedLoadFactors(
                    $allEndUsers,              $demandEndUsers,
                    $standingForFixedEndUsers, $componentMap,
                    $volumesByEndUser,         $unitsByEndUser,
                    $daysInYear,               $powerFactorInModel,
                    $loadFactors,
                ),
            ),
        },
        arithmetic => $inputOnly ? '=A1/A4'
        : '=IF(ISERROR(A3/A2),A9,A1/A4)',
    );

}

1;
