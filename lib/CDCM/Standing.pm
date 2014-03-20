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

sub standingCharges {

    my (
        $model,                    $standingFactors,
        $drmLevels,                $drmExitLevels,
        $operatingDrmExitLevels,   $chargingDrmExitLevels,
        $demandEndUsers,           $demandTariffsByEndUser,
        $standingForFixedEndUsers, $standingForFixedTariffsByEndUser,
        $loadFactors,              $unitsByEndUser,
        $volumesByEndUser,         $unitsInYear,
        $volumeData,               $modelCostToAml,
        $modelCostToSml,           $operatingCostToSml,
        $costToSml,                $diversityAllowances,
        $lineLossFactors,          $proportionCoveredByContributions,
        $powerFactorInModel,       $daysInYear,
        $forecastAmlUnits,         $componentMap,
        $paygUnitYardstick,        @paygUnitRates
    ) = @_;

    push @{ $model->{standingResults} },
      my $costToAml =
      $model->{useLvAml}
      ? Stack(
        name    => 'All costs based on aggregate maximum load (£/kW/year)',
        cols    => $chargingDrmExitLevels,
        sources => [
              $model->{useLvAml}
            ? $modelCostToAml
            : Arithmetic(
                name => 'Network model annuities based'
                  . ' on aggregate maximum load (£/kW/year)',
                rows       => 0,
                cols       => $drmLevels,
                arithmetic => '=IV1/(1+IV2)',
                arguments  => {
                    IV1 => $modelCostToSml,
                    IV2 => $diversityAllowances
                }
            ),
            Arithmetic(
                name => 'Network operating costs based'
                  . ' on aggregate maximum load (£/kW/year)',
                rows       => 0,
                cols       => $operatingDrmExitLevels,
                arithmetic => '=IV1/(1+IV2)',
                arguments  => {
                    IV1 => $operatingCostToSml,
                    IV2 => $diversityAllowances
                }
            ),
            $costToSml
        ]
      )
      : Arithmetic(
        name       => 'Costs based on aggregate maximum load (£/kW/year)',
        rows       => 0,
        cols       => $chargingDrmExitLevels,
        arithmetic => '=IV1/(1+IV2)',
        arguments  => {
            IV1 => $costToSml,
            IV2 => $diversityAllowances
        }
      );

    my $capacityCharges = GroupBy
      name   => 'Capacity charge p/kVA/day',
      rows   => $demandTariffsByEndUser,
      cols   => 0,
      source => Arithmetic(
        name  => 'Capacity elements p/kVA/day',
        rows  => $demandTariffsByEndUser,
        cols  => $chargingDrmExitLevels,
        lines => 'This calculation uses aggregate '
          . 'maximum load and no coincidence factor.',
        arithmetic => '=100*IV5*IV1*IV2*IV3/IV4*(1-IV6)',
        arguments  => {
            IV1 => $lineLossFactors,
            IV6 => $proportionCoveredByContributions,
            IV2 => $costToAml,
            IV3 => $powerFactorInModel,
            IV4 => $daysInYear,
            IV5 => $standingFactors,
        },
        defaultFormat => '0.000softnz'
      );

    my $unauthorisedDemandCharges =
      $model->{unauth} && $model->{unauth} =~ /day/i
      ? GroupBy
      name   => 'Exceeded capacity charge p/kVA/day',
      rows   => $demandTariffsByEndUser,
      cols   => 0,
      source => Arithmetic(
        name       => 'Exceeded capacity charge elements p/kVA/day',
        rows       => $demandTariffsByEndUser,
        cols       => $chargingDrmExitLevels,
        arithmetic => $model->{unauth} =~ /same/i
        ? '=100*IV5*IV1*IV2*IV3/IV4*(1-IV6)'
        : '=100*IV5*IV1*IV2*IV3/IV4',
        arguments => {
            IV1 => $lineLossFactors,
            IV6 => $proportionCoveredByContributions,
            IV2 => $costToAml,
            IV3 => $powerFactorInModel,
            IV4 => $daysInYear,
            IV5 => $standingFactors,
        },
        defaultFormat => '0.000softnz'
      )
      : GroupBy
      name   => 'Unauthorised demand charge p/kVAh',
      rows   => $demandTariffsByEndUser,
      cols   => 0,
      source => Arithmetic(
        name       => 'Unauthorised demand charge elements p/kVAh',
        rows       => $demandTariffsByEndUser,
        cols       => $chargingDrmExitLevels,
        arithmetic => '=100*IV5*IV1*IV2*IV3/IV4',
        arguments  => {
            IV1 => $lineLossFactors,
            IV6 => $proportionCoveredByContributions,
            IV2 => $costToAml,
            IV3 => $powerFactorInModel,
            IV4 => Dataset(
                name       => 'Unauthorised demand maximum annual hours',
                validation => {
                    validate => 'decimal',
                    criteria => '>',
                    value    => 0,
                },
                appendTo => $model->{inputTables},
                dataset  => $model->{dataset},
                number   => 1089,
                data     => [40]
            ),
            IV5 => $standingFactors,
        },
        defaultFormat => '0.000softnz'
      ) if $model->{unauth};

    my $maxKvaByEndUser;

    # The lvCosts and fixedCap options interact and
    # do not behave in a particularly natural way.

    if ( $model->{fixedCap} && $model->{fixedCap} =~ /group|1-4/i ) {

    # get $maxKvaByEndUser from tariff grouping - to apply to all network levels

        my $tariffGroupset = Labelset(
            list => [
                $model->{fixedCap} =~ /1-4/
                ? (
                    'LV domestic and small non-domestic tariffs',
                    'LV medium non-domestic tariffs'
                  )
                : (
                    'LV domestic tariffs',
                    'LV network non-domestic aggregated tariffs'
                ),
                'LV substation aggregated tariffs',
                'HV network aggregated tariffs',
            ]
        );

        my $mapping = Constant(
            name          => 'Mapping of tariffs to tariff groups',
            defaultFormat => '0connz',
            rows          => $standingForFixedEndUsers,
            cols          => $tariffGroupset,
            data          => $model->{fixedCap} =~ /1-4/
            ? [
                map {
                        /(additional|related) mpan/i ? [qw(0 0 0 0)]
                      : /domestic|1p|single/i && !/non.?dom/i ? [qw(1 0 0 0)]
                      : /small/i  ? [qw(1 0 0 0)]
                      : /lv sub/i ? [qw(0 0 1 0)]
                      : /hv/i     ? [qw(0 0 0 1)]
                      :             [qw(0 1 0 0)];
                } @{ $standingForFixedEndUsers->{list} }
              ]
            : [
                map {
                        /(additional|related) mpan/i ? [qw(0 0 0 0)]
                      : /domestic|1p|single/i && !/non.?dom/i ? [qw(1 0 0 0)]
                      : /lv sub/i ? [qw(0 0 1 0)]
                      : /hv/i     ? [qw(0 0 0 1)]
                      :             [qw(0 1 0 0)];
                } @{ $standingForFixedEndUsers->{list} }
            ],
            byrow => 1,
        );

        my $numerator = SumProduct(
            matrix        => $mapping,
            name          => 'Aggregate capacity (kW)',
            defaultFormat => '0softnz',
            vector        => Arithmetic(
                name => Label(
                        'Unit-based contributions to aggregate '
                      . 'maximum load (kW)'
                ),
                arithmetic => '=IV1/IV2/(24*IV9)*1000',
                rows       => $standingForFixedEndUsers,
                arguments  => {
                    IV1 => $unitsByEndUser,
                    IV2 => $loadFactors,
                    IV9 => $daysInYear,
                },
                defaultFormat => '0softnz',
            )
        );

        my $denominator = SumProduct(
            matrix => $mapping,
            name   => 'Aggregate number of users charged '
              . 'for capacity on an exit point basis',
            defaultFormat => '0softnz',
            vector        => Stack(
                rows    => $standingForFixedEndUsers,
                sources => [ $volumeData->{'Fixed charge p/MPAN/day'} ]
            ),
        );

        my $maxKvaAverageLv = Arithmetic(
            name       => 'Average maximum kVA by exit point',
            arithmetic => '=IF(IV5,IV1/IV2'
              . ( $model->{spareCap} ? '*IV4' : '/IV4' ) . ',0)',
            arguments => {
                IV1 => $numerator,
                IV2 => $denominator,
                IV5 => $denominator,
                IV4 => $model->{spareCap} || $powerFactorInModel,
            }
        );

        Columnset(
            name => 'Capacity use for tariffs charged '
              . 'for capacity on an exit point basis',
            columns =>
              [ map { $maxKvaAverageLv->{arguments}{$_}{vector} } qw(IV1 IV2) ]
        );

        $maxKvaByEndUser = SumProduct(
            name   => 'Deemed average maximum kVA for each tariff',
            matrix => $mapping,
            vector => $maxKvaAverageLv
        );

    }

    else {    # tariff by tariff
        $maxKvaByEndUser =
          $model->{spareCap}
          ? Arithmetic(
            name => Label(
                'Average maximum kVA/MPAN',
                'Average maximum kVA/MPAN by end user class,'
                  . ' for user classes without an agreed import capacity'
            ),
            rows          => $standingForFixedEndUsers,
            arithmetic    => '=IF(IV6>0,IV2/IV3*IV4/IV1/(24*IV5)*1000,0)',
            defaultFormat => '0.000softnz',
            arguments     => {
                IV1 => $loadFactors,
                IV2 => $unitsByEndUser,
                IV6 => $volumesByEndUser->{'Fixed charge p/MPAN/day'},
                IV3 => $volumesByEndUser->{'Fixed charge p/MPAN/day'},
                IV4 => $model->{spareCap},
                IV5 => $daysInYear
            }
          )
          : Arithmetic(
            name => Label(
                'Average maximum kVA/MPAN',
                'Average maximum kVA/MPAN by end user class,'
                  . ' for user classes without an agreed import capacity'
            ),
            rows          => $standingForFixedEndUsers,
            arithmetic    => '=IF(IV6>0,IV2/IV3/IV4/IV1/(24*IV5)*1000,0)',
            defaultFormat => '0.000softnz',
            arguments     => {
                IV1 => $loadFactors,
                IV2 => $unitsByEndUser,
                IV6 => $volumesByEndUser->{'Fixed charge p/MPAN/day'},
                IV3 => $volumesByEndUser->{'Fixed charge p/MPAN/day'},
                IV4 => $powerFactorInModel,
                IV5 => $daysInYear
            }
          );
    }

    push @{ $model->{standingNhh} },
      my $capacityUserElements = Arithmetic(
        name => 'Capacity-driven fixed charge elements'
          . ' from standing charges factors p/MPAN/day',
        arithmetic => '=IV1*IV3',
        rows       => $standingForFixedTariffsByEndUser,
        arguments  => {
            IV3 => $maxKvaByEndUser,
            IV1 => $capacityCharges->{source},
        },
        defaultFormat => '0.000softnz'
      );

    if (   $model->{fixedCap}
        || $model->{lvCosts} && $model->{lvCosts} =~ /cap/i )
    {

        # no special grouping for LV circuits

    }

    elsif ( $model->{lvCosts} && $model->{lvCosts} =~ /group/i ) {

        # several tariff groups for LV circuit costs

        push @{ $model->{optionLines} },
          'LV circuit costs by exit point, separately for three tariff groups';

        my $lvStandingForFixedTariffsByEndUser = Labelset(
            name => 'LV demand tariffs charged by exit point',
            $standingForFixedTariffsByEndUser->{groups}
            ? 'groups'
            : 'list' => [
                grep { /LV/i and !/LV sub/i } @{
                         $standingForFixedTariffsByEndUser->{groups}
                      || $standingForFixedTariffsByEndUser->{list}
                }
            ]
        );

        my $tariffGroupset = Labelset(
            list => [
                'LV domestic tariffs',
                'LV non-domestic aggregated tariffs',
                'LV non-domestic CT tariffs',
            ]
        );

        my $mapping = Constant(
            name => 'Mapping of tariffs to tariff groups for LV circuit costs',
            defaultFormat => '0connz',
            rows          => $lvStandingForFixedTariffsByEndUser,
            cols          => $tariffGroupset,
            data          => [
                map {
                        /(additional|related) mpan/i ? [qw(0 0 0)]
                      : /domestic|1p|single/i && !/non.?dom/i ? [qw(1 0 0)]
                      : /wc|small/i ? [qw(0 1 0)]
                      :               [qw(0 0 1)];
                } @{ $lvStandingForFixedTariffsByEndUser->{list} }
            ],
            byrow => 1,
        );

        if ( $model->{lvCosts} =~ /dnd/ ) {

            $tariffGroupset =
              Labelset( list => [ 'LV domestic', 'LV non domestic', ] );

            $mapping = Constant(
                name =>
                  'Mapping of tariffs to tariff groups for LV circuit costs',
                defaultFormat => '0connz',
                rows          => $lvStandingForFixedTariffsByEndUser,
                cols          => $tariffGroupset,
                data          => [
                    map {
                            /(additional|related) mpan/i ? [qw(0 0 0)]
                          : /domestic|1p|single/i
                          && !/non.?dom/i ? [qw(1 0)]
                          : [qw(0 1)];
                    } @{ $lvStandingForFixedTariffsByEndUser->{list} }
                ],
                byrow => 1,
            );

        }

        my $maxKvaAverageLv = Arithmetic(
            name => 'Average maximum kVA of tariffs '
              . 'charged on an exit point basis for LV circuits',
            arithmetic => '=IV1/IV2' . ( $model->{spareCap} ? '*IV4' : '/IV4' ),
            arguments => {
                IV1 => SumProduct(
                    matrix => $mapping,
                    name   => 'Aggregate capacity of tariffs charged '
                      . 'charged for LV circuits on an exit point basis (kW)',
                    defaultFormat => '0softnz',
                    vector        => Arithmetic(
                        name => Label(
                                'Unit-based contributions to aggregate '
                              . 'maximum load by network level (kW)'
                        ),
                        arithmetic => '=IV1/IV2/(24*IV9)*1000',
                        rows       => $lvStandingForFixedTariffsByEndUser,
                        arguments  => {
                            IV1 => $unitsInYear,
                            IV2 => $loadFactors,
                            IV9 => $daysInYear,
                        },
                        defaultFormat => '0softnz',
                    )
                ),
                IV2 => SumProduct(
                    matrix => $mapping,
                    name   => 'Aggregate number of users charged '
                      . 'for LV circuits on an exit point basis',
                    defaultFormat => '0softnz',
                    vector        => Stack(
                        rows    => $lvStandingForFixedTariffsByEndUser,
                        sources => [ $volumeData->{'Fixed charge p/MPAN/day'} ]
                    ),
                ),
                IV4 => $model->{spareCap} || $powerFactorInModel,
            }
        );

        push @{ $model->{standingNhh} },
          Columnset(
            name => 'Capacity use for tariffs charged '
              . 'for LV circuits on an exit point basis',
            columns =>
              [ map { $maxKvaAverageLv->{arguments}{$_}{vector} } qw(IV1 IV2) ]
          );

        push @{ $model->{standingNhh} },
          $capacityUserElements = Stack(
            name => 'Fixed charge elements from '
              . 'standing charges factors p/MPAN/day',
            cols    => $chargingDrmExitLevels,
            rows    => $standingForFixedTariffsByEndUser,
            sources => [
                Arithmetic(
                    name => 'LV fixed charge elements from '
                      . 'standing charges factors p/MPAN/day',
                    rows => $lvStandingForFixedTariffsByEndUser,
                    cols => Labelset(
                        list => [
                            grep { /lv circuit/i }
                              @{ $chargingDrmExitLevels->{list} }
                        ]
                    ),
                    arithmetic => '=IV1*IV3',
                    arguments  => {
                        IV3 => SumProduct(
                            name =>
                              'Deemed average maximum kVA for each tariff',
                            matrix => $mapping,
                            vector => $maxKvaAverageLv
                        ),
                        IV1 => $capacityCharges->{source},
                    },
                    defaultFormat => '0.000softnz'
                ),
                $capacityUserElements
            ],
            defaultFormat => '0.000copynz'
          );

    }

    else {

# model 100 approach: one group for profile classes 1-4 (or optionally for all NHH)

        my $allNhh = $model->{lvCosts} && $model->{lvCosts} =~ /nhh/i;

        push @{ $model->{optionLines} },
          'LV circuit costs by exit point for '
          . (
            $allNhh
            ? 'all NHH demand'
            : 'all small NHH demand'
          );

        my $lvStandingForFixedTariffs = Labelset(
            name => 'LV demand tariffs charged by exit point',
            $standingForFixedTariffsByEndUser->{groups}
            ? 'groups'
            : 'list' => [
                grep {
                          /LV/i
                      and !/LV sub/i
                      and $allNhh
                      || !/(profile|pc).*[5-8]|medium|\bCT\b/i
                  } @{
                    $standingForFixedTariffsByEndUser->{groups}
                      || $standingForFixedTariffsByEndUser->{list}
                  }
            ]
        );

        my $lvRelatedMpanTariffs = Labelset(
            name => 'LV related MPAN users without capacity charge',
            $lvStandingForFixedTariffs->{groups} ? 'groups' : 'list' => [
                grep { /(additional|related) mpan/i } @{
                         $lvStandingForFixedTariffs->{groups}
                      || $lvStandingForFixedTariffs->{list}
                }
            ]
        );

        my $lvCircuitLevel =
          Labelset(
            list => [ grep { /lv circuit/i } @{ $drmExitLevels->{list} } ] );

        my $lvRouteingFactors = Stack(
            name => 'Use of LV circuits by each tariff '
              . 'charged on an exit point basis',
            rows    => $lvStandingForFixedTariffs,
            cols    => $lvCircuitLevel,
            sources => [$lineLossFactors]
        );

        my $maxKvaAverageLv = Arithmetic(
            name => 'Average maximum kVA of tariffs '
              . 'charged on an exit point basis for LV circuits',
            arithmetic => '=IV1/IV2' . ( $model->{spareCap} ? '*IV4' : '/IV4' ),
            rows       => 0,
            cols       => 0,
            arguments  => {
                IV1 => SumProduct(
                    matrix => $lvRouteingFactors,
                    name   => 'Aggregate capacity of tariffs charged '
                      . 'charged for LV circuits on an exit point basis (kW)',
                    defaultFormat => '0softnz',
                    vector        => Arithmetic(
                        name => Label(
                                'Unit-based contributions to aggregate '
                              . 'maximum load by network level (kW)'
                        ),
                        arithmetic => '=IV1/IV2/(24*IV9)*1000',
                        rows       => $lvStandingForFixedTariffs,
                        arguments  => {
                            IV1 => $unitsInYear,
                            IV2 => $loadFactors,
                            IV9 => $daysInYear,
                        },
                        defaultFormat => '0softnz',
                    )
                ),
                IV2 => SumProduct(
                    matrix => $lvRouteingFactors,
                    name   => 'Aggregate number of users charged '
                      . 'for LV circuits on an exit point basis',
                    defaultFormat => '0softnz',
                    vector        => Stack(
                        name          => 'Relevant MPAN count',
                        defaultFormat => '0copynz',
                        rows          => $lvStandingForFixedTariffs,
                        sources       => [
                            Constant(
                                name => 'Zero for related MPANs',
                                rows => $lvRelatedMpanTariffs,
                                data => [
                                    [
                                        map { 0 }
                                          @{ $lvRelatedMpanTariffs->{list} }
                                    ]
                                ]
                            ),
                            $volumeData->{'Fixed charge p/MPAN/day'}
                        ]
                    )
                ),
                IV4 => $model->{spareCap} || $powerFactorInModel,
            }
        );

        push @{ $model->{standingNhh} },
          Columnset(
            name => 'Capacity use for tariffs charged '
              . 'for LV circuits on an exit point basis',
            columns => [
                $lvRouteingFactors,
                map { $maxKvaAverageLv->{arguments}{$_}{vector} } qw(IV1 IV2)
            ]
          ),
          Columnset(
            name => 'Aggregate data for tariffs charged '
              . 'for LV circuits on an exit point basis',
            columns => [
                @{ $maxKvaAverageLv->{arguments} }{qw(IV1 IV2)},
                $maxKvaAverageLv
            ]
          );

        my $lvCircuitLevels =
          Labelset( list =>
              [ grep { /lv circuit/i } @{ $chargingDrmExitLevels->{list} } ] );

        push @{ $model->{standingNhh} }, $capacityUserElements = Stack(
            name => 'Fixed charge elements from '
              . 'standing charges factors p/MPAN/day',
            cols    => $chargingDrmExitLevels,
            rows    => $standingForFixedTariffsByEndUser,
            sources => [
                Constant(
                    name => 'Zero for related MPANs',
                    rows => $lvRelatedMpanTariffs,
                    cols => $lvCircuitLevels,
                    data => [
                        map {
                            [ map { 0 } @{ $lvRelatedMpanTariffs->{list} } ]
                        } 1 .. 2
                    ]
                ),
                Arithmetic(
                    name => 'LV fixed charge elements from '
                      . 'standing charges factors p/MPAN/day',
                    arithmetic => '=IV1*IV3' . ( $model->{pcd} ? '' : '*IV5' ),
                    rows      => $lvStandingForFixedTariffs,
                    cols      => $lvCircuitLevels,
                    arguments => {
                        IV3 => $maxKvaAverageLv,
                        IV1 => $capacityCharges->{source},
                        $model->{pcd} ? () : ( IV5 => $lvRouteingFactors ),
                    },
                    defaultFormat => '0.000softnz'
                ),
                $capacityUserElements
            ],
            defaultFormat => '0.000copynz'
        );

    }

    my $capacityUser = GroupBy(
        name   => 'Fixed charge from standing charges factors p/MPAN/day',
        rows   => $standingForFixedTariffsByEndUser,
        cols   => 0,
        source => $capacityUserElements
    );

    push @{ $model->{standingResults} },
      $model->{showSums}
      ? Columnset(
        name    => 'Capacity charges from standing charges factors',
        columns => [ $capacityCharges->{source}, $capacityCharges ]
      )
      : $capacityCharges->{source};

    push @{ $model->{unauthorisedDemand} },
      $model->{showSums}
      ? Columnset(
        name => $model->{unauth} && $model->{unauth} =~ /day/i
        ? 'Exceeded capacity charges from standing charges factors'
        : 'Unauthorised demand charges from standing charges factors',
        columns =>
          [ $unauthorisedDemandCharges->{source}, $unauthorisedDemandCharges ]
      )
      : $unauthorisedDemandCharges->{source};

    push @{ $model->{standingNhh} },
      $model->{showSums}
      ? Columnset(
        name    => 'Fixed charges from standing charges factors',
        columns => [ $capacityUser->{source}, $capacityUser ]
      )
      : $capacityUser->{source};

# The unrestricted yardstick used in reactive power calculations even if $model->{alwaysUseRAG}

    my $unitYardstick = GroupBy(
        name   => 'Yardstick total p/kWh (taking account of standing charges)',
        rows   => $demandTariffsByEndUser,
        source => Arithmetic(
            rows => $demandTariffsByEndUser,
            name =>
              'Yardstick components p/kWh (taking account of standing charges)',
            arithmetic => '=(1-IV2)*IV1',
            arguments  => {
                IV1 => $paygUnitYardstick->{source},
                IV2 => $standingFactors,
            },
            defaultFormat => '0.000softnz'
        )
    );

    push @{ $model->{standingResults} },
      $model->{showSums}
      ? Columnset(
        name =>
          'Yardstick unit rate p/kWh (taking account of standing charges)',
        columns => [ $unitYardstick->{source}, $unitYardstick ]
      )
      : $unitYardstick->{source};

    my @unitRates = map {
        my $relevantTariffs = Labelset(
            name => 'Demand ' . ( 1 + $_ ) . '-rate tariffs',
            $paygUnitRates[$_]{rows}{groups} ? 'groups' : 'list' => [
                grep { !/gener/i } @{
                         $paygUnitRates[$_]{rows}{groups}
                      || $paygUnitRates[$_]{rows}{list}
                }
            ]
        );
        my $c = Arithmetic(
            name => 'Contributions to unit rate '
              . ( 1 + $_ )
              . ' p/kWh by network level (taking account of standing charges)',
            defaultFormat => '0.000softnz',
            arithmetic    => '=(1-IV2)*IV1',
            rows          => $relevantTariffs,
            arguments     => {
                IV1 => $paygUnitRates[$_]{source},
                IV2 => $standingFactors,
            }
        );
        my $a = GroupBy
          rows => $c->{rows},
          name => 'Unit rate '
          . ( 1 + $_ )
          . ' total p/kWh (taking account of standing charges)',
          source => $c;
        push @{ $model->{standingResults} },
          $model->{showSums}
          ? Columnset(
            name => 'Unit rate '
              . ( 1 + $_ )
              . ' (taking account of standing charges)',
            columns => [ $c, $a ]
          )
          : $c;
        $a;
    } 0 .. ( $model->{maxUnitRates} > 1 ? $model->{maxUnitRates} - 1 : -1 );

    $capacityCharges, $unauthorisedDemandCharges, $capacityUser, $unitYardstick,
      @unitRates;

}

1;
