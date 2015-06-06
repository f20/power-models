package CDCM;

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

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Shortcuts ':all';
use YAML;

sub volumes {

    my (
        $model,                 $allTariffsByEndUser, $allEndUsers,
        $nonExcludedComponents, $componentMap,        $unitsAdjustmentFactor
    ) = @_;

    my %componentVolumeName = (
        (
            map { ( "Unit rate $_ p/kWh", "Rate $_ units (MWh)" ) }
              1 .. $model->{maxUnitRates}
        ),
        split "\n",
        <<'EOL');
Fixed charge p/MPAN/day
MPANs
Capacity charge p/kVA/day
Import capacity (kVA)
Unauthorised demand charge p/kVAh
Unauthorised demand (MVAh)
Exceeded capacity charge p/kVA/day
Exceeded capacity (kVA)
Generation capacity rate p/kW/day
Generation capacity (kW)
Reactive power charge p/kVArh
Reactive power units (MVArh)
EOL

    # MPANs are included for all tariffs, even if the model does not need them

    my %componentVolumeData = map {
        my $component = $_;
        my @data;
        foreach ( $allTariffsByEndUser->indices ) {
            $data[$_] = 0
              if $component eq 'Fixed charge p/MPAN/day'
              || $componentMap->{ $allTariffsByEndUser->{list}[$_] }
              {$component};
        }
        $_ => \@data;
    } @$nonExcludedComponents;

    my %volumeData = map {
        $_ => Dataset(
            name => Label(
                $componentVolumeName{$_},
                "$componentVolumeName{$_} by tariff"
            ),
            rows       => $allTariffsByEndUser,
            validation => {
                validate      => 'decimal',
                criteria      => '>=',
                value         => 0,
                input_title   => 'Volume data:',
                input_message => $componentVolumeName{$_}
                  . ( /kVA/ ? ' (except where excluded revenue)' : '' ),
                error_title   => 'Volume data error',
                error_message => 'The volume must be a non-negative number.'
            },
            data          => $componentVolumeData{$_},
            defaultFormat => /k(?:W|VAr)h/ ? '0.000hard' : '0hard',
          )
    } @$nonExcludedComponents;

    Columnset(
        $model->{addVolumes} && $model->{addVolumes} =~ /matching/i
        ? ( name => 'Historical volume data' )
        : (
            name  => 'Volume forecasts for the charging year',
            lines => [
                <<'EOL'
Source: forecast.
Please include MPAN counts for tariffs with no fixed charge (e.g. off-peak tariffs), but exclude MPANs on tariffs with a fixed
charge that are not subject to a fixed charge due to a site grouping arrangement.
EOL
            ]
        ),
        number   => 1053,
        appendTo => $model->{inputTables},
        dataset  => $model->{dataset},
        columns  => [ @volumeData{@$nonExcludedComponents} ]
    );

    return ( \%volumeData )
      unless ref $unitsAdjustmentFactor || !$unitsAdjustmentFactor;

    my %volumesAdjusted;

    if ($unitsAdjustmentFactor) {
        my @adjustedColumns;
        %volumesAdjusted = map {
            if (/Unit rate/i) {
                my $adj = Arithmetic(
                    name       => "$componentVolumeName{$_} loss adjusted",
                    arithmetic => '=IV1*(1+IV2)',
                    arguments  => {
                        IV1 => $volumeData{$_},
                        IV2 => $unitsAdjustmentFactor,
                    }
                );
                push @adjustedColumns, $adj;
                $_ => $adj;
            }
            else {
                $_ => 1
                  ? Stack( sources => [ $volumeData{$_} ] )
                  : $volumeData{$_};
            }
        } @$nonExcludedComponents;
        Columnset(
            name =>
              'Unit volumes adjusted for losses (for embedded network tariffs)',
            columns => 1
            ? [ @volumesAdjusted{@$nonExcludedComponents} ]
            : \@adjustedColumns
        );
    }
    else {
        %volumesAdjusted = %volumeData;
    }

    my $unitsInYear = Arithmetic(
        name =>
          Label( 'All units (MWh)', 'All units aggregated by tariff (MWh)' ),
        arithmetic => '='
          . join( '+', map { "IV$_" } 1 .. $model->{maxUnitRates} ),
        arguments => {
            map { ( "IV$_" => $volumesAdjusted{"Unit rate $_ p/kWh"} ) }
              1 .. $model->{maxUnitRates}
        },
        defaultFormat => '0softnz',
    );

    my %volumesByEndUser = map {
        $_ => GroupBy(
            name => Label(
                $componentVolumeName{$_},
                "$componentVolumeName{$_} aggregated by end user"
            ),
            rows          => $allEndUsers,
            source        => $volumesAdjusted{$_},
            defaultFormat => '0softnz'
          )
    } @$nonExcludedComponents;

    my $unitsByEndUser = GroupBy(
        name =>
          Label( 'All units (MWh)', 'All units aggregated by end user (MWh)' ),
        source        => $unitsInYear,
        rows          => $allEndUsers,
        defaultFormat => '0softnz',
    );

    $unitsByEndUser = Arithmetic(
        name =>
          Label( 'All units (MWh)', 'All units aggregated by end user (MWh)' ),
        arithmetic => '='
          . join( '+', map { "IV$_" } 1 .. $model->{maxUnitRates} ),
        arguments => {
            map { ( "IV$_" => $volumesByEndUser{"Unit rate $_ p/kWh"} ) }
              1 .. $model->{maxUnitRates}
        },
        defaultFormat => '0softnz',
    );

    push @{ $model->{volumeData} }, Columnset(
        name =>
          'Volume forecasts for the charging year, aggregated by end user',
        $model->{portfolio}
          || $model->{boundary} ? () : ( lines => <<EOIDNONOTE),
This table is a straight copy of the original demand forecasts because there are no portfolio tariffs for embedded networks in this version of the model.
If embedded network tariffs are included, this table calculates the aggregate of volume forecasts by type of end user, taking directly connected end users and end users connected to embedded networks together.
EOIDNONOTE
        columns =>
          [ @volumesByEndUser{@$nonExcludedComponents}, $unitsByEndUser ]
    );

    \%volumeData, \%volumesAdjusted, \%volumesByEndUser, $unitsInYear,
      $unitsByEndUser;

}

sub loadProfiles {

    my ( $model, $allEndUsers, $componentMap ) = @_;

    my $unitsEndUsers =
      ( grep { !$componentMap->{$_}{'Unit rate 1 p/kWh'} }
          @{ $allEndUsers->{list} } )
      ? Labelset(
        name => 'Units end users',
        list => [
            grep { $componentMap->{$_}{'Unit rate 1 p/kWh'} }
              @{ $allEndUsers->{list} }
        ]
      )
      : $allEndUsers;

    my $generationEndUsers = Labelset(
        name => 'Generation end users',
        list => [ grep { /generat/i } @{ $allEndUsers->{list} } ]
    );

    my $demandEndUsers = Labelset(
        name => 'Demand end users',
        list => [
            grep { !/generat/i && $componentMap->{$_}{'Unit rate 1 p/kWh'} }
              @{ $allEndUsers->{list} }
        ]
    );

    my $standingForFixedEndUsers = Labelset
      name => 'End users with fixed charges based on standing charges factors',
      list => [
        grep {
                 !/generat/i
              && !/unmeter/i
              && $componentMap->{$_}{'Fixed charge p/MPAN/day'}
              && !$componentMap->{$_}{'Capacity charge p/kVA/day'}
        } @{ $allEndUsers->{list} }
      ];

    my $generationUnitsEndUsers = Labelset
      name => 'Generation unit end users',
      list =>
      [ grep { !$componentMap->{$_}{'Generation capacity rate p/kW/day'} }
          @{ $generationEndUsers->{list} } ];

    my $coincidenceFactors = $model->{coincidenceAdj}
      && $model->{coincidenceAdj} =~ /none/i ? Constant(
        name => 'Not used',
        rows => $allEndUsers,
        data => [],
      ) : Dataset(
        rows       => $demandEndUsers,
        validation => {
            validate      => 'decimal',
            criteria      => 'between',
            minimum       => 0,
            maximum       => 1,
            input_title   => 'Coincidence:',
            input_message => 'Percentage',
            error_message => 'The coincidence factor'
              . ' must be between 0% and 100%.'
        },
        data => [
            map {
                $componentMap->{$_}{'Unit rate 0 p/kWh'}
                  && /(?:related|additional)/i ? undef
                  : /lv/i && /domestic/i && !/non.*domestic/i && /unr/i ? 0.9
                  : /lv/i
                  && /domestic/i && !/non.*domestic/i && /rates/i ? 0.3
                  : /lv/i && /off/i ? 0
                  : 0.5
            } @{ $demandEndUsers->{list} }
        ],
        name => Label(
            'Coincidence factor to system maximum load'
              . ' for each type of demand user',
            'Coincidence factor'
        )
      );

    my $loadFactors;

    unless ( $model->{impliedLoadFactors}
        && $model->{coincidenceAdj}
        && $model->{coincidenceAdj} =~ /none/i )
    {

        $loadFactors = Dataset(
            rows       => $demandEndUsers,
            validation => {
                validate      => 'decimal',
                criteria      => 'between',
                minimum       => 0,
                maximum       => 1,
                input_title   => 'Load factor:',
                input_message => 'Percentage',
                error_message => 'The load factor'
                  . ' must be between 0% and 100%.'
            },
            data => [
                map {
                        /domestic/i && !/non.*domestic/i && /unr/i   ? 0.4
                      : /domestic/i && !/non.*domestic/i && /rates/i ? 0.4
                      : /off/i ? 0.2
                      : /small/i && /unr/i   ? 0.5
                      : /small/i && /rates/i ? 0.5
                      : 0.5
                } @{ $demandEndUsers->{list} }
            ],
            name => Label(
                'Load factor',
                'Load factor' . ' for each type of demand user'
            )
        );

        Columnset(
            name     => 'Load profile data for demand users',
            lines    => 'Source: load data analysis.',
            number   => 1041,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            columns  => [
                grep { ref $_ eq 'SpreadsheetModel::Dataset' }
                  $coincidenceFactors,
                $loadFactors
            ],
        );

    }

    my $loadCoefficients;

    if (   $model->{coincidenceAdj}
        && $model->{coincidenceAdj} =~ /none/i )
    {

        $loadCoefficients = Constant(
            name => 'Demand/generation indicator; notional load coefficient',
            rows => $allEndUsers,
            data => [ map { /gener/i ? -1 : 1; } @{ $allEndUsers->{list} } ],
        );

    }
    else {

        my $generationCoefficient = Constant(
            rows => $generationEndUsers,
            data => [ map { 1 } @{ $generationEndUsers->{list} } ],
            name => Label(
                'Generation coefficient',
                'Generation coefficient (negative of load coefficient)'
            ),
            lines  => 'Source: assumption.',
            number => 1044,
        );

        push @{ $model->{loadProfiles} },
          my $demandCoefficient = Arithmetic(
            name => Label(
                'Demand coefficient',
                'Demand coefficient (load at time of '
                  . 'system maximum load divided by average load)'
            ),
            arithmetic => '=IV1/IV2',
            arguments  => { IV1 => $coincidenceFactors, IV2 => $loadFactors }
          );

        my $negGC = Arithmetic(
            name       => 'Negative of generation coefficient',
            arithmetic => '=-1*IV1',
            arguments  => { IV1 => $generationCoefficient }
        );

        $negGC = Constant(
            name => 'Negative of generation coefficient; set to -1',
            rows => $generationEndUsers,
            cols => 0,
            data => [ [ map { -1 } @{ $generationEndUsers->{list} } ] ]
        );

        $loadCoefficients = Stack(
            name    => 'Load coefficient',
            rows    => $allEndUsers,
            sources => [ $demandCoefficient, $negGC ]
        );
    }

    push @{ $model->{loadProfiles} }, $loadCoefficients;

    my $fFactors;
    my $generationCapacityEndUsers = Labelset
      name => 'Generation capacity end users',
      list =>
      [ grep { $componentMap->{$_}{'Generation capacity rate p/kW/day'} }
          @{ $generationEndUsers->{list} } ];
    if ( @{ $generationCapacityEndUsers->{list} } ) {
        $model->{hasGenerationCapacity} = 1;
        push @{ $model->{loadProfiles} }, $fFactors = Dataset(
            rows       => $generationCapacityEndUsers,
            validation => {
                validate      => 'decimal',
                criteria      => 'between',
                minimum       => 0,
                maximum       => 1,
                input_title   => 'F factor:',
                input_message => 'Percentage',
                error_message => 'The F factor'
                  . ' must be between 0% and 100%.'
            },
            number => 1043,
            data   => [
                map {
                        /wind/i       ? 0.24
                      : /non[- ]chp/i ? 0.73
                      : /chp/i        ? 0.67
                      : 0.36
                } @{ $generationCapacityEndUsers->{list} }
            ],
            name => Label( 'F factor', 'F factors by generation technology' ),
            lines =>
              'Source: assumption based on Engineering Recommendation P2/6.'
        );
    }

    $unitsEndUsers, $generationEndUsers, $demandEndUsers,
      $standingForFixedEndUsers, $generationCapacityEndUsers,
      $generationUnitsEndUsers, $fFactors, $loadCoefficients, $loadFactors;

}

1;
