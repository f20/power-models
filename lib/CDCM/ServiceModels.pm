package CDCM;

# Copyright 2009-2011 Energy Networks Association Limited and others.
# Copyright 2016-2019 Franck Latrémolière, Reckon LLP and others.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice,
# this list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
# this list of conditions and the following disclaimer in the documentation
# and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
# EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
# THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Shortcuts ':all';

sub serviceModels {

    my ( $model, $daysInYear, $modelLife, $annuityRate, $proportionChargeable,
        $allTariffsByEndUser, $componentMap )
      = @_;

    if ( $model->{noSM} ) {
        push @{ $model->{optionLines} }, 'No service models';
        return;
    }

    my %serviceModels;
    my @serviceModelAssetArray;
    my ( $serviceModelAssetsPerAnnualMwh, $serviceModelCostPerAnnualMwh );

    $serviceModels{LV} = Labelset(
        name => 'LV service models',
        list => [

            Labelset(
                name => 'LV service model 1',
                list => [ map { $_ ? $_ : ' ' } split /\n/, <<'EOL'] ),
(service model for 20 domestic connections)
Same side underground connection (single-phase 185mm waveform)
Cross road underground connection (single-phase 185mm waveform)

Other assets
EOL

            Labelset(
                name => 'LV service model 2',
                list => [ map { $_ ? $_ : ' ' } split /\n/, <<'EOL'] ),
(service model for one 24-60kVA connection)
Underground connection (three-phase 185mm waveform)

Other assets
EOL

            Labelset(
                name => 'LV service model 3',
                list => [ map { $_ ? $_ : ' ' } split /\n/, <<'EOL'] ),
(service model for one larger LV connection)
Underground connection (three-phase 300mm waveform)

Other assets
EOL

            Labelset(
                name => 'LV service model 4',
                list => [ map { $_ ? $_ : ' ' } split /\n/, <<'EOL'] ),
(service model for 40 lampposts)
Some cables and stuff

Other assets
EOL

            Labelset(
                name => 'LV service model 5',
                list => [ map { $_ ? $_ : ' ' } split /\n/, <<'EOL'] ),
(service model for additional generator protection at LV)
Some circuit breakers and stuff

Other assets
EOL

            Labelset(
                name => 'LV service model 6',
                list => [ map { $_ ? $_ : ' ' } split /\n/, <<'EOL'] ),
(spare)

Other assets
EOL

            Labelset(
                name => 'LV service model 7',
                list => [ map { $_ ? $_ : ' ' } split /\n/, <<'EOL'] ),
(spare)

Other assets
EOL

            Labelset(
                name => 'LV service model 8',
                list => [ map { $_ ? $_ : ' ' } split /\n/, <<'EOL'] ),
(spare)

Other assets
EOL

        ]
    );

    $serviceModels{HV} = Labelset(
        name => 'HV service models',
        list => [

            Labelset(
                name => 'HV service model 1',
                list => [ map { $_ ? $_ : ' ' } split /\n/, <<'EOL'] ),
(service model for one HV connection up to 1.5MVA)
Ring main unit
CTs and VTs
Man-hours for live jointing

Other assets
EOL

            Labelset(
                name => 'HV service model 2',
                list => [ map { $_ ? $_ : ' ' } split /\n/, <<'EOL'] ),
(service model for one HV connection above 1.5MVA)
Bigger ring main unit
Bigger CTs and VTs
Man-hours for live jointing

Other assets
EOL

            Labelset(
                name => 'HV service model 3',
                list => [ map { $_ ? $_ : ' ' } split /\n/, <<'EOL'] ),
(service model for additional generator protection at HV)
Some HV circuit breakers and stuff

Other assets
EOL

            Labelset(
                name => 'HV service model 4',
                list => [ map { $_ ? $_ : ' ' } split /\n/, <<'EOL'] ),
(spare)

Other assets
EOL

            Labelset(
                name => 'HV service model 5',
                list => [ map { $_ ? $_ : ' ' } split /\n/, <<'EOL'] ),
(spare)

Other assets
EOL

        ]
    );

    foreach my $voltage (qw(LV HV)) {

        my $relevantLevel = Labelset( list => ["Assets\t$voltage customer"] );

        my @everything =
          map { ( '', @{ $_->{list} } ) } @{ $serviceModels{$voltage}{list} };

        my @items = map {
                /Same side underground connection \(single-phase/  ? 12
              : /Cross road underground connection \(single-phase/ ? 8
              : /three-phase|ring main|cts and vts/i               ? 1
              : /hours/i                                           ? 3
              : /stuff/i                                           ? 1
              : /other assets/i                                    ? 1
              : 0

        } @everything;

        my @cost = map {
                /Same side underground connection \(single-phase/  ? 160
              : /Cross road underground connection \(single-phase/ ? 240
              : /connection \(three-phase 185mm/                   ? 400
              : /connection \(three-phase 300mm/                   ? 600
              : /bigger ring main unit/i                           ? 10000
              : /ring main unit/i                                  ? 8000
              : /bigger CTs and VTs/i                              ? 3000
              : /CTs and VTs/i                                     ? 2000
              : /hours/i                                           ? 40
              : /HV circuit breakers/i                             ? 1500
              : /stuff/i                                           ? 1000
              : 0
        } @everything;

        foreach ( @{ $serviceModels{$voltage}{list} } ) {
            my $name = $_->{name};
            $_ = Labelset
              name     => $name,
              editable => Dataset(
                lines => 'Source: service models.',
                name  => "Names of network components in $name",
                rows  => Labelset(
                    name => "Possible components in $name",
                    list => [ map { "$name item $_" } 1 .. @{ $_->{list} } ]
                ),
                defaultFormat => 'texthard',
                data          => $_->{list}
              );
        }

        my $itemsByServiceModel = Labelset(
            name   => "Items in $voltage service models",
            groups => $serviceModels{$voltage}{list}
        );

        my $itemCounts = Dataset(
            name       => 'Item count',
            validation => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
            defaultFormat => '0hardnz',
            rows          => $itemsByServiceModel,
            data          => \@items
        );

        my $unitCosts = Dataset(
            name       => 'Item cost (£)',
            validation => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
            rows          => $itemsByServiceModel,
            defaultFormat => '0hardnz',
            data          => \@cost
        );

        my $replacement = Arithmetic(
            name          => 'Asset cost (£)',
            arithmetic    => '=A1*A2',
            defaultFormat => '0softnz',
            arguments     => {
                A1 => $itemCounts,
                A2 => $unitCosts,
            }
        );

        push @{ $model->{serviceModels} },
          Columnset(
            name   => "$voltage service models",
            lines  => 'Source: service models.',
            number => $voltage eq 'LV' ? 1024 : $voltage eq 'HV' ? 1027 : undef,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            columns  => [ $itemCounts, $unitCosts, $replacement ]
          ) if $model->{detailedCosts};

        my $serviceModelTotal = GroupBy(
            name          => $voltage . ' service model asset cost (£)',
            defaultFormat => '0soft',
            cols          => $serviceModels{$voltage},
            rows          => 0,
            source        => $replacement
        );

        $serviceModelTotal = Dataset(
            name       => $voltage . ' service model asset cost (£)',
            validation => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
            number => $voltage =~ /lv/i ? 1022 : 1023,
            appendTo      => $model->{inputTables},
            dataset       => $model->{dataset},
            defaultFormat => '0hard',
            cols          => $serviceModels{$voltage},
            rows          => 0,
            data =>
              [ 10000, map { 0 } 1 .. $#{ $serviceModels{$voltage}{list} } ]
        ) unless $model->{detailedCosts};

        push @{ $model->{serviceModels} }, $serviceModelTotal;

        $model->{sharedData}->addStats( 'Input data (Schedule 16 paragraph 33)',
            $model, $serviceModelTotal )
          if $model->{sharedData};

        if (
            my @list =
            grep {
                     /^$voltage/
                  && !/(additional|related) MPAN/i
                  && !/edcm/i
                  && $componentMap->{$_}{'Fixed charge p/MPAN/day'}
            } map { $allTariffsByEndUser->{list}[$_] }
            $allTariffsByEndUser->indices
          )
        {

            my $relevantTariffs = Labelset( list => \@list );

            my $serviceModelMatrix = Dataset(
                name => 'Matrix of applicability of '
                  . "$voltage service models to tariffs with fixed charges",
                rows  => $relevantTariffs,
                cols  => $serviceModels{$voltage},
                byrow => 1,
                data  => [
                    map {
                        [ map { 0 } 1 .. @{ $serviceModels{$voltage}{list} } ]
                    } @{ $relevantTariffs->{list} }
                ],
                number => $voltage eq 'LV' ? 1025
                : $voltage eq 'HV' ? 1028
                : undef,
                appendTo      => $model->{inputTables},
                dataset       => $model->{dataset},
                defaultFormat => '%hardnz',
                validation    => {
                    validate      => 'decimal',
                    criteria      => 'between',
                    minimum       => 0,
                    maximum       => 1,
                    input_title   => 'Percentage of service model',
                    input_message => 'Between 0% and 100%',
                    error_message => 'The number in this cell'
                      . ' must be between 0% and 100%.'
                },
            );

            $model->{sharedData}
              ->addStats( 'Input data (Schedule 16 paragraph 35A)',
                $model, $serviceModelMatrix )
              if $model->{sharedData};

            push @{ $model->{serviceModels} },
              my $serviceModelAssetsPerCustomer = SumProduct(
                name   => "Asset £/customer from $voltage service models",
                matrix => $serviceModelMatrix,
                vector => $serviceModelTotal,
                cols   => $relevantLevel,
                defaultFormat => '0softnz',
              );

            push @serviceModelAssetArray, $serviceModelAssetsPerCustomer;

            $model->{sharedData}
              ->addStats( 'Input data (15 months notice of service models)',
                $model, $serviceModelAssetsPerCustomer )
              if $model->{sharedData};

        }

        if (
            $voltage eq 'LV' && (
                my @list = grep {
                         /^$voltage/
                      && !$componentMap->{$_}{'Fixed charge p/MPAN/day'}
                      && !/(additional|related) MPAN/i
                      && !/microgen/i
                } map { $allTariffsByEndUser->{list}[$_] }
                $allTariffsByEndUser->indices
            )
          )
        {

            my $relevantTariffs = Labelset( list => \@list );

            my $serviceModelMatrix = Dataset(
                validation => {
                    validate => 'decimal',
                    criteria => '>=',
                    value    => 0,
                },
                lines => [
                    'Source: service models',
                    'Proportion of service model involved'
                      . ' in connecting load of 1 MWh/year',
                ],
                $model->{tariffs} && $model->{tariffs} =~ /dcp130/i
                ? (
                    name => Label(
                        "All $voltage unmetered tariffs",
                        'Matrix of applicability of '
                          . "$voltage service models to unmetered tariffs",
                    )
                  )
                : (
                    name => 'Matrix of applicability of '
                      . "$voltage service models to unmetered tariffs",
                    rows => $relevantTariffs
                ),
                cols  => $serviceModels{$voltage},
                byrow => 1,
                data  => [
                    map {
                        /un-?met/i
                          && !($model->{opCoded}
                            && $model->{opCoded} =~ /ums/i )
                          ? [ map { 0 }
                              1 .. @{ $serviceModels{$voltage}{list} } ]
                          : [ map { 0 } @{ $serviceModels{$voltage}{list} } ]
                    } @{ $relevantTariffs->{list} }
                ],
                number   => 1026,
                appendTo => $model->{inputTables},
                dataset  => $model->{dataset},
                1 ? () : ( defaultFormat => '%hardnz' )
            );

            $serviceModelAssetsPerAnnualMwh = SumProduct(
                name => "$voltage unmetered service model assets £/(MWh/year)",
                matrix        => $serviceModelMatrix,
                vector        => $serviceModelTotal,
                cols          => $relevantLevel,
                defaultFormat => '0softnz',
            );

            $model->{sharedData}
              ->addStats( 'Input data (15 months notice of service models)',
                $model, $serviceModelAssetsPerAnnualMwh )
              if $model->{sharedData};

            $serviceModelCostPerAnnualMwh = Arithmetic(
                name => "$voltage unmetered service model asset charge (p/kWh)",
                arithmetic => '=0.1*A5*A1*A3',
                arguments  => {
                    A1 => $serviceModelAssetsPerAnnualMwh,
                    A3 => $annuityRate,
                    A5 => $proportionChargeable
                }
            );

            push @{ $model->{serviceModels} }, $serviceModelCostPerAnnualMwh;

        }
    }

    my $lvHvCustomerLabelset =
      Labelset( list => [ "Assets\tLV customer", "Assets\tHV customer" ] );

    push @serviceModelAssetArray,
      Constant(
        name          => 'Zero if not applicable',
        defaultFormat => '0connz',
        rows          => $allTariffsByEndUser,
        cols          => $lvHvCustomerLabelset,
        data          => [ [0], [0] ],
      ) unless $#{ $allTariffsByEndUser->{list} };

    my $serviceModelAssetsPerCustomer = Stack(
        name    => 'Service model assets by tariff (£)',
        rows    => $allTariffsByEndUser,
        cols    => $lvHvCustomerLabelset,
        sources => \@serviceModelAssetArray
    );

    my $serviceModelCostPerCustomerDetail = Arithmetic(
        name       => 'Service model p/MPAN/day charge',
        arithmetic => '=100/A2*A1*A3*A5',
        arguments  => {
            A1 => $serviceModelAssetsPerCustomer,
            A2 => $daysInYear,
            A3 => $annuityRate,
            A5 => $proportionChargeable
        }
    );

    my $serviceModelCostPerCustomer = GroupBy(
        rows   => $allTariffsByEndUser,
        cols   => 0,
        source => $serviceModelCostPerCustomerDetail,
        name   => 'Service model p/MPAN/day'
    );

    push @{ $model->{serviceModels} },
      Columnset(
        name => 'Replacement annuities for service models',
        columns =>
          [ $serviceModelCostPerCustomerDetail, $serviceModelCostPerCustomer ]
      );

    $serviceModelAssetsPerCustomer,    $serviceModelCostPerCustomer,
      $serviceModelAssetsPerAnnualMwh, $serviceModelCostPerAnnualMwh;
}

sub siteSpecificSoleUse {

    my ( $model, $drmLevels, $modelLife, $annuityRate, $proportionChargeable, )
      = @_;

    my $sssua = Dataset(
        name => 'Site specific sole use assets (£/year)',
        $model->{opAlloc}
        ? ()
        : (
            lines => [
'This table is for the MEAV of EDCM site-specific sole use assets.',
'You can put the assets in any network level -- it makes no difference.'
            ]
        ),
        validation => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
        defaultFormat => '0hardnz',
        number        => 1080,
        appendTo      => $model->{inputTables},
        dataset       => $model->{dataset},
        cols          => $model->{ehv} =~ /ss/ ? 0 : $drmLevels,
        data => [ map { /\sEHV$/i ? 100e6 : 0 } @{ $drmLevels->{list} } ]
    );

    push @{ $model->{siteSpecific} },
      my $ssr = Arithmetic(
        name => 'Income for site specific sole use asset replacement (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=A1*A3*A5',
        arguments     => {
            A1 => $sssua,
            A3 => $annuityRate,
            A5 => $proportionChargeable
        }
      );

    $sssua, $ssr;

}

1;
