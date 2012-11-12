package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 DCUSA Limited and contributors. All rights reserved.

Redistribution and use in source and binary forms, with or without modification,
are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation and/or
other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY DCUSA LIMITED AND CONTRIBUTORS "AS IS" AND ANY
EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL DCUSA LIMITED OR CONTRIBUTORS BE LIABLE FOR ANY
DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

=cut

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
                lines    => 'Source: service models.',
                location => $model->{configSheet},
                name     => "Names of network components in $name",
                rows     => Labelset(
                    name => "Possible components in $name",
                    list => [ map { "$name item $_" } 1 .. @{ $_->{list} } ]
                ),
                defaultFormat => 'textinput',
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
            arithmetic    => '=IV1*IV2',
            defaultFormat => '0softnz',
            arguments     => {
                IV1 => $itemCounts,
                IV2 => $unitCosts,
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
            data          => $voltage =~ /LV/i
            ? [ 6000, 500, 1200, 800, 1000, map { 0 } 1 .. 10 ]
            : [ 12000, 20000, 10000, map { 0 } 1 .. 10 ]
        ) unless $model->{detailedCosts};

        push @{ $model->{serviceModels} }, $serviceModelTotal;

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
                        /microgener|(additional|related) MPAN|gener.*non.*half/i
                          ? [ map { 0 }
                              1 .. @{ $serviceModels{$voltage}{list} } ]

                          : /domestic/i && !/non.*domestic/i
                          ? [ map { $_ == 1 ? 0.05 : 0 }
                              1 .. @{ $serviceModels{$voltage}{list} } ]

                          : /HV.*gener/i ? [ map { $_ == 3 ? 0 : 0 }
                              1 .. @{ $serviceModels{$voltage}{list} } ]

                          : /HV/ && /5-8/ ? [ map { $_ == 1 ? 1 : 0 }
                              1 .. @{ $serviceModels{$voltage}{list} } ]

                          : /HV sub/ ? [ map { $_ == 2 ? 1 : 0 }
                              1 .. @{ $serviceModels{$voltage}{list} } ]

                          : /HV/ ? [ map { $_ < 3 ? 0.5 : 0 }
                              1 .. @{ $serviceModels{$voltage}{list} } ]

                          : /LV.*gener/i ? [ map { $_ == 5 ? 0 : 0 }
                              1 .. @{ $serviceModels{$voltage}{list} } ]

                          : /small/i ? [ map { $_ == 2 ? 1 : 0 }
                              1 .. @{ $serviceModels{$voltage}{list} } ]

                          : /half/i ? [ map { $_ == 3 ? 1 : 0 }
                              1 .. @{ $serviceModels{$voltage}{list} } ]

                          : [ map { $_ == 2 ? .5 : $_ == 3 ? .5 : 0 }
                              1 .. @{ $serviceModels{$voltage}{list} } ]
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

            push @{ $model->{serviceModels} },
              my $serviceModelAssetsPerCustomer = SumProduct(
                name   => "Asset £/customer from $voltage service models",
                matrix => $serviceModelMatrix,
                vector => $serviceModelTotal,
                cols   => $relevantLevel,
                defaultFormat => '0softnz',
              );

            push @serviceModelAssetArray, $serviceModelAssetsPerCustomer;

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
                $model->{tariffs} =~ /dcp130/i
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
                          ? [ map { $_ == 4 ? 0.12 : 0 }
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
                name   => "Asset £/(MWh/year) from $voltage service models",
                matrix => $serviceModelMatrix,
                vector => $serviceModelTotal,
                cols   => $relevantLevel,
                defaultFormat => '0softnz',
            );

            $serviceModelCostPerAnnualMwh = Arithmetic(
                name => 'Service model asset p/kWh charge for'
                  . ' unmetered tariffs',
                arithmetic => '=0.1*IV5*IV1*IV3',
                arguments  => {
                    IV1 => $serviceModelAssetsPerAnnualMwh,
                    IV3 => $annuityRate,
                    IV5 => $proportionChargeable
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
        arithmetic => '=100/IV2*IV1*IV3*IV5',
        arguments  => {
            IV1 => $serviceModelAssetsPerCustomer,
            IV2 => $daysInYear,
            IV3 => $annuityRate,
            IV5 => $proportionChargeable
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

    push @{ $model->{siteSpecific} }, my $ssr = Arithmetic(
        name => 'Income for site specific sole use asset replacement (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => '=IV1*IV3*IV5',
        arguments     => {
            IV1 => $sssua,
            IV3 => $annuityRate,
            IV5 => $proportionChargeable
        }
    );

    $sssua, $ssr;

}

1;
