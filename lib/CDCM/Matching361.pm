package CDCM;

# Copyright 2020-2021 Franck Latrémolière and others.
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

sub matching361 {

    my (
        $model,                  $adderAmount,
        $componentMap,           $allTariffsByEndUser,
        $demandTariffsByEndUser, $allEndUsers,
        $chargingLevels,         $nonExcludedComponents,
        $allComponents,          $daysAfter,
        $volumeAfterUndoctored,  $loadCoefficients,
        $tariffsExMatching,      $daysFullYear,
        $volumeFullYear
    ) = @_;

    my @unchargeableFixed = grep { /related|unmet|ums/i; }
      @{ $volumeAfterUndoctored->{'Fixed charge p/MPAN/day'}->{rows}{list} };
    my $volumeAfter = {
        %$volumeAfterUndoctored,
        'Fixed charge p/MPAN/day' => Stack(
            name => 'Chargeable MPANs',
            rows => $volumeAfterUndoctored->{'Fixed charge p/MPAN/day'}->{rows},
            sources => [
                Constant(
                    defaultFormat => '0con',
                    rows          => Labelset( list => \@unchargeableFixed ),
                    data          => [ [ map { 0; } @unchargeableFixed ] ],
                ),
                $volumeAfterUndoctored->{'Fixed charge p/MPAN/day'},
            ],
        ),
    };

    my $tcrGroupAllocation = Dataset(
        name          => 'TCR group allocation',
        defaultFormat => '0hard',
        dataset       => $model->{dataset},
        number        => 1071,
        appendTo      => $model->{inputTables},
        rows          => $allTariffsByEndUser,
        data          => [
            map {
                    /domestic/i && !/non.domestic/i           ? 1
                  : /Non-Domestic Aggregated.* Band ([1-4])/i ? 4 + $1
                  : /LV.* Site Specific Band ([1-4])/i        ? 8 + $1
                  : /HV Site Specific Band ([1-4])/i          ? 12 + $1
                  : /unmetered/i                              ? 17
                  :                                             0;
            } @{ $demandTariffsByEndUser->{list} }
        ],
        validation => {
            validate      => 'list',
            value         => [ 0, 1, 5 .. 17 ],
            criteria      => '>=',
            value         => 0,
            input_title   => 'TCR group:',
            input_message => 'TCR group (between 0 and 17)',
            error_title   => 'Invalid TCR group',
            error_message =>
              'Invalid TCR group (must be an integer between 0 and 17).'
        },
        usePlaceholderData => 1,
    );

    my $tcrGroupset =
      Labelset( list => [ map { "TCR group $_"; } 1, 5 .. 17 ] );

    my $demandTariffsSubjectToScaling = Labelset( list =>
          [ grep { !/related mpan/i; } @{ $demandTariffsByEndUser->{list} } ] );

    my @columnsets;

    foreach my $tcrGroup ( 1, 5 .. 17 ) {

        my @allUnitsColumns = map {
            Arithmetic(
                name          => $_->objectShortName,
                defaultFormat => '0copy',
                arithmetic    => "=IF(A1=$tcrGroup,A2" . ',"")',
                arguments     => {
                    A1 => $tcrGroupAllocation,
                    A2 => $_,
                },
            );
        } @{$volumeAfter}{ @{$nonExcludedComponents}[ 0 .. 2 ] };

        my @volumeColumns = map {
            Arithmetic(
                name          => $_->objectShortName,
                defaultFormat => '0copy',
                rows          => $demandTariffsSubjectToScaling,
                arithmetic    => "=IF(A1=$tcrGroup,A2" . ',"")',
                arguments     => {
                    A1 => $tcrGroupAllocation,
                    A2 => $_,
                },
            );
        } @{$volumeAfter}{ @{$nonExcludedComponents}[ 0 .. 3 ] };

        my @priceColumns = map {
            Arithmetic(
                name => $tariffsExMatching->{$_}->objectShortName,
                rows => $demandTariffsSubjectToScaling,
                /day/ ? ( defaultFormat => '0.00copy' )
                : ( defaultFormat => '0.000copy' ),
                arithmetic => "=IF(A1=$tcrGroup,A2"
                  . (
                    defined $model->{pcd}{preRoundingFiddles}{$_} ? '+A3' : ''
                  )
                  . ',"")',
                arguments => {
                    A1 => $tcrGroupAllocation,
                    A2 => $tariffsExMatching->{$_},
                    defined $model->{pcd}{preRoundingFiddles}{$_}
                    ? ( A3 => $model->{pcd}{preRoundingFiddles}{$_} )
                    : (),
                },
            );
        } @{$nonExcludedComponents}[ 0 .. 3 ];

        Columnset(
            name => "Extraction of pre-matching data for TCR group $tcrGroup",
            columns => [ @volumeColumns, @priceColumns ],
        );

        Columnset(
            name => 'Extraction of all (including related MPAN)'
              . " pre-matching units related to TCR group $tcrGroup",
            columns => \@allUnitsColumns,
        );

        my $row                 = Labelset( list => ["TCR group $tcrGroup"] );
        my @minimumPriceColumns = map {
            Arithmetic(
                name => $_->objectShortName,
                rows => $row,
                $_->objectShortName =~ /day/
                ? ( defaultFormat => '0.00soft' )
                : (),
                arithmetic => '=MIN(A1_A2)',
                arguments  => { A1_A2 => $_, },
            );
        } @priceColumns;
        my @allVolumesColumns = map {
            Arithmetic(
                name          => $_->objectShortName,
                rows          => $row,
                defaultFormat => '0soft',
                arithmetic    => '=SUM(A1_A2)',
                arguments     => { A1_A2 => $_, },
            );
        } @allUnitsColumns, @volumeColumns;
        push @columnsets,
          Columnset(
            name => "Aggregation of pre-matching data for TCR group $tcrGroup",
            columns => [ @allVolumesColumns, @minimumPriceColumns, ]
          );
    }

    my @tcrGroupData;
    for ( my $c = 0 ; $c < @{ $columnsets[0]{columns} } ; ++$c ) {
        my $name = $columnsets[0]{columns}[$c]->objectShortName;
        push @tcrGroupData,
          Stack(
            name => $name,
            defaultFormat => $name !~ /\// ? '0copy'
            : $name =~ /day/ ? '0.00copy'
            : '0.000copy',
            rows    => $tcrGroupset,
            sources => [ map { $_->{columns}[$c]; } @columnsets ],
          );
    }

    Columnset(
        name    => 'Aggregation of pre-matching data for all TCR groups',
        columns => \@tcrGroupData,
    );

    push @{ $model->{adderResults} }, my $adderByGroup = Arithmetic(
        name => 'Residual amount to be applied to each TCR group (£/year)',
        defaultFormat => '0soft',
        arithmetic => '=(A1+A2+A3)*A4/(SUM(A15_A16)+SUM(A25_A26)+SUM(A35_A36))',
        arguments  => {
            A4 => $adderAmount,
            map {
                (
                    "A$_"           => $tcrGroupData[ $_ - 1 ],
                    "A${_}5_A${_}6" => $tcrGroupData[ $_ - 1 ]
                );
            } 1 .. 3,
        },
    );

    my $lowRateId = Arithmetic(
        name          => 'Lowest rate',
        defaultFormat => 'indexsoft',
        arithmetic    => '=IF(A1<A2,IF(A11<A3,1,3),IF(A21<A31,2,3))',
        arguments     => {
            A1  => $tcrGroupData[7],
            A11 => $tcrGroupData[7],
            A2  => $tcrGroupData[8],
            A21 => $tcrGroupData[8],
            A3  => $tcrGroupData[9],
            A31 => $tcrGroupData[9],
        },
    );

    my $highRateId = Arithmetic(
        name          => 'Highest rate',
        defaultFormat => 'indexsoft',
        arithmetic    => '=IF(A2<A3,IF(A1<A31,3,1),IF(A11<A21,2,1))',
        arguments     => {
            A1  => $tcrGroupData[7],
            A11 => $tcrGroupData[7],
            A2  => $tcrGroupData[8],
            A21 => $tcrGroupData[8],
            A3  => $tcrGroupData[9],
            A31 => $tcrGroupData[9],
        },
    );

    my $middleRateId = Arithmetic(
        name          => 'Middle rate',
        defaultFormat => 'indexsoft',
        arithmetic    => '=6-A1-A2',
        arguments     => {
            A1 => $lowRateId,
            A2 => $highRateId,
        },
    );

    my $lowRate = Arithmetic(
        name          => 'Low rate p/kWh',
        defaultFormat => '0.000copy',
        arithmetic    => '=IF(A8=1,A1,IF(A9=2,A2,A3))',
        arguments     => {
            A8 => $lowRateId,
            A9 => $lowRateId,
            A1 => $tcrGroupData[7],
            A2 => $tcrGroupData[8],
            A3 => $tcrGroupData[9],
        },
    );

    my $middleRate = Arithmetic(
        name          => 'Middle rate p/kWh',
        defaultFormat => '0.000copy',
        arithmetic    => '=IF(A8=1,A1,IF(A9=2,A2,A3))',
        arguments     => {
            A8 => $middleRateId,
            A9 => $middleRateId,
            A1 => $tcrGroupData[7],
            A2 => $tcrGroupData[8],
            A3 => $tcrGroupData[9],
        },
    );

    my $highRate = Arithmetic(
        name          => 'High rate p/kWh',
        defaultFormat => '0.000copy',
        arithmetic    => '=IF(A8=1,A1,IF(A9=2,A2,A3))',
        arguments     => {
            A8 => $highRateId,
            A9 => $highRateId,
            A1 => $tcrGroupData[7],
            A2 => $tcrGroupData[8],
            A3 => $tcrGroupData[9],
        },
    );

    my $lowUnits = Arithmetic(
        name          => 'Low rate MWh',
        defaultFormat => '0copy',
        arithmetic    => '=IF(A8=1,A1,IF(A9=2,A2,A3))',
        arguments     => {
            A8 => $lowRateId,
            A9 => $lowRateId,
            A1 => $tcrGroupData[3],
            A2 => $tcrGroupData[4],
            A3 => $tcrGroupData[5],
        },
    );

    my $middleUnits = Arithmetic(
        name          => 'Middle rate MWh',
        defaultFormat => '0copy',
        arithmetic    => '=IF(A8=1,A1,IF(A9=2,A2,A3))',
        arguments     => {
            A8 => $middleRateId,
            A9 => $middleRateId,
            A1 => $tcrGroupData[3],
            A2 => $tcrGroupData[4],
            A3 => $tcrGroupData[5],
        },
    );

    my $highUnits = Arithmetic(
        name          => 'High rate MWh',
        defaultFormat => '0copy',
        arithmetic    => '=IF(A8=1,A1,IF(A9=2,A2,A3))',
        arguments     => {
            A8 => $highRateId,
            A9 => $highRateId,
            A1 => $tcrGroupData[3],
            A2 => $tcrGroupData[4],
            A3 => $tcrGroupData[5],
        },
    );

    Columnset(
        name    => 'Units data for all TCR groups in priority order',
        columns => [
            $lowRateId,   $highRateId, $middleRateId,
            $lowRate,     $lowUnits,   $middleRate,
            $middleUnits, $highRate,   $highUnits,
        ],
    );

    my $fixedAdder = Arithmetic(
        name          => 'Adder p/MPAN/day',
        defaultFormat => '0.00soft',
        arithmetic    => '=IF(A2,MAX(A1/A3/A9*100,0-A4),0)',
        arguments     => {
            A1 => $adderByGroup,
            A2 => $tcrGroupData[6],
            A3 => $tcrGroupData[6],
            A4 => $tcrGroupData[10],
            A9 => $daysAfter,
        },
    );

    my $unitsAdder1 = Arithmetic(
        name          => 'Adder low units p/kWh',
        defaultFormat => '0.000soft',
        arithmetic    =>
          '=IF(A21+A22+A23,MAX((A1-A5*A6*A9*0.01)/(A31+A32+A33)*0.1,0-A4),0)',
        arguments => {
            A1  => $adderByGroup,
            A21 => $lowUnits,
            A22 => $middleUnits,
            A23 => $highUnits,
            A31 => $lowUnits,
            A32 => $middleUnits,
            A33 => $highUnits,
            A4  => $lowRate,
            A5  => $fixedAdder,
            A6  => $tcrGroupData[6],
            A9  => $daysAfter,
        },
    );

    my $unitsAdder2 = Arithmetic(
        name          => 'Adder middle units p/kWh',
        defaultFormat => '0.000soft',
        arithmetic    =>
'=IF(A22+A23,MAX((A1-A5*A6*A9*0.01-A21*A51*10)/(A32+A33)*0.1,0-A4),0)',
        arguments => {
            A1  => $adderByGroup,
            A21 => $lowUnits,
            A22 => $middleUnits,
            A23 => $highUnits,
            A32 => $middleUnits,
            A33 => $highUnits,
            A4  => $middleRate,
            A5  => $fixedAdder,
            A51 => $unitsAdder1,
            A6  => $tcrGroupData[6],
            A9  => $daysAfter,
        },
    );

    my $unitsAdder3 = Arithmetic(
        name          => 'Adder high units p/kWh',
        defaultFormat => '0.000soft',
        arithmetic    =>
'=IF(A23,MAX((A1-A5*A6*A9*0.01-A21*A51*10-A22*A52*10)/A33*0.1,0-A4),0)',
        arguments => {
            A1  => $adderByGroup,
            A21 => $lowUnits,
            A22 => $middleUnits,
            A23 => $highUnits,
            A33 => $highUnits,
            A4  => $highRate,
            A5  => $fixedAdder,
            A51 => $unitsAdder1,
            A52 => $unitsAdder2,
            A6  => $tcrGroupData[6],
            A9  => $daysAfter,
        },
    );

    push @{ $model->{adderResults} },
      Columnset(
        name    => 'Adders by TCR group',
        columns => [ $fixedAdder, $unitsAdder1, $unitsAdder2, $unitsAdder3, ]
      );

    my $adderTable;

    foreach (@$nonExcludedComponents) {
        next unless /kWh|MPAN/;
        $adderTable->{$_} = Arithmetic(
            name => "Adder on $_",
            cols => Labelset( list => ['Adder'] ),
            /MPAN/
            ? (
                defaultFormat => '0.00soft',
                arithmetic    => '=IF(A11,INDEX(A2_A3,A1-IF(A12>1,3,0)),0)',
                arguments     => {
                    A2_A3 => $fixedAdder,
                    A1    => $tcrGroupAllocation,
                    A11   => $tcrGroupAllocation,
                    A12   => $tcrGroupAllocation,
                },
              )
            : (
                arithmetic => '=IF(A11,MAX('
                  . '0-INDEX(A4_A5,A13-IF(A14>1,3,0)),'
                  . 'INDEX(A2_A3,A1-IF(A12>1,3,0))),0)',
                arguments => {
                    A4_A5 => $tcrGroupData[ 6 + (/rate ([0-9])/)[0] ],
                    A2_A3 => $unitsAdder3,
                    A1    => $tcrGroupAllocation,
                    A11   => $tcrGroupAllocation,
                    A12   => $tcrGroupAllocation,
                    A13   => $tcrGroupAllocation,
                    A14   => $tcrGroupAllocation,
                },
            ),
        );
    }

    my $revenuesFromAdder;

    {
        my @termsNoDays;
        my @termsWithDays;
        my %args = ( A400 => $daysAfter );
        my $i    = 1;

        foreach ( grep { $adderTable->{$_} } @$nonExcludedComponents ) {
            ++$i;
            my $pad = "$i";
            $pad = "0$pad" while length $pad < 3;
            if (m#/day#) {
                push @termsWithDays, "A2$pad*A3$pad";
            }
            else {
                push @termsNoDays, "A2$pad*A3$pad";
            }
            $args{"A2$pad"} = $adderTable->{$_};
            $args{"A3$pad"} = $volumeAfter->{$_};
        }

        $revenuesFromAdder = Arithmetic(
            name       => 'Net revenues by tariff from adder',
            rows       => $allTariffsByEndUser,
            arithmetic => '='
              . join( '+',
                @termsWithDays
                ? ( '0.01*A400*(' . join( '+', @termsWithDays ) . ')' )
                : ('0'),
                @termsNoDays ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
                : ('0'),
              ),
            arguments     => \%args,
            defaultFormat => '0softnz'
        );

        push @{ $model->{adderResults} },
          Columnset(
            name    => 'Adder',
            columns => [
                ( grep { $_ } @{$adderTable}{@$nonExcludedComponents} ),
                $revenuesFromAdder
            ]
          );

    }

    my $totalRevenuesFromAdder = GroupBy(
        name          => 'Total net revenues from adder (£/year)',
        rows          => 0,
        cols          => 0,
        source        => $revenuesFromAdder,
        defaultFormat => '0softnz'
    );

    my $siteSpecificCharges;    # not implemented in this option

    $totalRevenuesFromAdder, $siteSpecificCharges, $adderTable;

}

1;
