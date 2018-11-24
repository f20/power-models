package CDCM;

# Copyright 2012 Energy Networks Association Limited and others.
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

sub revenueShortfall {

    my (
        $model,                 $allTariffsByEndUser,
        $nonExcludedComponents, $daysInYear,
        $volumeData,            $revenueBefore,
        $tariffsExMatching,     $siteSpecificOperatingCost,
        $siteSpecificReplacement,
    ) = @_;

    my $revenuesSoFar;

    {
        my @termsNoDays;
        my @termsWithDays;
        my %args = ( A400 => $daysInYear );
        my $i = 1;
        foreach (@$nonExcludedComponents) {
            ++$i;
            my $pad = "$i";
            $pad = "0$pad" while length $pad < 3;
            if (m#/day#) {
                push @termsWithDays, "A2$pad*A3$pad";
            }
            else {
                push @termsNoDays, "A2$pad*A3$pad";
            }
            $args{"A2$pad"} = $tariffsExMatching->{$_};
            $args{"A3$pad"} = $volumeData->{$_};
        }
        $revenuesSoFar = Arithmetic(
            name => Label(
                'Net revenues', 'Net revenues by tariff before matching (£)'
            ),
            rows       => $allTariffsByEndUser,
            arithmetic => '='
              . join( '+',
                @termsWithDays
                ? ( '0.01*A400*(' . join( '+', @termsWithDays ) . ')' )
                : (),
                @termsNoDays ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
                : ('0'),
              ),
            arguments     => \%args,
            defaultFormat => '0soft'
          )
    };

    push @{ $model->{revenueMatching} }, $revenuesSoFar;

    my $totalRevenuesSoFar = GroupBy(
        name          => 'Total net revenues before matching (£)',
        rows          => 0,
        cols          => 0,
        source        => $revenuesSoFar,
        defaultFormat => '0soft'
    );

    my @allowedRevenueItems = (
        Dataset(
            name          => '"Allowed revenue" (£/year)',
            data          => [280e6],
            defaultFormat => '0hard',
            validation    => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
        ),
        Dataset(
            name          => '"Pass-through charges" (£/year)',
            data          => [0],
            defaultFormat => '0hard',
            validation    => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
        ),
        Dataset(
            name =>
              'Adjustment for previous year\'s under (over) recovery (£/year)',
            data          => [0],
            defaultFormat => '0hard',
        ),
    );

    my ( $allowedRevenue, $revenueFromElsewhere );

    if ( $model->{targetRevenue} ) {
        if ( $model->{targetRevenue} =~ /dcp249|dcp273|2016/i ) {
            $allowedRevenue = $model->table1001_2016;
        }
        elsif ( $model->{targetRevenue} =~ /dcp132|2012/i ) {
            $allowedRevenue = $model->table1001_2012;
        }
        elsif ( $model->{targetRevenue} =~ /single/i ) {
            $allowedRevenue = Dataset(
                name          => 'Target CDCM net revenue (£/year)',
                data          => [300e6],
                defaultFormat => '0hard',
                validation    => {
                    validate => 'decimal',
                    criteria => '>=',
                    value    => 0,
                },
                number   => 1076,
                appendTo => $model->{inputTables},
                dataset  => $model->{dataset},
                lines => 'Source: mostly forecasts and price control formulae.',
            );
        }
    }
    else {
        $allowedRevenue = Arithmetic(
            name => 'Target net income from all use of system charges (£/year)',
            arithmetic    => '=A1+A2+A3',
            defaultFormat => '0soft',
            arguments =>
              { map { ( "A$_" => $allowedRevenueItems[ $_ - 1 ] ) } 1 .. 3 }
        );
        $model->{edcmTables}[0][4] = Stack(
            name => 'The amount of money that the DNO wants to raise from use'
              . ' of system charges (£/year)',
            sources => [$allowedRevenue],
        ) if $model->{edcmTables};
        $revenueFromElsewhere = Dataset(
            name          => 'Revenue raised outside this model (£/year)',
            defaultFormat => '0hard',
            validation    => {
                validate      => 'decimal',
                criteria      => '>',
                value         => -1e9,
                input_message => 'Enter the total net amount of revenue'
                  . ' expected from relevant charges outside this model.',
                error_message => 'Please enter a number in this cell.'
            },
            data => [5_000_000],
        );

        Columnset(
            name    => 'Target revenue',
            columns => [
                @allowedRevenueItems,
                $revenueFromElsewhere,
                $model->{adjustRevenuesBefore}
                ? @{ $model->{adjustRevenuesBefore} }
                : (),
            ],
            number   => 1076,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            lines    => 'Source: mostly forecasts and price control formulae.'
        );

    }

    my ( $revenueShortfall, $totalSiteSpecificReplacement,
        $totalSiteSpecificOperating );

    if ($siteSpecificReplacement) {
        $totalSiteSpecificReplacement = GroupBy(
            name          => 'Revenue from site specific replacement (£)',
            source        => $siteSpecificReplacement,
            defaultFormat => '0soft'
        );
        $totalSiteSpecificOperating = GroupBy(
            name          => 'Revenue from site specific operation (£)',
            source        => $siteSpecificOperatingCost,
            defaultFormat => '0soft'
        );
        $revenueShortfall = Arithmetic(
            name       => 'Revenue shortfall (surplus) £',
            rows       => 0,
            arithmetic => '=A1'
              . ( $revenueFromElsewhere ? '-A3' : '' )
              . '-A4-A5-A2'
              . ( $revenueBefore ? '-A9' : '' ),
            arguments => {
                A1 => $allowedRevenue,
                $revenueFromElsewhere ? ( A3 => $revenueFromElsewhere ) : (),
                A4 => $totalSiteSpecificReplacement,
                A5 => $totalSiteSpecificOperating,
                A2 => $totalRevenuesSoFar,
                $revenueBefore ? ( A9 => $revenueBefore ) : (),
            },
            defaultFormat => '0soft'
        );
    }
    else {
        $revenueShortfall = Arithmetic(
            name       => 'Revenue shortfall (surplus) £',
            rows       => 0,
            arithmetic => '=A1'
              . ( $revenueFromElsewhere ? '-A3' : '' ) . '-A2'
              . ( $revenueBefore        ? '-A9' : '' ),
            arguments => {
                A1 => $allowedRevenue,
                $revenueFromElsewhere ? ( A3 => $revenueFromElsewhere ) : (),
                A2 => $totalRevenuesSoFar,
                $revenueBefore ? ( A9 => $revenueBefore ) : (),
            },
            defaultFormat => '0soft'
        );
    }

    push @{ $model->{revenueMatching} },
      Columnset(
        name    => 'Revenue surplus or shortfall',
        columns => [
            $totalRevenuesSoFar,
            !$revenueBefore ? ()
            : $revenueBefore->{appendTo} ? Stack( sources => [$revenueBefore] )
            : $revenueBefore,
            $totalSiteSpecificReplacement
            ? ( $totalSiteSpecificReplacement, $totalSiteSpecificOperating )
            : (),
            $revenueShortfall,
        ],
      );

    $revenueShortfall, $totalRevenuesSoFar, $revenuesSoFar, $allowedRevenue,
      $revenueFromElsewhere, $totalSiteSpecificReplacement;

}

1;
