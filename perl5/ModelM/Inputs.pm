package ModelM;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others. All rights reserved.

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

sub splits {

    my ( $model, $allocLevelset ) = @_;

    my $lvSplit = Dataset(
        name          => 'DNO LV main usage',
        data          => [ [0.3] ],
        defaultFormat => '%hard',
        number        => 1301,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    );

    my $hvSplit = Dataset(
        name          => 'DNO HV main usage',
        data          => [ [0.3] ],
        defaultFormat => '%hard',
        number        => 1302,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    );

    $lvSplit, $hvSplit;

}

sub checks {

    my ( $model, $allocLevelset ) = @_;

    my $discounts = Columnset(
        name    => 'Current LDNO discounts',
        columns => [
            Dataset(
                name          => 'LDNO LV: LV user',
                data          => [ ['#VALUE!'] ],
                defaultFormat => '%hard',
                dataset       => $model->{dataset},
                validation    => {
                    validate => 'decimal',
                    criteria => '>=',
                    value    => 0,
                },
            ),
            Dataset(
                name          => 'LDNO HV: LV user',
                data          => [ ['#VALUE!'] ],
                defaultFormat => '%hard',
                dataset       => $model->{dataset},
                validation    => {
                    validate => 'decimal',
                    criteria => '>=',
                    value    => 0,
                },
            ),
            Dataset(
                name          => 'LDNO HV: LV Sub user',
                data          => [ ['#VALUE!'] ],
                defaultFormat => '%hard',
                dataset       => $model->{dataset},
                validation    => {
                    validate => 'decimal',
                    criteria => '>=',
                    value    => 0,
                },
            ),
            Dataset(
                name          => 'LDNO HV: HV user',
                data          => [ ['#VALUE!'] ],
                defaultFormat => '%hard',
                dataset       => $model->{dataset},
                validation    => {
                    validate => 'decimal',
                    criteria => '>=',
                    value    => 0,
                },
            ),
        ],
        number   => 1399,
        dataset  => $model->{dataset},
        appendTo => $model->{inputTables},
    );

    $discounts;

}

sub totalDpcr {

    my ($model) = @_;

    my @columns = (
        Dataset(
            name          => 'Aggregate return',
            data          => [ [100] ],
            defaultFormat => '0hard',
            dataset       => $model->{dataset},
            validation    => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
        ),
        Dataset(
            name          => 'Aggregate depreciation',
            data          => [ [100] ],
            defaultFormat => '0hard',
            dataset       => $model->{dataset},
            validation    => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
        ),
        Dataset(
            name          => 'Aggregate operating',
            data          => [ [100] ],
            defaultFormat => '0hard',
            dataset       => $model->{dataset},
            validation    => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
        ),
    );

    Columnset(
        name     => 'DPCR4 aggregate allowances',
        lines    => 'From sheet Final Allocation, cells D45, D46, D47',
        columns  => \@columns,
        number   => 1310,
        dataset  => $model->{dataset},
        appendTo => $model->{inputTables},
    );

    @columns;

}

sub oneYearDpcr {

    my ($model) = @_;

    my @columns = (
        Dataset(
            name          => 'Total revenue',
            data          => [ [100] ],
            defaultFormat => '0hard',
            dataset       => $model->{dataset},
            validation    => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
        ),
        Dataset(
            name          => 'Net incentive revenue',
            data          => [ [100] ],
            defaultFormat => '0hard',
            dataset       => $model->{dataset},
        ),
        1 ? () : Dataset(
            name          => 'Pension deficit payment',
            data          => [ [100] ],
            defaultFormat => '0hard',
            dataset       => $model->{dataset},
            validation    => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
        ),
    );

    Columnset(
        name     => 'Analysis of allowed revenue for 2007/2008',
        lines    => 'From sheet Final Allocation, cells G64, G61, G66',
        columns  => \@columns,
        number   => 1315,
        dataset  => $model->{dataset},
        appendTo => $model->{inputTables},
    );

    @columns;
}

sub units {

    my ( $model, $allocLevelset ) = @_;

    Dataset(
        name          => 'Units flowing',
        lines         => 'From sheet Calc-Units, cells C21, C21, D21, E21',
        data          => [ map { 100 } @{ $allocLevelset->{list} } ],
        defaultFormat => '0hard',
        number        => 1320,
        cols          => $allocLevelset,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    );

}

sub allocated {

    my ( $model, $allocLevelset, $expenditureSet, ) = @_;

    Dataset(
        name => 'Allocated costs',
        lines =>
'From sheet Opex Allocation, starting at cell H6, reversing column order',
        data => [
            map {
                [ map { 0 } @{ $expenditureSet->{list} } ]
            } @{ $allocLevelset->{list} }
        ],
        defaultFormat => '0hard',
        number        => 1330,
        cols          => $allocLevelset,
        rows          => $expenditureSet,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables},
    );

}

sub expenditure {

    my ( $model, $expenditureSet ) = @_;

    Dataset(
        name          => 'Total costs',
        lines         => 'From sheet Opex Allocation, starting at cell D8',
        data          => [ map { 0 } @{ $expenditureSet->{list} } ],
        defaultFormat => '0hard',
        number        => 1335,
        rows          => $expenditureSet,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables},
    );

}

sub meavPercentages {

    my ( $model, $allocLevelset ) = @_;

    Dataset(
        name => 'MEAV percentages',
        lines =>
          'From sheet Calc-Drivers row 22 or Calc-MEAV starting at cell H6',
        data          => [ map { 0 } @{ $allocLevelset->{list} } ],
        defaultFormat => '%hard',
        number        => 1350,
        cols          => $allocLevelset,
        dataset       => $model->{dataset},
        appendTo   => $model->{inputTables},
        validation => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    );

}

sub netCapexPercentages {

    my ( $model, $allocLevelset ) = @_;

    Dataset(
        name => 'Net capex percentages',
        lines =>
'From sheet Calc-Drivers row 17 or Calc-Net capex starting at cell H6',
        data          => [ map { 0 } @{ $allocLevelset->{list} } ],
        defaultFormat => '%hard',
        number        => 1370,
        cols          => $allocLevelset,
        dataset       => $model->{dataset},
        appendTo   => $model->{inputTables},
        validation => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    );

}

sub networkLengthPercentages {

    my ( $model, $allocLevelset ) = @_;

    Dataset(
        name => 'Network length percentages',
        lines =>
'From sheet Calc-Drivers row 17 or Calc-Net capex starting at cell H6',
        data          => [ map { 0 } @{ $allocLevelset->{list} } ],
        defaultFormat => '%hard',
        number        => 1375,
        cols          => $allocLevelset,
        dataset       => $model->{dataset},
        appendTo   => $model->{inputTables},
        validation => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    );

}

sub customerNumbersPercentages {

    my ( $model, $allocLevelset ) = @_;

    Dataset(
        name          => 'Customer numbers percentages',
        data          => [ map { 0 } @{ $allocLevelset->{list} } ],
        defaultFormat => '%hard',
        number        => 1377,
        cols          => $allocLevelset,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    );

}

sub meavPercentageServiceLV {

    my ( $model, $lvOnly, $lvServiceOnly ) = @_;

    Dataset(
        name          => 'MEAV: ratio of LV service to LV total',
        data          => [.5],
        defaultFormat => '%hard',
        number        => 1360,
        cols          => $lvOnly,
        rows          => $lvServiceOnly,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    );

}

sub netCapexPercentageServiceLV {

    my ( $model, $lvOnly, $lvServiceOnly ) = @_;

    Dataset(
        name          => 'Net capex: ratio of LV service to LV total',
        data          => [.5],
        defaultFormat => '%hard',
        number        => 1380,
        cols          => $lvOnly,
        rows          => $lvServiceOnly,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    );

}

sub networkLengthPercentageServiceLV {

    my ( $model, $lvOnly, $lvServiceOnly ) = @_;

    Dataset(
        name          => 'Network length: ratio of LV service to LV total',
        data          => [.5],
        defaultFormat => '%hard',
        number        => 1385,
        cols          => $lvOnly,
        rows          => $lvServiceOnly,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    );

}

sub customerNumbersPercentageServiceLV {

    my ( $model, $lvOnly, $lvServiceOnly ) = @_;

    Dataset(
        name          => 'Customer numbers: ratio of LV service to LV total',
        data          => [.5],
        defaultFormat => '%hard',
        number        => 1387,
        cols          => $lvOnly,
        rows          => $lvServiceOnly,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    );

}

1;
