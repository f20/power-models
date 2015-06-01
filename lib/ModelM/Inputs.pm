package ModelM;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others.
Copyright 2012-2015 Franck Latrémolière, Reckon LLP and others.

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

sub lvSplit {
    my ($model) = @_;
    $model->{objects}{lvSplit} ||= Dataset(
        name  => 'DNO LV mains usage',
        lines => 'DNO-specific LV mains split'
          . ' calculated in accordance with paragraph 114 of Schedule 16',
        data => [ [0.1] ],
        defaultFormat => '%hard',
        number        => 1301,
        dataset       => $model->{dataset},
        appendTo      => $model->{objects}{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    );
}

sub hvSplit {
    my ($model) = @_;
    $model->{objects}{hvSplit} ||= Dataset(
        name  => 'DNO HV mains usage',
        lines => 'HV mains usage value provided each year'
          . ' by the Nominated Calculation Agent',
        data => [ [0.4] ],
        defaultFormat => '%hard',
        number        => 1302,
        dataset       => $model->{dataset},
        appendTo      => $model->{objects}{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    );
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
        appendTo => $model->{objects}{inputTables},
    );
    $discounts;
}

sub totalDpcr {
    my ($model) = @_;
    return @{ $model->{objects}{totalDcpr}{columns} }
      if $model->{objects}{totalDcpr};
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
    $model->{objects}{totalDcpr} = Columnset(
        name  => 'DPCR4 aggregate allowances (£)',
        lines => [
                'In a legacy Method M workbook, these data are on'
              . ' sheet Calc-Allocation, possibly cells C47, C48, C49.',
        ],
        columns  => \@columns,
        number   => 1310,
        dataset  => $model->{dataset},
        appendTo => $model->{objects}{inputTables},
    );
    @columns;
}

sub oneYearDpcr {
    my ($model) = @_;
    return @{ $model->{objects}{oneYearDpcr}{columns} }
      if $model->{objects}{oneYearDpcr};
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
    $model->{objects}{oneYearDpcr} = Columnset(
        name  => 'Analysis of allowed revenue for 2007/2008',
        lines => 'In a legacy Method M workbook, these data are on'
          . ' sheet Calc-Allocation, possibly cells F66 and F63.',
        columns  => \@columns,
        number   => 1315,
        dataset  => $model->{dataset},
        appendTo => $model->{objects}{inputTables},
    );
    @columns;
}

sub allocated {
    my ( $model, $allocLevelset, $expenditureSet, ) = @_;
    $model->{objects}{allocated}{ 0 + $allocLevelset }{ 0 + $expenditureSet }
      ||= Dataset(
        name => 'Allocated costs',
        lines =>
          'In a legacy Method M workbook, these data are on sheet Calc-Opex, '
          . 'possibly starting at cell H7, reversing column order.',
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
        appendTo      => $model->{objects}{inputTables},
      );
}

sub expenditure {
    my ( $model, $expenditureSet ) = @_;
    $model->{objects}{expenditure}{ 0 + $expenditureSet } ||= Dataset(
        name  => 'Total costs',
        lines => 'In a legacy Method M workbook, these data are on'
          . ' sheet Calc-Opex, starting at cell D7.',
        data          => [ map { 0 } @{ $expenditureSet->{list} } ],
        defaultFormat => '0hard',
        number        => 1335,
        rows          => $expenditureSet,
        dataset       => $model->{dataset},
        appendTo => $model->{objects}{inputTables},
    );
}

sub networkLengthPercentages {
    my ( $model, $allocLevelset ) = @_;
    $model->{objects}{networkLengthPercentages}{ 0 + $allocLevelset } ||=
      Dataset(
        name          => 'Network length percentages',
        data          => [ map { 0 } @{ $allocLevelset->{list} } ],
        defaultFormat => '%hard',
        number        => 1375,
        cols          => $allocLevelset,
        dataset       => $model->{dataset},
        appendTo      => $model->{objects}{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
      );
}

sub customerNumbersPercentages {
    my ( $model, $allocLevelset ) = @_;
    $model->{objects}{customerNumbersPercentages}{ 0 + $allocLevelset } ||=
      Dataset(
        name          => 'Customer numbers percentages',
        data          => [ map { 0 } @{ $allocLevelset->{list} } ],
        defaultFormat => '%hard',
        number        => 1377,
        cols          => $allocLevelset,
        dataset       => $model->{dataset},
        appendTo      => $model->{objects}{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
      );
}

sub networkLengthPercentageServiceLV {
    my ( $model, $lvOnly, $lvServiceOnly ) = @_;
    $model->{objects}{networkLengthPercentageServiceLV}{ 0 + $lvOnly }
      { 0 + $lvServiceOnly } ||= Dataset(
        name          => 'Network length: ratio of LV services to LV total',
        data          => [.5],
        defaultFormat => '%hard',
        number        => 1385,
        cols          => $lvOnly,
        rows          => $lvServiceOnly,
        dataset       => $model->{dataset},
        appendTo      => $model->{objects}{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
      );
}

sub customerNumbersPercentageServiceLV {
    my ( $model, $lvOnly, $lvServiceOnly ) = @_;
    $model->{objects}{customerNumbersPercentageServiceLV}{ 0 + $lvOnly }
      { 0 + $lvServiceOnly } ||= Dataset(
        name          => 'Customer numbers: ratio of LV services to LV total',
        data          => [.5],
        defaultFormat => '%hard',
        number        => 1387,
        cols          => $lvOnly,
        rows          => $lvServiceOnly,
        dataset       => $model->{dataset},
        appendTo      => $model->{objects}{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
      );
}

1;
