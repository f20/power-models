package ModelM;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others.
Copyright 2012-2013 Franck Latrémolière, Reckon LLP and others.

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

sub meavRawData {

    my ($model) = @_;
    return unless $model->{meav};
    return $model->{meav} if ref $model->{meav};
    my $rows = Labelset( name => 'MEAV rows', list => [ split /\n/, <<EOL] );
LV main overhead line km
LV service overhead
LV overhead support
LV main underground consac km
LV main underground plastic km
LV main underground paper km
LV service underground 
LV pillar indoors
LV pillar outdoors
LV board wall-mounted
LV board underground
LV fuse pole-mounted
LV fuse tower-mounted
6.6/11kV overhead open km
6.6/11kV overhead covered km
20kV overhead open pm
20kV overhead covered km
6.6/11kV overhead support
20kV overhead support
6.6/11kV underground km
20kV underground km
HV submarine km
6.6/11kV breaker pole-mounted
6.6/11kV breaker ground-mounted
6.6/11kV switch pole-mounted
6.6/11kV switch ground-mounted
6.6/11kV ring main unit
6.6/11kV other switchgear pole-mounted
6.6/11kV other switchgear ground-mounted
20kV breaker pole-mounted
20kV breaker ground-mounted
20kV switch pole-mounted
20kV switch ground-mounted
20kV ring main unit
20kV other switchgear pole-mounted
20kV other switchgear ground-mounted
6.6/11kV transformer pole-mounted
6.6/11kV transformer ground-mounted
20kV transformer pole-mounted
20kV transformer ground-mounted
33kV overhead pole line km
33kV overhead tower line km
66kV overhead pole line km
66kV overhead tower line km
33kV pole
33kV tower
66kV pole
66kV tower
33kV underground non-pressurised km
33kV underground oil km
33kV underground gas km
66kV underground non-pressurised km
66kV underground oil km
66kV underground gas km
EHV submarine km
33kV breaker indoors
33kV breaker outdoors
33kV switch ground-mounted
33kV switch pole-mounted
33kV ring main unit
33kV other switchgear
66kV breaker
66kV other switchgear
33kV transformer pole-mounted
33kV transformer ground-mounted
33kV auxiliary transformer
66kV transformer
66kV auxiliary transformer
132kV overhead pole conductor km
132kV overhead tower conductor km 
132kV pole
132kV tower
132kV tower fittings
132kV underground non-pressurised km
132kV underground oil km
132kV underground gas km
132kV submarine km
132kV breaker
132kV other switchgear
132kV transformer
132kV auxiliary transformer
132kV/EHV remote terminal unit pole-mounted
132kV/EHV remote terminal unit ground-mounted
HV remote terminal unit pole-mounted
HV remote terminal unit ground-mounted
EOL
    my @d = (
        Dataset(
            name          => 'Asset quantity',
            defaultFormat => '0hard',
            rows          => $rows,
            data          => [ map { 0 } @{ $rows->{list} } ],
            validation    => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
        ),
        Dataset(
            name          => 'Unit MEAV (£)',
            defaultFormat => '0hard',
            rows          => $rows,
            data          => [ map { 0 } @{ $rows->{list} } ],
            validation    => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
        ),
    );
    Columnset(
        name          => 'MEAV data',
        lines         => 'From sheet Calc-MEAV or Data-MEAV',
        defaultFormat => '%hard',
        number        => 1355,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables},
        columns       => \@d,
    );
    $model->{meav} = Arithmetic(
        name          => 'MEAV (£)',
        defaultFormat => 'millionsoft',
        arithmetic    => '=IV1*IV2',
        arguments     => { IV1 => $d[0], IV2 => $d[1], },
    );

}

sub meavPercentages {

    my ( $model, $allocLevelset ) = @_;

    my $meav = $model->meavRawData;
    return Dataset(
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
    ) unless $meav;

    my $mapping = Constant(
        name          => 'MEAV mapping',
        rows          => $meav->{rows},
        cols          => $allocLevelset,
        byrow         => 1,
        defaultFormat => '0con',
        data          => [
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
        ],
    );

    if ( ref( $meav->{location} ) eq 'SpreadsheetModel::Columnset' ) {
        push @{ $meav->{location}{columns} }, $mapping;
        $mapping->{location} = $meav->{location};
    }
    else {
        Columnset(
            name    => 'MEAV calculations',
            columns => [ $meav, $mapping ]
        );
    }

    my $totalMeav = SumProduct(
        name          => 'MEAV by network level (£)',
        defaultFormat => $meav->{defaultFormat},
        matrix        => $mapping,
        vector        => $meav,
    );

    Arithmetic(
        name          => 'MEAV percentages',
        defaultFormat => '%soft',
        cols          => $allocLevelset,
        arithmetic    => '=IV6/SUM(IV8_IV9)',
        arguments     => {
            IV6     => $totalMeav,
            IV8_IV9 => $totalMeav,
        },
    );

}

sub netCapexRawData {
}

sub netCapexPercentages {

    my ( $model, $allocLevelset ) = @_;

    my $netCapex = $model->netCapexRawData;
    return Dataset(
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
    ) unless $netCapex;

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

    my $meav = $model->meavRawData;
    return Dataset(
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
    ) unless $meav;

    my $lvService = Constant(
        name          => 'MEAV mapping: LV services',
        rows          => $meav->{rows},
        defaultFormat => '0con',
        data          => [ undef, 1, ( map { undef } 1 .. 4 ), 1 ],
    );

    my $lvTotal = Constant(
        name          => 'MEAV mapping: LV total',
        rows          => $meav->{rows},
        defaultFormat => '0con',
        data          => [ map { 1; } 1 .. 13 ],
    );

    if ( ref( $meav->{location} ) eq 'SpreadsheetModel::Columnset' ) {
        push @{ $meav->{location}{columns} }, $lvService, $lvTotal;
        $_->{location} = $meav->{location} foreach $lvService, $lvTotal;
    }
    else {
        Columnset(
            name    => 'MEAV calculations',
            columns => [ $meav, $lvService, $lvTotal, ]
        );
    }

    Arithmetic(
        name          => 'MEAV: ratio of LV service to LV total',
        defaultFormat => '%soft',
        cols          => $lvOnly,
        rows          => $lvServiceOnly,
        arithmetic =>
          '=SUMPRODUCT(IV2_IV3,IV4_IV5)/SUMPRODUCT(IV6_IV7,IV8_IV9)',
        arguments => {
            IV2_IV3 => $lvService,
            IV4_IV5 => $meav,
            IV6_IV7 => $lvTotal,
            IV8_IV9 => $meav,
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
