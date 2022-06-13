package GenericTester::WaterfallTester;

# Copyright 2021-2022 Franck Latrémolière and others.
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
use SpreadsheetModel::WaterfallChartset;

sub new {
    my ( $class, $model, @options ) = @_;
    bless { model => $model, @options }, $class;
}

sub rowset {
    my ($component) = @_;
    return $component->{rowset} if $component->{rowset};
    if ( $component->{rowList} ) {
        $component->{rowset} =
          Labelset( list => $component->{rowList} );
        return $component->{rowset}{rowsetInputs} = $component->{rowset};
    }
    my $rowsetInputs = Labelset(
        defaultFormat => 'thitem',
        list          => [ 1 .. ( $component->{numItems} || 7 ) ]
    );
    $component->{itemNames} //= Dataset(
        name          => 'Item names',
        defaultFormat => 'texthard',
        rows          => $rowsetInputs,
        data => [ map { 'Item ' . ( 1 + $_ ); } $rowsetInputs->indices ],
    );
    $component->{rowset} = Labelset(
        editable     => $component->{itemNames},
        accepts      => [$rowsetInputs],
        rowsetInputs => $rowsetInputs,
    );
}

sub dataColumnsRef {
    my ($component) = @_;
    $component->{dataColumnsRef} //= [
        map {
            Dataset(
                name          => $_,
                defaultFormat => '0.00hard',
                rows          => $component->rowset->{rowsetInputs},
                data          => [ map { 0; } $component->rowset->indices ],
            );
        } $component->{stepList}
        ? @{ $component->{stepList} }
        : ( 'End point', 'Start point', map { "Step $_"; } 1 .. 6 )
    ];
}

sub inputTables {
    my ($component) = @_;
    $component->{inputColumnset} //= Columnset(
        name    => 'Waterfall chart data',
        columns => [
            $component->{itemNames} ? $component->{itemNames} : (),
            @{ $component->dataColumnsRef },
        ],
        dataset => $component->{model}{dataset},
        number  => 110,
    );
}

sub waterfallSettings {
    my ($component) = @_;
    my %settings;
    %settings = %{ $component->{chartOptions} } if $component->{chartOptions};
    $settings{chartTitlesMaker} = sub { $_[0]; };
    $settings{rows}             = $component->rowset;
    \%settings;
}

sub myTablesAndCharts {
    my ($component) = @_;
    $component->{myTablesAndCharts} //= [
        SpreadsheetModel::WaterfallChartset->tablesAndCharts(
            $component->waterfallSettings,
            $component->dataColumnsRef,
        )
    ];
}

sub calcTables {
    my ($component) = @_;
    @{ $component->myTablesAndCharts->[0] };
}

sub charts {
    my ($component) = @_;
    @{ $component->myTablesAndCharts->[1] };
}

1;
