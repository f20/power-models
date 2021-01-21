package ElecHarness;

# Copyright 2021 Franck Latrémolière and others.
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

sub takeResults {
    my ( $me, $model, @tables ) = @_;
    unless ( $me->{modelData}{ 0 + $model } ) {
        push @{ $me->{modelList} }, $model;
        $me->{modelData}{ 0 + $model } = [];
    }
    push @{ $me->{modelData}{ 0 + $model } }, @tables;
}

sub makeWaterfallColumnsets {
    my ($me) = @_;
    my ( @columnsets, @rows, %rows );
    foreach (
        map { $_->{rows} ? @{ $_->{rows}{list} } : ''; }
        map { @{ $me->{modelData}{ 0 + $_ } }; } @{ $me->{modelList} }
      )
    {
        next if exists $rows{$_};
        push @rows, $_;
        undef $rows{$_};
    }
    push @columnsets, Columnset(
        name    => 'Total revenue for waterfall charts',
        columns => [
            map {
                Stack(
                    name => $_->{nickName} || $_,
                    defaultFormat => '0copy',
                    sources       => $me->{modelData}{ 0 + $_ },
                );
            } @{ $me->{modelList} }
        ],
    ) if exists $rows{''};
    my $rowset = Labelset( list => [ grep { $_ ne ''; } @rows ] );
    push @columnsets, Columnset(
        name    => 'Revenue for waterfall charts',
        columns => [
            map {
                Stack(
                    name => $_->{nickName} || $_,
                    defaultFormat => '0copy',
                    rows          => $rowset,
                    sources       => $me->{modelData}{ 0 + $_ },
                );
            } @{ $me->{modelList} }
        ],
    ) if @{ $rowset->{list} };
    @columnsets;
}

sub waterfallTablesAndCharts {
    my ($me) = @_;
    return $me->{waterfallTablesAndCharts} if $me->{waterfallTablesAndCharts};
    my ( @tables, @charts );
    foreach ( $me->makeWaterfallColumnsets ) {
        my ( $t, $c ) = SpreadsheetModel::WaterfallChartset->tablesAndCharts(
            {
                scaling_factor => 1.5,
                instructions   => [
                    set_x_axis => [
                        num_format => '#,##0,"k";[red](#,##0,"k");0',
                        num_font   => { size => 18 },
                    ]
                ],
            },
            $_->{columns}
        );
        push @tables, @$t;
        push @charts, @$c;
    }
    $me->{waterfallTablesAndCharts} = [ \@tables, \@charts ];
}

1;
