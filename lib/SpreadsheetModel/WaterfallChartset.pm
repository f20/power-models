package SpreadsheetModel::WaterfallChartset;

# Copyright 2017-2021 Franck Latrémolière and others.
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
use SpreadsheetModel::WaterfallChart;
use SpreadsheetModel::ChartSeries;

sub tablesAndCharts {

    my ( $class, $settings, $cols, $chartTitlesMapOptional, ) = @_;
    my ( @tables, @charts );
    my $rows = $settings->{rows} || $cols->[0]{rows};
    my $csetName = ( $cols->[0]{location} || $cols->[0] )->objectShortName;

    # Column order:
    # 0 = end point
    # 1 = start point
    # 2 = first step
    # 3 = second step
    # etc

    my ( @formatting, @value_pos, @value_neg, @padding, @increase, @decrease,
        @increase_pre, @decrease_pre );

    if ( local $_ = $cols->[0]->{defaultFormat} ) {
        s/(?:hard|soft|copy)(?:|nz)$/soft/s;
        @formatting = ( defaultFormat => $_ );
    }

    my @rowIndices = $rows ? $rows->indices : 0;

    my $naFactory = sub {
        Constant(
            name => $_[0],
            @formatting,
            rows => $rows,
            data => [ map { '=NA()'; } @rowIndices ],
        );
    };

    unless ( $settings->{mergeFirstStep} ) {
        my $stepName = $cols->[1]->objectShortName;
        @value_pos = Arithmetic(
            name => $stepName,
            @formatting,
            rows       => $rows,
            arithmetic => '=MAX(A1,0)',
            arguments  => { A1 => $cols->[1] },
        );
        @value_neg = Arithmetic(
            name => $stepName,
            @formatting,
            rows       => $rows,
            arithmetic => '=MIN(A1,0)',
            arguments  => { A1 => $cols->[1] },
        );
        @padding      = $naFactory->($stepName);
        @increase     = $naFactory->($stepName);
        @decrease     = $naFactory->($stepName);
        @increase_pre = $naFactory->($stepName);
        @decrease_pre = $naFactory->($stepName);
    }

    for ( my $index = 2 ; $index <= @$cols ; ++$index ) {

        my $index_before = $index - 1;
        my $index_after  = $index % @$cols;
        my $stepName     = $cols->[$index_after]->objectShortName;

        if ( $index_after
            and !$settings->{mergeFirstStep} || $index_before > 1 )
        {
            push @value_pos, $naFactory->($stepName);
            push @value_neg, $naFactory->($stepName);
        }
        else {
            push @value_pos,
              Arithmetic(
                name => $stepName,
                @formatting,
                rows       => $rows,
                arithmetic => '=MAX(A1,0)',
                arguments =>
                  { A1 => $cols->[ $index_after ? $index_before : 0 ] },
              );
            push @value_neg,
              Arithmetic(
                name => $stepName,
                @formatting,
                rows       => $rows,
                arithmetic => '=MIN(A1,0)',
                arguments =>
                  { A1 => $cols->[ $index_after ? $index_before : 0 ] },
              );
        }

        push @padding,
          Arithmetic(
            name => $stepName,
            @formatting,
            rows => $rows,
            arithmetic =>
              '=IF(A1<0,IF(A2<0,MAX(A11,A21),0),IF(A23<0,0,MIN(A12,A22)))',
            arguments => {
                A1  => $cols->[$index_before],
                A11 => $cols->[$index_before],
                A12 => $cols->[$index_before],
                A2  => $cols->[$index_after],
                A21 => $cols->[$index_after],
                A22 => $cols->[$index_after],
                A23 => $cols->[$index_after],
            },
          );

        push @increase,
          Arithmetic(
            name => $stepName,
            @formatting,
            rows => $rows,
            arithmetic =>
              '=IF(A1<0,MIN(0,MIN(0,A21)-A11),MAX(0,A12-MAX(0,A22)))',
            arguments => {
                A1  => $cols->[$index_after],
                A11 => $cols->[$index_after],
                A12 => $cols->[$index_after],
                A21 => $cols->[$index_before],
                A22 => $cols->[$index_before],
            },
          );

        push @increase_pre,
          Arithmetic(
            name => $stepName,
            @formatting,
            rows       => $rows,
            arithmetic => '=IF(A1<0,0,MIN(0,A22))',
            arguments  => {
                A1  => $cols->[$index_after],
                A22 => $cols->[$index_before],
            },
          );

        push @decrease,
          Arithmetic(
            name => $stepName,
            @formatting,
            rows => $rows,
            arithmetic =>
              '=IF(A1<0,MIN(0,A11-MIN(0,A21)),MAX(0,MAX(0,A22)-A12))',
            arguments => {
                A1  => $cols->[$index_after],
                A11 => $cols->[$index_after],
                A12 => $cols->[$index_after],
                A21 => $cols->[$index_before],
                A22 => $cols->[$index_before],
            },
          );

        push @decrease_pre,
          Arithmetic(
            name => $stepName,
            @formatting,
            rows       => $rows,
            arithmetic => '=IF(A1<0,MAX(0,A21),0)',
            arguments  => {
                A1  => $cols->[$index_after],
                A21 => $cols->[$index_before],
            },
          );

    }

    push @tables,
      my $valueColumnset = Columnset(
        name    => "$csetName: baseline positive values",
        columns => \@value_pos,
      );
    push @tables,
      my $valueNegColumnset = Columnset(
        name    => "$csetName: baseline negative values",
        columns => \@value_neg,
      );
    push @tables,
      my $paddingColumnset = Columnset(
        name    => "$csetName: padding",
        columns => \@padding,
      );
    push @tables,
      my $increaseColumnset = Columnset(
        name    => "$csetName: increases",
        columns => \@increase,
      );
    push @tables,
      my $preincreaseColumnset = Columnset(
        name    => "$csetName: pre-increases",
        columns => \@increase_pre,
      );
    push @tables,
      my $decreaseColumnset = Columnset(
        name    => "$csetName: decreases",
        columns => \@decrease,
      );
    push @tables,
      my $predecreaseColumnset = Columnset(
        name    => "$csetName: pre-decreases",
        columns => \@decrease_pre,
      );

    for my $r (@rowIndices) {
        next if $rows && $rows->{groupid} && !defined $rows->{groupid}[$r];
        local $_ = $rows ? $rows->{list}[$r] : $csetName;
        if ($chartTitlesMapOptional) {
            $_ = $chartTitlesMapOptional->{$_};
        }
        else {
            s/.*\n//s;
            $_ = "$csetName for $_";
        }
        push @charts,
          SpreadsheetModel::WaterfallChart->new(
            name           => $_,
            scaling_factor => $settings->{scaling_factor},
            width          => $settings->{width},
            height         => $settings->{height},
            $settings->{instructions}
            ? ( instructions => [ @{ $settings->{instructions} } ] )
            : (),
            grey_rightwards =>
              SpreadsheetModel::ChartSeries->new( $valueColumnset, $r ),
            grey_leftwards =>
              SpreadsheetModel::ChartSeries->new( $valueNegColumnset, $r ),
            padding =>
              SpreadsheetModel::ChartSeries->new( $paddingColumnset, $r ),
            blue_light =>
              SpreadsheetModel::ChartSeries->new( $preincreaseColumnset, $r ),
            blue_rightwards =>
              SpreadsheetModel::ChartSeries->new( $increaseColumnset, $r ),
            orange_light =>
              SpreadsheetModel::ChartSeries->new( $predecreaseColumnset, $r ),
            orange_leftwards =>
              SpreadsheetModel::ChartSeries->new( $decreaseColumnset, $r ),
          );
    }

    \@tables, \@charts;

}

1;
