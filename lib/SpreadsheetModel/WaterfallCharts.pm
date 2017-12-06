package SpreadsheetModel::WaterfallCharts;

=head Copyright licence and disclaimer

Copyright 2017 Franck Latrémolière, Reckon LLP and others.

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
use SpreadsheetModel::WaterfallChart;
use SpreadsheetModel::ChartSeries;

sub tablesAndCharts {

    my ( $class, $settings, $cols ) = @_;
    my ( @tables, @charts );
    my $rows = $settings->{rows} || $cols->[1]{rows};
    my $csetName = $cols->[1]{location}->objectShortName;

    my $n     = $cols->[1]->objectShortName;
    my @value = (
        Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=MAX(A1,0)',
            arguments  => { A1 => $cols->[1] },
        )
    );
    my @value_neg = (
        Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=MIN(A1,0)',
            arguments  => { A1 => $cols->[1] },
        )
    );
    my @padding = (
        Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=NA()',
            arguments  => { A1 => $cols->[1] },
        )
    );
    my @increase = (
        Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=NA()',
            arguments  => { A1 => $cols->[1] },
        )
    );
    my @increase_neg = (
        Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=NA()',
            arguments  => { A1 => $cols->[1] },
        )
    );
    my @decrease = (
        Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=NA()',
            arguments  => { A1 => $cols->[1] },
        )
    );
    my @decrease_neg = (
        Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=NA()',
            arguments  => { A1 => $cols->[1] },
        )
    );

    my $limit = @$cols + ( $settings->{secondBar} ? 0 : 1 );
    for ( my $i = 2 ; $i < $limit ; ++$i ) {
        my $j = $i % @$cols;
        $n = $cols->[$j]->objectShortName;
        push @value,
          Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=NA()',
            arguments  => { A1 => $cols->[$j] },
          );
        push @value_neg,
          Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=NA()',
            arguments  => { A1 => $cols->[$j] },
          );
        push @padding,
          Arithmetic(
            name => $n,
            rows => $rows,
            arithmetic =>
              '=IF(A1<0,IF(A2<0,MAX(A11,A21),0),IF(A23<0,0,MIN(A12,A22)))',
            arguments => {
                A1  => $cols->[ $i - 1 ],
                A11 => $cols->[ $i - 1 ],
                A12 => $cols->[ $i - 1 ],
                A2  => $cols->[$j],
                A21 => $cols->[$j],
                A22 => $cols->[$j],
                A23 => $cols->[$j],
            },
          );
        push @increase,
          Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=MAX(0,MAX(A2,0)-MAX(A1,0))',
            arguments  => { A1 => $cols->[ $i - 1 ], A2 => $cols->[$j], },
          );
        push @increase_neg,
          Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=MIN(0,MIN(A1,0)-MIN(A2,0))',
            arguments  => { A1 => $cols->[ $i - 1 ], A2 => $cols->[$j], },
          );
        push @decrease,
          Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=MAX(0,MAX(A1,0)-MAX(A2,0))',
            arguments  => { A1 => $cols->[ $i - 1 ], A2 => $cols->[$j], },
          );
        push @decrease_neg,
          Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=MIN(0,MIN(A2,0)-MIN(A1,0))',
            arguments  => { A1 => $cols->[ $i - 1 ], A2 => $cols->[$j], },
          );
    }

    if ( $settings->{secondBar} ) {
        $n = $cols->[0]->objectShortName;
        push @value,
          Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=MAX(A1,0)',
            arguments  => { A1 => $cols->[0] },
          );
        push @value_neg,
          Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=MIN(A1,0)',
            arguments  => { A1 => $cols->[0] },
          );
        push @padding,
          Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=NA()',
            arguments  => { A1 => $cols->[0] },
          );
        push @increase,
          Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=NA()',
            arguments  => { A1 => $cols->[0] },
          );
        push @increase_neg,
          Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=NA()',
            arguments  => { A1 => $cols->[0] },
          );
        push @decrease,
          Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=NA()',
            arguments  => { A1 => $cols->[0] },
          );
        push @decrease_neg,
          Arithmetic(
            name       => $n,
            rows       => $rows,
            arithmetic => '=NA()',
            arguments  => { A1 => $cols->[0] },
          );
    }

    push @tables,
      my $valueColumnset = Columnset(
        name    => "$csetName: baseline positive values",
        columns => \@value,
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
        name    => "$csetName: positive increases",
        columns => \@increase,
      );
    push @tables,
      my $increaseNegColumnset = Columnset(
        name    => "$csetName: negative increases",
        columns => \@increase_neg,
      );
    push @tables,
      my $decreaseColumnset = Columnset(
        name    => "$csetName: positive decreases",
        columns => \@decrease,
      );
    push @tables,
      my $decreaseNegColumnset = Columnset(
        name    => "$csetName: negative decreases",
        columns => \@decrease_neg,
      );

    for my $r ( $rows->indices ) {
        next if $rows->{groupid} && !defined $rows->{groupid}[$r];
        local $_ = $rows->{list}[$r];
        s/.*\n//s;
        push @charts,
          SpreadsheetModel::WaterfallChart->new(
            name => "$csetName for $_",
            $settings->{instructions}
            ? ( instructions => [ @{ $settings->{instructions} } ] )
            : (),
            value => SpreadsheetModel::ChartSeries->new( $valueColumnset, $r ),
            value_neg =>
              SpreadsheetModel::ChartSeries->new( $valueNegColumnset, $r ),
            padding =>
              SpreadsheetModel::ChartSeries->new( $paddingColumnset, $r ),
            increase =>
              SpreadsheetModel::ChartSeries->new( $increaseColumnset, $r ),
            increase_neg =>
              SpreadsheetModel::ChartSeries->new( $increaseNegColumnset, $r ),
            decrease =>
              SpreadsheetModel::ChartSeries->new( $decreaseColumnset, $r ),
            decrease_neg =>
              SpreadsheetModel::ChartSeries->new( $decreaseNegColumnset, $r ),
          );
    }

    \@tables, \@charts;

}

1;
