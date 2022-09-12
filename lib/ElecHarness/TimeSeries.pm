package ElecHarness;

# Copyright 2022 Franck Latrémolière and others.
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

sub useForTimeSeries {
    my ( $me, $model, $dataName, $dataTable ) = @_;
    my $dataId           = $dataName // $dataTable->objectShortName;
    my $rowsetForSummary = $me->{rowsets}{$dataId};
    $rowsetForSummary = $me->{rowsets}{$dataId} = Labelset(
        list => [ map { "$dataName: $_" } @{ $dataTable->{rows}{list} } ], )
      if $dataTable->{rows} && !$rowsetForSummary;
    push @{ $rowsetForSummary->{accepts} }, $dataTable->{rows}
      if $rowsetForSummary;
    $me->{modelNumberSuffixMap}{ $model->{'~datasetId'} } =
      $model->{modelNumberSuffix}
      ? "TM$model->{modelNumberSuffix}"
      : 'Main model';
    if ( $dataTable->{cols} ) {
        $me->_add(
            $model,
            ( $dataName ? "$dataName: " : '' ) . $_,
            Stack(
                cols    => Labelset( list => [$_] ),
                rows    => $rowsetForSummary,
                sources => [$dataTable]
            )
        ) foreach @{ $dataTable->{cols}{list} };
    }
    else {
        my $defaultFormat = $dataTable->{defaultFormat} || '0.000soft';
        $defaultFormat =~ s/(hard|soft|con)/copy/ unless ref $defaultFormat;
        $me->_add(
            $model,
            $dataName,
            Arithmetic(
                rows          => $rowsetForSummary,
                arithmetic    => '=A1',
                arguments     => { A1 => $dataTable },
                defaultFormat => $defaultFormat,
            )
        );
    }
}

sub _add {
    my ( $me, $model, $dataName, $clone ) = @_;
    push @{ $me->{seriesList} }, $dataName unless $me->{$dataName};
    $me->{$dataName}{ $model->{'~datasetId'} } = $clone;
}

sub tsTables {

    my ($me) = @_;
    my @tables;

    push @tables, Notes( name => 'Time series' );

    push @tables, Columnset(
        name          => '',
        noHeaders     => 1,
        singleRowName => 'Model',
        columns       => [
            map {
                Constant(
                    defaultFormat => 'puretextcopy',
                    data          => [ [ $me->{modelNumberSuffixMap}{$_} ] ]
                );
            } @{ $me->{datasetIdsForRuns} }
        ],
    );

    push @tables, Columnset(
        name          => '',
        noHeaders     => 1,
        singleRowName => 'Baseline model',
        columns       => [
            map {
                Constant(
                    defaultFormat => 'puretextcopy',
                    data          => [ [ $me->{modelNumberSuffixMap}{$_} ] ]
                );
            } @{ $me->{datasetIdsForBaseline} }
        ],
    ) if $me->{datasetIdsForBaseline};

    foreach my $series ( @{ $me->{seriesList} } ) {
        my @runs = @{ $me->{$series} }{ @{ $me->{datasetIdsForRuns} } };
        push @tables,
          Columnset(
            $runs[0]{rows}
            ? ( name => $me->{datasetIdsForBaseline}
                ? "Scenario — $series"
                : '' )
            : (
                name          => '',
                singleRowName => "$series",
            ),
            noHeaders => 1,
            columns   => \@runs,
          );
        if ( $me->{datasetIdsForBaseline} ) {
            my @baseline =
              @{ $me->{$series} }{ @{ $me->{datasetIdsForBaseline} } };
            push @tables,
              Columnset(
                $runs[0]{rows}
                ? ( name => "Baseline — $series" )
                : (
                    name          => '',
                    singleRowName => "$series (baseline)",
                ),
                noHeaders => 1,
                columns   => \@baseline,
              );
            if (   $runs[0]{defaultFormat}
                && $runs[0]{defaultFormat} =~ /^0(\.0+)?[a-z]/s )
            {
                my $dec = $1 // '';
                push @tables, Columnset(
                    $runs[0]{rows}
                    ? ( name => "Increment — $series" )
                    : (
                        name          => '',
                        singleRowName => "$series (increment)",
                    ),
                    noHeaders => 1,
                    columns   => [
                        map {
                            Arithmetic(
                                name          => '',
                                defaultFormat => "0${dec}soft",
                                arithmetic    => '=A1-A2',
                                arguments     =>
                                  { A1 => $runs[$_], A2 => $baseline[$_], },
                            );
                        } 0 .. $#runs
                    ]
                );
            }
            else {
                push @tables, Columnset(
                    $runs[0]{rows}
                    ? ( name => "Comparison — $series" )
                    : (
                        name          => '',
                        singleRowName => "$series (comparison)",
                    ),
                    noHeaders => 1,
                    columns   => [
                        map {
                            Arithmetic(
                                name          => '',
                                defaultFormat => 'puretextsoft',
                                arithmetic => '=IF(A1=A2,"Same","Different")',
                                arguments  =>
                                  { A1 => $runs[$_], A2 => $baseline[$_], },
                            );
                        } 0 .. $#runs
                    ]
                );
            }
        }
    }

    @tables;

}

1;
