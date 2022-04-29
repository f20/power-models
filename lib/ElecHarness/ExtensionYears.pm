package ElecHarness::ExtensionYears;

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

# This module (pmod -ext=ElecHarness::ExtensionYears,-7) extends a model that has
# input data table 1501 "Period covered by the model" forward or back in time.

use warnings;
use strict;
use SpreadsheetModel::Book::DerivativeDatasetMaker;

sub process {

    my ( $self, $maker, $options ) = @_;
    my $numExtraYears = $options =~ /([0-9]+)/ ? $1  : 2;
    my $direction     = $options =~ /-/        ? '-' : '+';

    my $customPicker = sub {

        my ( $addToList, $datasetsRef, $rulesetsRef ) = @_;
        my ( $rulesetHarness, $ruleset ) = @$rulesetsRef;
        my $rulesetCompact = {
            %$ruleset,
            compact => 1,
            noLinks => 1,
        };
        my ( $datasetObj, $datasetObjRef ) = @$datasetsRef;
        my $datasetName = $datasetObj->{'~datasetName'};
        $datasetObj->{'~datasetId'} = $rulesetHarness->{datasetIdsForRuns}[
            $direction eq '-'
          ? $numExtraYears
          : 0
        ] = 'Run 0';
        $datasetObj->{modelNumberSuffix} = '';

        $addToList->( $datasetObj, $rulesetHarness );
        $addToList->( $datasetObj, $ruleset );

        if ($datasetObjRef) {
            $datasetObjRef->{'~datasetId'} =
              $rulesetHarness->{datasetIdsForBaseline}[
                $direction eq '-'
              ? $numExtraYears
              : 0
              ] = 'Baseline 0';
            $datasetObjRef->{modelNumberSuffix} = '.B';
            $addToList->( $datasetObjRef, $rulesetCompact );
        }

        my $datasetCallback = sub {
            my ($model) = @_;
            SpreadsheetModel::Book::DerivativeDatasetMaker
              ->applySourceModelsToDataset(
                $model,
                {
                    previous => $model->{sourceModels}{previous}
                },
                {
                    1501 => sub {
                        my ($cell) = @_;
                        defined $cell
                          ? "=DATE(YEAR($cell)${direction}1,"
                          . "MONTH($cell),DAY($cell))"
                          : undef;
                    }
                },
              );
        };

        foreach my $datasetId ( 1 .. $numExtraYears ) {
            $addToList->(
                {
                    '~datasetId' => (
                        $rulesetHarness->{datasetIdsForRuns}[
                          $direction eq '-'
                        ? $numExtraYears - $datasetId
                        : $datasetId
                          ]
                          = 'Run ' . $datasetId
                    ),
                    '~datasetName'    => $datasetName,
                    modelNumberSuffix => $direction . $datasetId,
                    dataset           => {
                        sourceModelsIds =>
                          { previous => 'Run ' . ( $datasetId - 1 ) },
                        datasetCallback => $datasetCallback,
                    },
                },
                $rulesetCompact,
            );
            $addToList->(
                {
                    '~datasetId' => (
                        $rulesetHarness->{datasetIdsForBaseline}[
                          $direction eq '-'
                        ? $numExtraYears - $datasetId
                        : $datasetId
                          ]
                          = 'Baseline ' . $datasetId
                    ),
                    '~datasetName'    => $datasetName,
                    modelNumberSuffix => $direction . $datasetId . '.B',
                    dataset           => {
                        sourceModelsIds =>
                          { previous => 'Baseline ' . ( $datasetId - 1 ) },
                        datasetCallback => $datasetCallback,
                    },
                },
                $rulesetCompact,
            ) if $datasetObjRef;
        }
    };

    $maker->{setting}->( customPicker => $customPicker );

}

1;
