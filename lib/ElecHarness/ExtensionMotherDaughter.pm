package ElecHarness::ExtensionMotherDaughter;

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

sub process {
    my ( $self, $maker, $options ) = @_;
    $maker->{setting}->(
        customPicker => sub {
            my ( $addToList, $datasetsRef, $rulesetsRef ) = @_;
            my ( $datasetHarness, $datasetMother, $datasetDaughter ) =
              @$datasetsRef;
            my ( $rulesetHarness, $rulesetMother, $rulesetDaughter ) =
              @$rulesetsRef;
            my $datasetName = $datasetHarness->{'~datasetName'};
            $_->{'~datasetName'} = $datasetName
              foreach $datasetMother, $datasetDaughter;
            $datasetMother->{modelNumberSuffix}   = '.M';
            $datasetMother->{modelSheetPriority}  = 100;
            $datasetDaughter->{modelNumberSuffix} = '.D';
            $_->{'template'}                      = '%' foreach @$rulesetsRef;
            $datasetDaughter->{'~datasetId'}      = '_daughter';
            $datasetDaughter->{dataset}{sourceModelsIds} =
              { mother => $datasetMother->{'~datasetId'} = '_mother', };
            $datasetDaughter->{dataset}{datasetCallback} = sub {
                my ($model) = @_;
                require SpreadsheetModel::Book::DerivativeDatasetMaker;
                SpreadsheetModel::Book::DerivativeDatasetMaker
                  ->applySourceModelsToDataset(
                    $model,
                    { baseline => $model->{sourceModels}{mother} },
                    $model->{customActionMap}
                  );
            };
            $addToList->( $datasetHarness,  $rulesetHarness );
            $addToList->( $datasetMother,   $rulesetMother );
            $addToList->( $datasetDaughter, $rulesetDaughter );
        }
    );
}

1;
