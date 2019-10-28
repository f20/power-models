package ModelM::Hybridise;

# Copyright 2017-2019 Franck Latrémolière, Reckon LLP and others.
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
use SpreadsheetModel::Book::DerivativeDatasetMaker;

sub _combinedRules {
    my ( $ruleset, @amendments ) = @_;
    return $ruleset unless @amendments;
    my %newRuleset = %$ruleset;
    delete $newRuleset{$_} foreach @amendments;
    $newRuleset{AdditionalRules} =
      [ map { _combinedRules( $_, @amendments ); }
          @{ $newRuleset{AdditionalRules} } ]
      if $newRuleset{AdditionalRules};
    \%newRuleset;
}

sub process {

    my ( $self, $maker, $arg ) = @_;

    $arg =
        '1301+1302,1310,1315,1321+1322,1328,1329'
      . ',1330,1331+1332,1335,1355,1369,1380'
      unless $arg =~ /[0-9]/;
    $arg =~ s/^,+//s;
    my %allRuleChanges;
    my @hybridisationRules = map {
        my $hybridNickname;
        $hybridNickname = $1 if s/\[(.*)\]//s;
        my @changes = /([a-zA-Z0-9]+)/g;
        undef $allRuleChanges{$_} foreach grep { !/^[0-9]+$/s; } @changes;
        $hybridNickname ||=
          @changes > 1
          ? 'Tables ' . join( '&', @changes )
          : "Table @changes";
        [ $hybridNickname, @changes ];
    } split /,/, $arg;

    $maker->{setting}->(
        customPicker => sub {

            my ( $addToList, $datasetsRef, $rulesetsRef ) = @_;

            my %datasetsByDno;
            require SpreadsheetModel::Data::DnoAreas;
            push @{
                $datasetsByDno{
                    SpreadsheetModel::Data::DnoAreas::normaliseDnoName(
                        $_->{'~datasetName'} =~ m#(.*)-20[0-9]{2}-[0-9]+#
                    )
                }
              },
              $_
              foreach @$datasetsRef;

            foreach my $dno ( keys %datasetsByDno ) {

                my %sourceModelsIds = (
                    baseline => (
                        $datasetsByDno{$dno}[0]{'~datasetId'} = "Baseline $dno"
                    ),
                    scenario => (
                        $datasetsByDno{$dno}[1]{'~datasetId'} = "Scenario $dno"
                    ),
                );

                $addToList->(
                    $datasetsByDno{$dno}[0],
                    _combinedRules( $rulesetsRef->[0], keys %allRuleChanges )
                );

                $addToList->( $datasetsByDno{$dno}[1], $rulesetsRef->[0] );

                my @tableAccumulator     = (1300);
                my %remainingRuleChanges = %allRuleChanges;

                foreach (@hybridisationRules) {
                    my ( $hybridNickname, @changes ) = @$_;
                    push @tableAccumulator, grep { /^[0-9]+$/s; } @changes;
                    delete $remainingRuleChanges{$_}
                      foreach grep { !/^[0-9]+$/s; } @changes;
                    my @tableAccumulatorForClosure = @tableAccumulator;
                    $addToList->(
                        {
                            '~datasetName' =>
                              $datasetsByDno{$dno}[1]{'~datasetName'}
                              . " ($hybridNickname)",
                            dataset => {
                                1300 => [
                                    {},
                                    {},
                                    {},
                                    {
                                        'Company charging year data version' =>
                                          $hybridNickname
                                    }
                                ],
                                sourceModelsIds => \%sourceModelsIds,
                                datasetCallback => sub {
                                    my ($model) = @_;
                                    SpreadsheetModel::Book::DerivativeDatasetMaker
                                      ->applySourceModelsToDataset(
                                        $model,
                                        {
                                            baseline =>
                                              $model->{sourceModels}{baseline},
                                            map {
                                                ( $_ => $model->{sourceModels}
                                                      {scenario} );
                                            } @tableAccumulatorForClosure
                                        }
                                      );
                                },
                            }
                        },
                        _combinedRules(
                            $rulesetsRef->[0],
                            keys %remainingRuleChanges
                        ),
                    );
                }

            }

        }
    );

}

1;
