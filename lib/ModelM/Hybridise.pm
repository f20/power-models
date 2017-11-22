package ModelM::Hybridise;

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

=head About this extensions

This hybridises some of the input data, on a single set of rules.

=cut

use warnings;
use strict;

sub process {

    my ( $self, $maker, $arg ) = @_;

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

                $addToList->( $datasetsByDno{$dno}[0], $rulesetsRef->[0] );

                $addToList->( $datasetsByDno{$dno}[1], $rulesetsRef->[0] );

                my %sourceModelsDatasetNameMatches = (
                    baseline => $datasetsByDno{$dno}[0]{'~datasetName'},
                    scenario => $datasetsByDno{$dno}[1]{'~datasetName'},
                );
                my @tables;

                foreach ( [1315], [ 1321, 1322 ], [1335], ) {
                    push @tables, @$_;
                    $addToList->(
                        {
                            '~datasetName' =>
                              $datasetsByDno{$dno}[1]{'~datasetName'} . ' blah',
                            dataset => {
                                1300 => [
                                    {},
                                    {},
                                    {},
                                    {
                                        'Company charging year data version' =>
                                          "Tables @$_"
                                    }
                                ],
                                scenarioised => [@tables],
                                sourceModelsDatasetNameMatches =>
                                  \%sourceModelsDatasetNameMatches,
                            }
                        },
                        $rulesetsRef->[0],
                    );
                }

            }

        }
    );

}

sub apply {

    my ( $model, $sourceModels, $tableList ) = @_;

    my %actions;
    my %sources = %$sourceModels;
    my $copy = sub { my ($cell) = @_; "=$cell"; };
    foreach (@$tableList) {
        $sources{$_} = $sourceModels->{scenario};
        $actions{$_} = $copy;
    }

    require SpreadsheetModel::Book::Derivative;

    # Adaptation bodge for CDCM-style API
    $_->{inputTables} = $_->{objects}{inputTables}
      foreach values %$sourceModels;
    SpreadsheetModel::Book::Derivative::registerSourceModels( $model,
        \%sources, \%actions );

}

1;
