package ModelM;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others.
Copyright 2012-2014 Franck Latrémolière, Reckon LLP and others.

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

sub netCapexRawData {
}

sub netCapexPercentages {

    my ( $model, $allocLevelset ) = @_;

    my $netCapex = $model->netCapexRawData;
    return Dataset(
        name  => 'Net capex percentages',
        lines => 'From pre-DCP 118 sheet Calc-Net capex starting at cell H6.',
        data  => [ map { 0 } @{ $allocLevelset->{list} } ],
        defaultFormat => '%hard',
        number        => 1370,
        cols          => $allocLevelset,
        dataset       => $model->{dataset},
        appendTo      => $model->{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    ) unless $netCapex;

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

1;
