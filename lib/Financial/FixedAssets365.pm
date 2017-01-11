package Financial::FixedAssets365;

=head Copyright licence and disclaimer

Copyright 2015 Franck Latrémolière, Reckon LLP and others.

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
use base 'Financial::FixedAssets';

sub netValue {
    my ( $assets, $periods ) = @_;
    $assets->{netValue}{ 0 + $periods } ||= GroupBy(
        name          => $periods->decorate('Fixed assets (£)'),
        cols          => $periods->labelset,
        defaultFormat => '0soft',
        source        => Arithmetic(
            name => $periods->decorate('Details of net asset value (£)'),
            defaultFormat => '0soft',
            rows          => $assets->labelset,
            cols          => $periods->labelset,
            arithmetic    => '=IF(A203,IF(A351,A601+(A501-A602)*MAX(0,'
              . '1-IF(A702,(A901+1-A201)/A701/365,0)),0),0)',
            arguments => {
                A201 => $assets->comDate,
                A203 => $assets->comDate,
                A351 => $assets->grossValue($periods)->{source},
                A501 => $assets->cost,
                A601 => $assets->scrapValuation,
                A602 => $assets->scrapValuation,
                A701 => $assets->life,
                A702 => $assets->life,
                A901 => $periods->lastDay,
            }
        ),
    );
}

sub depreciationEndDate {
    my ($assets) = @_;
    $assets->{depreciationEndDate} ||= Arithmetic(
        name => 'Expected depreciation end date at time of commissioning',
        defaultFormat => 'datesoft',
        arithmetic    => '=A1+365*A2',
        arguments     => {
            A1 => $assets->comDate,
            A2 => $assets->life,
        }
    );
}

sub depreciationCharge {
    my ( $assets, $periods ) = @_;
    $assets->{depreciationCharge}{ 0 + $periods } ||= GroupBy(
        name          => $periods->decorate('Depreciation charge (£)'),
        cols          => $periods->labelset,
        defaultFormat => '0soft',
        source        => Arithmetic(
            name => $periods->decorate('Details of depreciation charge (£)'),
            defaultFormat => '0soft',
            arithmetic    => '=IF(A702,MAX(0,A1+1-A2)/A701/365*(A601-A501),0)',
            arguments     => {
                A1 => Arithmetic(
                    name          => 'End of depreciation period',
                    defaultFormat => 'datesoft',
                    rows          => $assets->labelset,
                    cols          => $periods->labelset,
                    arithmetic =>
                      '=IF(A302,MIN(A901,A301,A251),MIN(A902,A252))',
                    arguments => {
                        A251 => $assets->depreciationEndDate,
                        A252 => $assets->depreciationEndDate,
                        A301 => $assets->decomDate,
                        A302 => $assets->decomDate,
                        A702 => $assets->life,
                        A901 => $periods->lastDay,
                        A902 => $periods->lastDay,
                    },
                ),
                A2 => Arithmetic(
                    name          => 'Start of depreciation period',
                    defaultFormat => 'datesoft',
                    rows          => $assets->labelset,
                    cols          => $periods->labelset,
                    arithmetic    => '=MAX(A202,A802)',
                    arguments     => {
                        A202 => $assets->comDate,
                        A802 => $periods->firstDay,
                    },
                ),
                A501 => $assets->cost,
                A601 => $assets->scrapValuation,
                A701 => $assets->life,
                A702 => $assets->life,
            }
        ),
    );
}

sub disposalGainLoss {
    my ( $assets, $periods ) = @_;
    $assets->{disposalGainLoss}{ 0 + $periods } ||= GroupBy(
        name => $periods->decorate('Gain/loss on fixed asset disposal (£)'),
        cols => $periods->labelset,
        defaultFormat => '0soft',
        source        => Arithmetic(
            name => $periods->decorate(
                'Details of gain/loss on disposal of fixed assets (£)'),
            defaultFormat => '0soft',
            rows          => $assets->labelset,
            cols          => $periods->labelset,
            arithmetic    => '=IF(OR(NOT(A203),A301<A801,A302>A901),0,'
              . 'A651-A601-(A501-A602)*MAX(0,'
              . '1-IF(A702,(A303+1-A201)/A701/365,0)))',
            arguments => {
                A201 => $assets->comDate,
                A203 => $assets->comDate,
                A301 => $assets->decomDate,
                A302 => $assets->decomDate,
                A303 => $assets->decomDate,
                A501 => $assets->cost,
                A601 => $assets->scrapValuation,
                A602 => $assets->scrapValuation,
                A651 => $assets->scrappedValue,
                A701 => $assets->life,
                A702 => $assets->life,
                A801 => $periods->firstDay,
                A901 => $periods->lastDay,
            }
        ),
    );
}

1;
