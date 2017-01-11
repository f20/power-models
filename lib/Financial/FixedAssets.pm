package Financial::FixedAssets;

=head Copyright licence and disclaimer

Copyright 2015-2017 Franck Latrémolière, Reckon LLP and others.

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

sub new {
    my ( $class, %args ) = @_;
    die __PACKAGE__ . ' needs a model attribute' unless $args{model};
    bless \%args, $class;
}

sub finish {
    my ($assets) = @_;
    return unless $assets->{inputDataColumns};
    my @columns =
      grep { ref $_ =~ /Dataset/; }
      @{ $assets->{inputDataColumns} }
      {qw(names comDate decomDate cost life scrapValuation scrappedValue)};
    Columnset(
        name     => 'Fixed assets',
        columns  => \@columns,
        appendTo => $assets->{model}{inputTables},
        dataset  => $assets->{model}{dataset},
        number   => 1450,
    );
}

sub labelset {
    my ($assets) = @_;
    $assets->{labelset} ||= Labelset(
        name     => 'Fixed assets with names',
        editable => (
            $assets->{inputDataColumns}{names} ||= Dataset(
                name          => 'Asset name',
                defaultFormat => 'texthard',
                rows          => $assets->labelsetNoNames,
                data => [ map { '' } $assets->labelsetNoNames->indices ],
            )
        ),
    );
}

sub labelsetNoNames {
    my ($assets) = @_;
    $assets->{labelsetNoNames} ||= Labelset(
        name          => 'Fixed assets without names',
        defaultFormat => 'thitem',
        list          => [ 1 .. $assets->{model}{numAssets} || 4 ]
    );
}

sub life {
    my ($assets) = @_;
    $assets->{inputDataColumns}{life} ||= Dataset(
        name          => 'Straight-line depreciation period (years)',
        defaultFormat => '0.0hard',
        rows          => $assets->labelsetNoNames,
        data          => [ map { 0 } @{ $assets->labelsetNoNames->{list} } ],
    );
}

sub cost {
    my ($assets) = @_;
    $assets->{inputDataColumns}{cost} ||= Dataset(
        name          => 'Cost (£)',
        defaultFormat => '0hard',
        rows          => $assets->labelsetNoNames,
        data          => [ map { 0 } @{ $assets->labelsetNoNames->{list} } ],
    );
}

sub scrapValuation {
    my ($assets) = @_;
    $assets->{inputDataColumns}{scrapValuation} ||= Dataset(
        name          => 'Estimated scrap value at time of commissioning (£)',
        defaultFormat => '0hard',
        rows          => $assets->labelsetNoNames,
        data          => [ map { 0 } @{ $assets->labelsetNoNames->{list} } ],
    );
}

sub scrappedValue {
    my ($assets) = @_;
    $assets->{inputDataColumns}{scrappedValue} ||= Dataset(
        name          => 'Proceeds of scrapping (£)',
        defaultFormat => '0hard',
        rows          => $assets->labelsetNoNames,
        data          => [ map { 0 } @{ $assets->labelsetNoNames->{list} } ],
    );
}

sub comDate {
    my ($assets) = @_;
    $assets->{inputDataColumns}{comDate} ||= Dataset(
        name          => 'Commissioning date',
        defaultFormat => 'datehard',
        rows          => $assets->labelsetNoNames,
        data          => [ map { '' } @{ $assets->labelsetNoNames->{list} } ],
    );
}

sub decomDate {
    my ($assets) = @_;
    $assets->{inputDataColumns}{decomDate} ||= Dataset(
        name          => 'Withdrawal date',
        defaultFormat => 'datehard',
        rows          => $assets->labelsetNoNames,
        data          => [ map { '' } @{ $assets->labelsetNoNames->{list} } ],
    );
}

sub grossValue {
    my ( $assets, $periods ) = @_;
    $assets->{grossValue}{ 0 + $periods } ||= GroupBy(
        name          => $periods->decorate('Gross asset value (£)'),
        cols          => $periods->labelset,
        defaultFormat => '0soft',
        source        => Arithmetic(
            name => $periods->decorate('Details of gross asset value (£)'),
            defaultFormat => '0soft',
            rows          => $assets->labelset,
            cols          => $periods->labelset,
            arithmetic => '=IF(A201>A901,0,IF(A301,IF(A302>A902,A501,0),A502))',
            arguments  => {
                A201 => $assets->comDate,
                A301 => $assets->decomDate,
                A302 => $assets->decomDate,
                A501 => $assets->cost,
                A502 => $assets->cost,
                A901 => $periods->lastDay,
                A902 => $periods->lastDay,
            }
        ),
    );
}

sub assetInventory {
    my ( $assets, $periods ) = @_;
    return () unless $assets->{capitalExp};
    $assets->{assetInventory}{ 0 + $periods } ||= Arithmetic(
        name          => $periods->decorate('Assets held as inventory (£)'),
        defaultFormat => '0soft',
        arithmetic    => '=0-A1-A2',
        arguments     => {
            A1 => $assets->{capitalExp}->aggregate($periods),
            A2 => GroupBy(
                name => $periods->decorate('Net recognition to date (£)'),
                cols => $periods->labelset,
                defaultFormat => '0soft',
                source        => Arithmetic(
                    name => $periods->decorate(
                        'Details of net recognition to date (£)'),
                    defaultFormat => '0soft',
                    rows          => $assets->labelset,
                    cols          => $periods->labelset,
                    arithmetic    => '=IF(A201>A901,0,'
                      . 'IF(A301,IF(A302>A902,A501,A503-A504),A502))',
                    arguments => {
                        A201 => $assets->comDate,
                        A301 => $assets->decomDate,
                        A302 => $assets->decomDate,
                        A501 => $assets->cost,
                        A502 => $assets->cost,
                        A503 => $assets->cost,
                        A504 => $assets->scrappedValue,
                        A901 => $periods->lastDay,
                        A902 => $periods->lastDay,
                    }
                ),
            ),
        },
    );
}

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
              . '1-IF(A702,('
              . '12*(YEAR(A901)-YEAR(A201))+MONTH(A902)-MONTH(A202)+1'
              . ')/A701/12,0)),0),0)',
            arguments => {
                A201 => $assets->comDate,
                A202 => $assets->comDate,
                A203 => $assets->comDate,
                A351 => $assets->grossValue($periods)->{source},
                A501 => $assets->cost,
                A601 => $assets->scrapValuation,
                A602 => $assets->scrapValuation,
                A701 => $assets->life,
                A702 => $assets->life,
                A901 => $periods->lastDay,
                A902 => $periods->lastDay,
            }
        ),
    );
}

sub depreciationEndDate {
    my ($assets) = @_;
    $assets->{depreciationEndDate} ||= Arithmetic(
        name          => 'Depreciation end date for each asset',
        defaultFormat => 'datesoft',
        rows          => $assets->labelset,
        arithmetic    => '=IF(A5,DATE(YEAR(A1),MONTH(A2)+12*A3,1)-1,0)',
        arguments     => {
            A1 => $assets->comDate,
            A2 => $assets->comDate,
            A3 => $assets->life,
            A4 => $assets->comDate,
            A5 => $assets->life,
        }
    );
}

sub depreciationCharge {

    my ( $assets, $periods, $recache ) = @_;

    if ( my $cached = $assets->{depreciationCharge}{ 0 + $periods } ) {
        return $cached unless $recache;
        return $assets->{depreciationCharge}{ 0 + $periods } = Stack(
            name    => $cached->objectShortName,
            sources => [$cached]
        );
    }

    my $start = Arithmetic(
        name          => 'Start of depreciation period',
        defaultFormat => 'datesoft',
        rows          => $assets->labelset,
        cols          => $periods->labelset,
        arithmetic    => '=IF(A702,MAX(A202,A802),0)',
        arguments     => {
            A202 => $assets->comDate,
            A702 => $assets->life,
            A802 => $periods->firstDay,
        },
    );

    my $end = Arithmetic(
        name          => 'End of the depreciation period',
        defaultFormat => 'datesoft',
        rows          => $assets->labelset,
        cols          => $periods->labelset,
        arithmetic    => '=IF(A302,MIN(A901,A301,A251),MIN(A902,A252))',
        arguments     => {
            A251 => $assets->depreciationEndDate,
            A252 => $assets->depreciationEndDate,
            A301 => $assets->decomDate,
            A302 => $assets->decomDate,
            A901 => $periods->lastDay,
            A902 => $periods->lastDay,
        },
    );

    $assets->{depreciationCharge}{ 0 + $periods } = GroupBy(
        name          => $periods->decorate('Depreciation charge (£)'),
        cols          => $periods->labelset,
        defaultFormat => '0soft',
        source        => Arithmetic(
            name => $periods->decorate('Details of depreciation charge (£)'),
            defaultFormat => '0soft',
            arithmetic =>
              '=IF(A702,MAX(0,YEAR(A1)*12+MONTH(A2)+1-YEAR(A3)*12-MONTH(A4))'
              . '/A701/12*(A601-A501),0)',
            arguments => {
                A1   => $end,
                A2   => $end,
                A3   => $start,
                A4   => $start,
                A501 => $assets->cost,
                A601 => $assets->scrapValuation,
                A701 => $assets->life,
                A702 => $assets->life,
            }
        ),
    );

}

sub capitalExpenditure {
    my ( $assets, $periods ) = @_;
    return $assets->{capitalExpenditure}{ 0 + $periods }
      if $assets->{capitalExpenditure}{ 0 + $periods };
    return $assets->{capitalExpenditure}{ 0 + $periods } =
      $assets->{capitalExp}->stream($periods)
      if $assets->{capitalExp};
    $assets->{capitalExpenditure}{ 0 + $periods } = GroupBy(
        name          => $periods->decorate('Capital expenditure (£)'),
        cols          => $periods->labelset,
        defaultFormat => '0soft',
        source        => Arithmetic(
            name => $periods->decorate('Details of capital expenditure (£)'),
            defaultFormat => '0soft',
            rows          => $assets->labelset,
            cols          => $periods->labelset,
            arithmetic    => '=IF(OR(A201<A802,A202>A902),0,0-A501)',
            arguments     => {
                A201 => $assets->comDate,
                A202 => $assets->comDate,
                A501 => $assets->cost,
                A802 => $periods->firstDay,
                A902 => $periods->lastDay,
            }
        ),
    );
}

sub capitalReceipts {
    my ( $assets, $periods ) = @_;
    $assets->{capitalReceipts}{ 0 + $periods } ||= GroupBy(
        name => $periods->decorate('Receipts from scrapping assets (£)'),
        cols => $periods->labelset,
        defaultFormat => '0soft',
        source        => Arithmetic(
            name => $periods->decorate(
                'Details of receipts from scrapping assets (£)'),
            defaultFormat => '0soft',
            rows          => $assets->labelset,
            cols          => $periods->labelset,
            arithmetic    => '=IF(OR(A301<A801,A302>A901),0,A601)',
            arguments     => {
                A301 => $assets->decomDate,
                A302 => $assets->decomDate,
                A601 => $assets->scrappedValue,
                A801 => $periods->firstDay,
                A901 => $periods->lastDay,
            }
        ),
    );
}

sub disposalGainLoss {
    my ( $assets, $periods, $recache ) = @_;
    return $assets->{disposalGainLoss}{ 0 + $periods } = Stack(
        name    => $recache->objectShortName,
        sources => [$recache]
    ) if $recache and $recache = $assets->{disposalGainLoss}{ 0 + $periods };
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
              . '1-IF(A702,('
              . '12*(YEAR(A303)-YEAR(A201))+MONTH(A304)-MONTH(A202)+1'
              . ')/A701/12,0)))',
            arguments => {
                A201 => $assets->comDate,
                A202 => $assets->comDate,
                A203 => $assets->comDate,
                A301 => $assets->decomDate,
                A302 => $assets->decomDate,
                A303 => $assets->decomDate,
                A304 => $assets->decomDate,
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

