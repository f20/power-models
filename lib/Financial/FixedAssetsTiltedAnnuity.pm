package Financial::FixedAssetsTiltedAnnuity;

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

sub annuityIndex {
    my ($assets) = @_;
    $assets->{inputDataColumns}{annuityIndex} ||= Dataset(
        name          => 'Notional annuity growth rate',
        defaultFormat => '%hard',
        rows          => $assets->labelsetNoNames,
        data          => [ map { 0 } @{ $assets->labelsetNoNames->{list} } ],
    );
}

sub rateOfReturn {
    my ($assets) = @_;
    $assets->{inputDataColumns}{rateOfReturn} ||= Dataset(
        name          => 'Notional rate of return',
        defaultFormat => '%hard',
        rows          => $assets->labelsetNoNames,
        data          => [ map { 0 } @{ $assets->labelsetNoNames->{list} } ],
    );
}

sub flattenFormula {
    my $argcounter = 0;
    my %arguments;
    my $flatten;
    $flatten = sub {
        ( local $_, my %args ) = @_;
        while ( my ( $k, $v ) = each %args ) {
            my $replacement;
            if ( ref $v eq 'ARRAY' ) {
                $replacement = '(' . $flatten->(@$v) . ')';
            }
            else {
                $replacement = 'F' . ++$argcounter;
                $arguments{$replacement} = $v;
            }
            s/\b$k\b/$replacement/;
        }
        $_;
    };
    arithmetic => '=' . $flatten->(@_), arguments => \%arguments;
}

sub depreciationArithmetic {
    my ( $assets, $periods, $start, $end ) = @_;
    my @indexationFactor = ( '1+A1', A1 => $assets->annuityIndex, );
    my @combinedFactor = (
        '(1+A1)/(1+A2)',
        A1 => $assets->annuityIndex,
        A2 => $assets->rateOfReturn,
    );
    my @denominator = (
        '1-A1^A2',
        A1 => [@combinedFactor],
        A2 => $assets->life,
    );
    my @numeratorStart = (
        'A1^A2*(1-A3^A4)',
        A1 => [@indexationFactor],
        A2 => [
            'YEAR(A1)-YEAR(A3)+(MONTH(A2)-MONTH(A4))/12',
            A1 => $start,
            A2 => $start,
            A3 => $assets->comDate,
            A4 => $assets->comDate,
        ],
        A3 => [@combinedFactor],
        A4 => [
            'A5+YEAR(A1)-YEAR(A3)+(MONTH(A2)-MONTH(A4))/12',
            A1 => $assets->comDate,
            A2 => $assets->comDate,
            A3 => $start,
            A4 => $start,
            A5 => $assets->life,
        ],
    );
    my @numeratorEnd = (
        'A1^A2*(1-A3^A4)',
        A1 => [@indexationFactor],
        A2 => [
            'YEAR(A1)-YEAR(A3)+(MONTH(A2)-MONTH(A4)+1)/12',
            A1 => $end,
            A2 => $end,
            A3 => $assets->comDate,
            A4 => $assets->comDate,
        ],
        A3 => [@combinedFactor],
        A4 => [
            'A5+YEAR(A1)-YEAR(A3)+(MONTH(A2)-MONTH(A4)-1)/12',
            A1 => $assets->comDate,
            A2 => $assets->comDate,
            A3 => $end,
            A4 => $end,
            A5 => $assets->life,
        ],
    );
    my @depreciableAmount = (
        'A1-A2',
        A1 => $assets->cost,
        A2 => $assets->scrapValuation,
    );
    flattenFormula(
        'IF(A1,IF(A8>A9,IF(A21<>A22,(A3-A4)/A5,A61*A62/A63)*A7,0),0)',
        A1  => $assets->life,
        A8  => $end,
        A9  => $start,
        A21 => $assets->annuityIndex,
        A22 => $assets->rateOfReturn,
        A3  => [@numeratorEnd],
        A4  => [@numeratorStart],
        A5  => [@denominator],
        A61 => [ 'A1^A2', @numeratorStart[ 1 .. 4 ] ],
        A62 => [
            'YEAR(A1)-YEAR(A3)+(MONTH(A2)-MONTH(A4)+1)/12',
            A1 => $start,
            A2 => $start,
            A3 => $end,
            A4 => $end,
        ],
        A63 => $assets->life,
        A7  => \@depreciableAmount,
    );
}

sub netValueArithmetic {
    my ( $assets, $periods ) = @_;
    my @indexationFactor = ( '1+A1', A1 => $assets->annuityIndex, );
    my @combinedFactor = (
        '(1+A1)/(1+A2)',
        A1 => $assets->annuityIndex,
        A2 => $assets->rateOfReturn,
    );
    my @denominator = (
        '1-A1^A2',
        A1 => [@combinedFactor],
        A2 => $assets->life,
    );
    my @numerator = (
        'A1^A2*(1-A3^A4)',
        A1 => [@indexationFactor],
        A2 => [
            'YEAR(A1)-YEAR(A3)+(MONTH(A2)-MONTH(A4)+1)/12',
            A1 => $periods->lastDay,
            A2 => $periods->lastDay,
            A3 => $assets->comDate,
            A4 => $assets->comDate,
        ],
        A3 => [@combinedFactor],
        A4 => [
            'A5+YEAR(A1)-YEAR(A3)+(MONTH(A2)-MONTH(A4)-1)/12',
            A1 => $assets->comDate,
            A2 => $assets->comDate,
            A3 => $periods->lastDay,
            A4 => $periods->lastDay,
            A5 => $assets->life,
        ],
    );
    my @depreciableAmount = (
        'A1-A2',
        A1 => $assets->cost,
        A2 => $assets->scrapValuation,
    );
    flattenFormula(
        'IF(A351,'
          . 'A601+IF(A1,IF(A21<>A22,'
          . 'MAX(0,A3/A5),A61*MAX(0,1-A62/A63)'
          . '),1)*A7,0)',
        A351 => $assets->grossValue($periods)->{source},
        A601 => $assets->scrapValuation,
        A1   => $assets->life,
        A21  => $assets->annuityIndex,
        A22  => $assets->rateOfReturn,
        A3   => [@numerator],
        A5   => [@denominator],
        A61  => [ 'A1^A2', @numerator[ 1 .. 4 ] ],
        A62  => [
            'YEAR(A1)-YEAR(A3)+(MONTH(A2)-MONTH(A4)+1)/12',
            A1 => $periods->lastDay,
            A2 => $periods->lastDay,
            A3 => $assets->comDate,
            A4 => $assets->comDate,
        ],
        A63 => $assets->life,
        A7  => \@depreciableAmount,
    );
}

sub disposalGainLossArithmetic {
    my ( $assets, $periods ) = @_;
    my @indexationFactor = ( '1+A1', A1 => $assets->annuityIndex, );
    my @combinedFactor = (
        '(1+A1)/(1+A2)',
        A1 => $assets->annuityIndex,
        A2 => $assets->rateOfReturn,
    );
    my @denominator = (
        '1-A1^A2',
        A1 => [@combinedFactor],
        A2 => $assets->life,
    );
    my @numerator = (
        'A1^A2*(1-A3^A4)',
        A1 => [@indexationFactor],
        A2 => [
            'YEAR(A1)-YEAR(A3)+(MONTH(A2)-MONTH(A4)+1)/12',
            A1 => $assets->decomDate,
            A2 => $assets->decomDate,
            A3 => $assets->comDate,
            A4 => $assets->comDate,
        ],
        A3 => [@combinedFactor],
        A4 => [
            'A5+YEAR(A1)-YEAR(A3)+(MONTH(A2)-MONTH(A4)-1)/12',
            A1 => $assets->comDate,
            A2 => $assets->comDate,
            A3 => $assets->decomDate,
            A4 => $assets->decomDate,
            A5 => $assets->life,
        ],
    );
    my @depreciableAmount = (
        'A1-A2',
        A1 => $assets->cost,
        A2 => $assets->scrapValuation,
    );
    flattenFormula(
        'IF(OR(NOT(A203),A301<A801,A302>A901),0,'
          . 'A651-A601-IF(A1,IF(A21<>A22,'
          . 'MAX(0,A3/A5),A61*MAX(0,1-A62/A63)'
          . '),1)*A7)',
        A203 => $assets->comDate,
        A301 => $assets->decomDate,
        A302 => $assets->decomDate,
        A801 => $periods->firstDay,
        A901 => $periods->lastDay,
        A601 => $assets->scrapValuation,
        A651 => $assets->scrappedValue,
        A1   => $assets->life,
        A21  => $assets->annuityIndex,
        A22  => $assets->rateOfReturn,
        A3   => [@numerator],
        A5   => [@denominator],
        A61  => [ 'A1^A2', @numerator[ 1 .. 4 ] ],
        A62  => [
            'YEAR(A1)-YEAR(A3)+(MONTH(A2)-MONTH(A4)+1)/12',
            A1 => $assets->decomDate,
            A2 => $assets->decomDate,
            A3 => $assets->comDate,
            A4 => $assets->comDate,
        ],
        A63 => $assets->life,
        A7  => \@depreciableAmount,
    );
}

1;
