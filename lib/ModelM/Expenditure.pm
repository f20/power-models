package ModelM;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others.
Copyright 2012-2015 Franck Latrémolière, Reckon LLP and others.

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

sub expenditureAlloc {

    my (
        $model,           $allocLevelset, $allocationRules, $capitalised,
        $directIndicator, $expenditure,   $meavPercentages,
    ) = @_;

    my $networkLengthPercentages;
    ( $networkLengthPercentages, ) =
      $model->networkLengthPercentages($allocLevelset)
      if $model->{allowNetworkLength};

    my $key = join '&', 'preAllocated?allocLevelset=' . ( 0 + $allocLevelset ),
      'allocationRows=' . ( 0 + $allocationRules->{rows} ),
      map { defined $model->{$_} ? "$_=$model->{$_}" : (); } qw(dcp117);

    my $preAllocated = $model->{objects}{$key};

    unless ($preAllocated) {

        ( $preAllocated, ) =
          $model->allocated( $allocLevelset, $allocationRules->{rows} );

        if ( $model->{dcp117} ) {
            if ( $model->{dcp117} =~ /half[ -]?baked/i ) {
                $preAllocated = Stack(
                    name    => 'Allocated costs after DCP 117 adjustments',
                    rows    => $preAllocated->{rows},
                    cols    => $preAllocated->{cols},
                    sources => [
                        Constant(
                            name => 'DCP 117: remove negative number',
                            rows => Labelset(
                                list => [
                                        'Load related new connections & '
                                      . 'reinforcement (net of contributions)'
                                ]
                            ),
                            cols => Labelset( list => ['LV'] ),
                            data => [ [0] ],
                            defaultFormat => '0connz',
                        ),
                        $preAllocated
                    ],
                );
            }
            elsif ( $model->{dcp117} =~ /201[34]/ ) {
                $preAllocated = Stack(
                    name => 'Allocated costs after DCP 117 adjustments',
                    defaultFormat => '0copy',
                    rows          => $preAllocated->{rows},
                    cols          => $preAllocated->{cols},
                    sources       => [
                        Dataset(
                            name =>
                              'Net new connections and reinforcement costs (£)',
                            lines => [
                                'These data are derived from a combination of the'
                                  . ' 2007/2008 regulatory reporting pack (table 2.4)'
                                  . ' and 10-year averages from the FBPQ.',
                                'In a post-DCP 117 legacy Method M workbook,'
                                  . ' these data are on sheet Calc-Opex,'
                                  . ' reading from right to left,'
                                  . ' possibly starting at cell H7.',
                            ],
                            rows => Labelset(
                                list => [ $preAllocated->{rows}{list}[0] ]
                            ),
                            cols          => $preAllocated->{cols},
                            defaultFormat => '0hard',
                            data =>
                              [ map { '' } @{ $preAllocated->{cols}{list} } ],
                            number   => 1329,
                            dataset  => $model->{dataset},
                            appendTo => $model->{objects}{inputTables},
                        ),
                        $preAllocated
                    ],
                );
            }
            else {
                $preAllocated =
                  $model->adjust117( $meavPercentages, $preAllocated );
            }
        }

        $model->{objects}{$key} = $preAllocated;
    }

    $key =
      join '&', 'expenditureAlloc?preAllocated=' . ( 0 + $preAllocated ),
      'allocationRules=' . ( 0 + $allocationRules ),
      'meavPercentages=' . ( 0 + $meavPercentages ),
      'allocLevelset=' .   ( 0 + $allocLevelset ),
      $networkLengthPercentages
      ? ( 'networkLengthPercentages=' . ( 0 + $networkLengthPercentages ) )
      : (),
      map { defined $model->{$_} ? "$_=$model->{$_}" : (); } qw(dcp097);

    return @{ $model->{objects}{$key} } if $model->{objects}{$key};

    my $keyAllocatedTotal =
      'allocatedTotal?preAllocated=' . ( 0 + $preAllocated );
    my $allocatedTotal = $model->{objects}{$keyAllocatedTotal};

    unless ($allocatedTotal) {

        # Avoid SUMIF across sheets: ugly and not compatible with Numbers.app.
        $expenditure = Stack( sources => [$expenditure] );

        $allocatedTotal = GroupBy(
            name          => 'Amounts already allocated',
            source        => $preAllocated,
            rows          => $preAllocated->{rows},
            defaultFormat => '0softnz'
        );

        Columnset(
            columns => [ $expenditure, $allocatedTotal ],
            name    => 'Expenditure data',
        );

        $model->{objects}{$keyAllocatedTotal} = $allocatedTotal;

    }

    my $lvOnly =
      $model->{objects}{ 'lvOnly?cols=' . ( 0 + $preAllocated->{cols} ) } ||=
      Constant(
        name          => 'LV only',
        cols          => $preAllocated->{cols},
        data          => [qw(1 0 0 0)],
        defaultFormat => '0connz'
      );

    my $ehvOnly =
      $model->{objects}{ 'ehvOnly?cols=' . ( 0 + $preAllocated->{cols} ) } ||=
      Constant(
        name          => 'EHV only',
        cols          => $preAllocated->{cols},
        data          => [qw(0 0 0 1)],
        defaultFormat => '0connz'
      );

    my $ar = $model->{dcp097} ? 'IF(A49="Customer numbers",A9,0)' : '0';
    $ar = qq%IF(A48="Network length",A8,$ar)% if $networkLengthPercentages;
    my $allAllocationPercentages = Arithmetic(
        name          => 'All allocation percentages',
        defaultFormat => '%softnz',
        rows          => $preAllocated->{rows},
        cols          => $preAllocated->{cols},
        arithmetic    => '=IF(A44="60%MEAV",0.4*A71+A51,'
          . 'IF(A45="MEAV",A52,'
          . 'IF(A46="EHV only",A6,'
          . 'IF(A47="LV only",A72,'
          . $ar . '))))',
        arguments => {
            A44 => $allocationRules,
            A45 => $allocationRules,
            A46 => $allocationRules,
            A47 => $allocationRules,
            A48 => $allocationRules,
            A51 => $meavPercentages,
            A52 => $meavPercentages,
            A6  => $ehvOnly,
            A71 => $lvOnly,
            A72 => $lvOnly,
            $networkLengthPercentages ? ( A8 => $networkLengthPercentages )
            : (),
            $model->{dcp097}
            ? (
                A49 => $allocationRules,
                A9  => $model->customerNumbersPercentages($allocLevelset)
              )
            : (),
        }
    );

    my $afterAllocation = Arithmetic(
        name          => 'Complete allocation',
        defaultFormat => '0softnz',
        arithmetic    => '=IF(A44="Kill",0,A1+(A2-A3)*A5)',
        arguments     => {
            A1  => $preAllocated,
            A2  => $expenditure,
            A3  => $allocatedTotal,
            A44 => $allocationRules,
            A5  => $allAllocationPercentages,
        }
    );

    my $omittingNegatives = $model->{allowNeg} ? $afterAllocation : Arithmetic(
        name          => 'Complete allocation, zeroing out negative numbers',
        defaultFormat => '0softnz',
        arithmetic    => '=MAX(0,A1)',
        arguments     => { A1 => $afterAllocation }
    );

    my $totalDirect = SumProduct(
        name          => 'Direct costs',
        defaultFormat => '0softnz',
        matrix        => $omittingNegatives,
        vector        => $directIndicator
    );

    my $total = GroupBy(
        name          => 'Total costs',
        defaultFormat => '0softnz',
        source        => $omittingNegatives,
        cols          => $omittingNegatives->{cols},
    );

    my $direct = Arithmetic(
        name          => 'Direct cost proportion for each network level',
        defaultFormat => '%soft',
        arithmetic    => '=A1/A2',
        arguments     => { A1 => $totalDirect, A2 => $total }
    );

    @{ $model->{objects}{$key} } = ( $afterAllocation, $direct );

}

1;
