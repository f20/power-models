package Financial::Pricing;

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
use base 'Financial::Inputs';
use SpreadsheetModel::Shortcuts ':all';
use SpreadsheetModel::CalcBlock;

sub finish {

    my ($pricing) = @_;

    Columnset(
        name => 'Purchase or construction cost'
          . ' (at time of bidding, winner\'s curse adjusted)',
        number        => 1405,
        singleRowName => 'Cost',
        appendTo      => $pricing->{model}{inputTables},
        dataset       => $pricing->{model}{dataset},
        columns       => [
            Dataset(
                name => 'Asset name' . ( $pricing->{suffix} || '' ),
                defaultFormat => 'texthard',
                data          => [''],
            ),
            $pricing->{constructionCost}[0],
        ],
    ) if $pricing->{constructionCost}[0];

    Columnset(
        name          => 'Purchase or construction cost (latest estimate)',
        number        => 1415,
        singleRowName => 'Cost',
        appendTo      => $pricing->{model}{inputTables},
        dataset       => $pricing->{model}{dataset},
        columns       => [
            Dataset(
                name => 'Asset name' . ( $pricing->{suffix} || '' ),
                defaultFormat => 'texthard',
                data          => [''],
            ),
            Dataset( name => 'Not used', data => [] ),
            $pricing->{constructionCost}[1],
        ],
    ) if $pricing->{constructionCost}[1];

    Columnset(
        name => 'Market cost of capital '
          . '(at time of bidding, winner\'s curse adjusted)',
        singleRowName => 'WACC',
        number        => 1407,
        appendTo      => $pricing->{model}{inputTables},
        dataset       => $pricing->{model}{dataset},
        columns       => [
            Constant(
                name          => ' ',
                defaultFormat => 'th',
                data          => ['Annual cost of capital'],
            ),
            $pricing->{costOfCapital}[0],
        ],
    ) if $pricing->{costOfCapital}[0];

    Columnset(
        name          => 'Market cost of capital (latest estimate)',
        singleRowName => ' ',
        number        => 1417,
        appendTo      => $pricing->{model}{inputTables},
        dataset       => $pricing->{model}{dataset},
        columns       => [
            Constant(
                name          => '',
                defaultFormat => 'th',
                data          => ['Annual cost of capital'],
            ),
            Dataset( name => 'Not used', data => [] ),
            $pricing->{costOfCapital}[1],
        ],
    ) if $pricing->{costOfCapital}[1];

    push @{ $pricing->{model}{preprocessingSheets} }, [
        Pricing => sub {
            my ( $wbook, $wsheet ) = @_;
            $wsheet->freeze_panes( 1, 1 );
            $wsheet->set_column( 0, 0,   42 );
            $wsheet->set_column( 1, 250, 15 );
            $_->wsWrite( $wbook, $wsheet )
              foreach Notes( name => 'Pricing' ),
              $pricing->{tables} ? @{ $pricing->{tables} } : ();
        }
    ];

}

sub periodset {
    my ( $pricing, $version ) = @_;
    $version ||= 0;
    $pricing->{periodset}[$version] ||= Labelset(
        list => [
            map { "Period $_"; }
              ( 1 + $version ) .. ( $pricing->{numPeriods} ||= 5 )
        ]
    );
}

sub constructionCost {
    my ( $pricing, $version ) = @_;
    $pricing->{constructionCost}[ $version || 0 ] ||= Dataset(
        name          => 'Future new asset purchase or construction cost (£)',
        defaultFormat => '0hard',
        cols          => $pricing->periodset($version),
        data => [ map { '' } @{ $pricing->periodset($version)->{list} } ],
    );
}

sub costOfCapital {
    my ( $pricing, $version ) = @_;
    $pricing->{costOfCapital}[ $version || 0 ] ||= Dataset(
        name          => 'Annual cost of capital',
        defaultFormat => '%hard',
        cols          => $pricing->periodset($version),
        data => [ map { 0.1 } @{ $pricing->periodset($version)->{list} } ],
    );
}

sub leaseRate {

    my ( $pricing, $version ) = @_;
    $pricing->{periodLengthInYears} ||= 8;
    my $periodset       = $pricing->periodset($version);
    my @relevantPeriods = reverse @{ $periodset->{list} };
    my $costOfCapital =
      Stack( sources => [ $pricing->costOfCapital($version) ] );
    my $constructionCost =
      Stack( sources => [ $pricing->constructionCost($version) ] );

    my $discountFactor = Arithmetic(
        name       => 'Discount factor over the period',
        arithmetic => '=(1+A1)^-' . $pricing->{periodLengthInYears},
        arguments  => { A1 => $costOfCapital, },
    );

    my $periodWeight = Arithmetic(
        name       => 'Present value weight of period income',
        arithmetic => '=(1-A1)/A2',
        arguments  => { A1 => $discountFactor, A2 => $costOfCapital, },
    );

    my $combinedFactor = SpreadsheetModel::Custom->new(
        name       => 'Combined discount and indexation factor',
        custom     => ['=A1*A2/A3'],
        arithmetic => '=A1*A2/(previous A3)',
        cols       => $periodset,
        arguments  => {
            A1 => $discountFactor,
            A2 => $constructionCost,
            A3 => $constructionCost,
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                return '', $wb->getFormat('unavailable')
                  unless $x == $#{ $periodset->{list} };
                '', $format, $formula->[0],
                  qr/\bA1\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A1}, $colh->{A1} + $x,
                  ),
                  qr/\bA2\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A2}, $colh->{A2} + $x
                  ),
                  qr/\bA3\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A3}, $colh->{A3} + $x - 1,
                  ),
                  ;
            };
        },
    );

    my @compoundDiscountFactor = Constant(
        name => "$relevantPeriods[0] compound discount factor",
        cols => $periodset,
        data => [
            map { $_ == $#{ $periodset->{list} } ? 1 : undef; }
              0 .. $#{ $periodset->{list} }
        ]
    );

    foreach my $period ( 1 .. $#relevantPeriods ) {
        push @compoundDiscountFactor, SpreadsheetModel::Custom->new(
            name       => "$relevantPeriods[$period] compound discount factor",
            cols       => $periodset,
            custom     => [ '=A1*A2', ],
            arithmetic => '= 0 or 1 or A1*A2',
            arguments  => {
                A1 => $compoundDiscountFactor[ $period - 1 ],
                A2 => $discountFactor,
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    return 0, $wb->getFormat('unavailable')
                      if $x + $period < $#{ $periodset->{list} };
                    return 1, $wb->getFormat('0.000con')
                      if $x + $period == $#{ $periodset->{list} };
                    '', $format, $formula->[0],
                      qr/\bA1\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{A1}, $colh->{A1} + $x,
                      ),
                      qr/\bA2\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{A2},
                        $colh->{A2} + $#{ $periodset->{list} } - $period );
                };
            },
        );
    }

    my @periodWeight;
    foreach my $period ( 0 .. $#relevantPeriods ) {
        my $foldedAtEnd = $pricing->{assetLifeInPeriods} - $period;
        push @periodWeight, SpreadsheetModel::Custom->new(
            name => "$relevantPeriods[$period] present value factor for income",
            cols => $periodset,
            custom => [
                '=A1',
                $foldedAtEnd > 1 ? "=A11*(1-A21^$foldedAtEnd)/(1-A22)" : ()
            ],
            arithmetic => '= 0 or A1'
              . (
                $foldedAtEnd > 1 ? " or A11*(1-A21^$foldedAtEnd)/(1-A22)"
                : ''
              ),
            arguments => {
                A1 => $periodWeight,
                $foldedAtEnd > 1
                ? (
                    A11 => $periodWeight,
                    A21 => $combinedFactor,
                    A22 => $combinedFactor,
                  )
                : (),
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    return 0, $wb->getFormat('unavailable')
                      if $x + $period > $#{ $periodset->{list} } +
                      $pricing->{assetLifeInPeriods} - 1;
                    return 0, $wb->getFormat('unavailable')
                      if $x + $period < $#{ $periodset->{list} };
                    '',
                      $foldedAtEnd > 1
                      && $x == $#{ $periodset->{list} } ? $format
                      : $wb->getFormat('0.000copy'),
                      $formula->[ $foldedAtEnd > 1
                      && $x == $#{ $periodset->{list} } ? 1 : 0 ],
                      qr/\bA1\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{A1},
                        $colh->{A1} + $#{ $periodset->{list} } - $period,
                      ),
                      $foldedAtEnd > 1
                      ? (
                        qr/\bA11\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A11},
                            $colh->{A11} + $#{ $periodset->{list} } - $period,
                          ),
                        qr/\bA21\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A21}, $colh->{A21} + $x
                          ),
                        qr/\bA22\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A22}, $colh->{A22} + $x
                          ),
                      )
                      : ();
                };
            },
        );
    }

    my @leaseRates;
    foreach my $period ( 0 .. $#relevantPeriods ) {
        push @leaseRates, SpreadsheetModel::Custom->new(
            name          => "$relevantPeriods[$period] lease rate estimation ",
            defaultFormat => '0soft',
            cols          => $periodset,
            custom        => [
                $period ? ( '=(A1-SUMPRODUCT(A2:A3,A4:A5,A6:A7))/A8', '=A9' )
                : '=A1/A8',
            ],
            arithmetic => $period ? '= A9 or (A1-SUMPRODUCT(A2,A4,A6))/A8'
            : '=A1/A8',
            arguments => {
                A1 => $constructionCost,
                A8 => $periodWeight[$period],
                $period
                ? (
                    A2 => $leaseRates[ $period - 1 ],
                    A3 => $leaseRates[ $period - 1 ],
                    A4 => $periodWeight[$period],
                    A5 => $periodWeight[$period],
                    A6 => $compoundDiscountFactor[$period],
                    A7 => $compoundDiscountFactor[$period],
                    A9 => $leaseRates[ $period - 1 ],
                  )
                : (),
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    return 0, $wb->getFormat('unavailable')
                      if $x + $period < $#{ $periodset->{list} };
                    '',
                      $x + $period == $#{ $periodset->{list} } ? $format
                      : $wb->getFormat('0copy'),
                      $formula->[
                      $x + $period == $#{ $periodset->{list} } ? 0
                      : 1
                      ],
                      qr/\bA1\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{A1},
                        $colh->{A1} + $#{ $periodset->{list} } - $period,
                      ),
                      qr/\bA8\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{A8}, $colh->{A8} + $x,
                      ),
                      $period
                      ? (
                        qr/\bA2\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A2}, $colh->{A2}, 0, 1,
                          ),
                        qr/\bA3\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A3}, $colh->{A3} + $#{ $periodset->{list} },
                            0, 1,
                          ),
                        qr/\bA4\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A4}, $colh->{A4}, 0, 1,
                          ),
                        qr/\bA5\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A5}, $colh->{A5} + $#{ $periodset->{list} },
                            0, 1,
                          ),
                        qr/\bA6\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A6}, $colh->{A6}, 0, 1,
                          ),
                        qr/\bA7\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A7}, $colh->{A7} + $#{ $periodset->{list} },
                            0, 1,
                          ),
                        qr/\bA9\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A9}, $colh->{A9} + $x,
                          ),
                      )
                      : ();
                };
            },
        );
    }

    push @{ $pricing->{tables} },
      CalcBlock(
        name => 'Annual lease rate estimation'
          . ( $version ? ' (revised pricing run)' : ' (initial pricing run)' ),
        items => [
            $costOfCapital,          $discountFactor,
            @compoundDiscountFactor, $constructionCost,
            $combinedFactor,         $periodWeight,
            @periodWeight,           @leaseRates,
        ],
      );

    @leaseRates;

}

sub leaseRateCombined {
    my ( $pricing, $flow ) = @_;
    return $pricing->{leaseRateCombined}{ 0 + $flow }
      if $pricing->{leaseRateCombined}{ 0 + $flow };
    my @leaseRate1 = $pricing->leaseRate;
    my @leaseRate2 = $pricing->leaseRate(1);
    push @{ $pricing->{tables} },
      my $leaseRateCombined = SpreadsheetModel::Custom->new(
        name          => 'Annual ' . $flow->{show_flow},
        defaultFormat => $flow->{show_formatBase} . 'copy',
        rows          => $flow->labelsetNoNames,
        custom        => [
            '=A1',
            ( map { "=A2$_"; } 0 .. ( $pricing->{assetLifeInPeriods} - 2 ) ),
            '=A5',
        ],
        arguments => {
            A1 => $leaseRate1[$#leaseRate1],
            (
                map { ( "A2$_" => $leaseRate2[ $#leaseRate2 - $_ ] ); }
                  0 .. ( $pricing->{assetLifeInPeriods} - 2 )
            ),
            A5 => $flow->{inputDataColumns}{annual},
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                '', $format,
                  $formula->[
                    $y > $pricing->{assetLifeInPeriods}
                  ? $pricing->{assetLifeInPeriods}
                  : $y
                  ],
                  qr/\bA1\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A1}, $colh->{A1} + $y,
                    1, 1, ),
                  qr/\bA5\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A5} + $y,
                    $colh->{A5},
                  ),
                  map {
                    (
                        qr/\bA2$_\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{"A2$_"}, $colh->{"A2$_"} + $y - 1,
                            1, 1,
                          )
                    );
                  } 0 .. ( $pricing->{assetLifeInPeriods} - 2 );
            };
        },
      );
    $pricing->{leaseRateCombined}{ 0 + $flow } = $leaseRateCombined;
}

sub sales {
    my ($pricing) = @_;
    (
        $pricing->SUPER::sales,
        name          => 'Sales',
        annualClosure => sub {
            my ($flow) = @_;
            $pricing->{assetLifeInPeriods} ||= 4;
            $flow->{inputDataColumns}{annual} = Constant(
                name          => 'Annual ' . $flow->{show_flow},
                defaultFormat => '0hard',
                rowFormats =>
                  [ map { 'unused'; } 1 .. $pricing->{assetLifeInPeriods}, ],
                rows => $flow->labelsetNoNames,
                data => [
                    (
                        map { 'calculated'; }
                          1 .. $pricing->{assetLifeInPeriods}
                    ),
                    (
                        map { ''; }
                          $pricing->{assetLifeInPeriods}
                          .. $#{ $flow->labelsetNoNames->{list} }
                    )
                ],
            );
            $pricing->leaseRateCombined($flow);
        },
    );
}

sub assets {
    my ($pricing) = @_;
    (
        $pricing->SUPER::assets,
        costClosure => sub {
            my ($assets) = @_;
            $assets->{inputDataColumns}{cost} = Dataset(
                name          => 'Cost (£)',
                defaultFormat => '0hard',
                rowFormats    => [ 'unused', ],
                rows          => $assets->labelsetNoNames,
                data          => [
                    'copied',
                    map { 0; } 2 .. @{ $assets->labelsetNoNames->{list} }
                ],
            );
            SpreadsheetModel::Custom->new(
                name          => 'Cost (£)',
                defaultFormat => '0copy',
                rows          => $assets->labelsetNoNames,
                custom        => [ '=A1', '=A2', ],
                arguments     => {
                    A1 => $pricing->constructionCost,
                    A2 => $assets->{inputDataColumns}{cost},
                },
                wsPrepare => sub {
                    my ( $self, $wb, $ws, $format, $formula, $pha, $rowh,
                        $colh ) = @_;
                    sub {
                        my ( $x, $y ) = @_;
                        return '', $format, $formula->[1],
                          qr/\bA2\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A2} + $y,
                            $colh->{A2} )
                          if $y;
                        '', $format, $formula->[0],
                          qr/\bA1\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A1}, $colh->{A1} );
                    };
                },
            );
        },
    );
}

sub capitalExp {
    my ($pricing) = @_;
    (
        $pricing->SUPER::capitalExp,
        amountClosure => sub {
            my ($flow) = @_;
            $flow->{inputDataColumns}{amount} = Dataset(
                name          => 'Total ' . $flow->{show_flow},
                defaultFormat => $flow->{show_formatBase} . 'hard',
                rows          => $flow->labelsetNoNames,
                rowFormats    => [ 'unused', ],
                data          => [
                    'copied',
                    map { 0; } 2 .. @{ $flow->labelsetNoNames->{list} }
                ],
            );
            SpreadsheetModel::Custom->new(
                name          => 'Total ' . $flow->{show_flow},
                defaultFormat => $flow->{show_formatBase} . 'copy',
                rows          => $flow->labelsetNoNames,
                custom        => [ '=A1', '=A2', ],
                arguments     => {
                    A1 => $pricing->constructionCost,
                    A2 => $flow->{inputDataColumns}{amount},
                },
                wsPrepare => sub {
                    my ( $self, $wb, $ws, $format, $formula, $pha, $rowh,
                        $colh ) = @_;
                    sub {
                        my ( $x, $y ) = @_;
                        return '', $format, $formula->[1],
                          qr/\bA2\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A2} + $y,
                            $colh->{A2} )
                          if $y;
                        '', $format, $formula->[0],
                          qr/\bA1\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A1}, $colh->{A1} );
                    };
                },
            );
        },
    );
}

1;
