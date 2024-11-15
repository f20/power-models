package CDCM;

# Copyright 2012-2024 Franck Latrémolière, Reckon LLP and others.
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
use Spreadsheet::WriteExcel::Utility;

sub table1001_2024 {

    my ($model) = @_;
    return @{ $model->{table1001_array} } if $model->{table1001_array};

    my @inputlines = split /\n/, <<EOL;
Fast money
Depreciation
Return
Licence Fee Payments
Prescribed Rates
Pass-through Transmission Connection Point Charges
Smart Meter Communication Licensee Costs
Smart Meter Information Technology Costs
Ring Fence Costs
Supplier of Last Resort Net Costs
Valid Bad Debt Claims
Pension Scheme Established Deficit repair expenditure
Failed Supplier Recovered Costs
Shetland Variable Energy Costs (SSEH only)
Assistance for high-cost distributors adjustment (SSEH only)
Return Adjustment
Equity issuance costs
Business plan incentive
Output delivery incentive
Other revenue allowances
Directly Remunerated Services
Tax allowance
Tax allowance adjustment
Real to nominal prices conversion factor (splice index for RIIO-2)
Correction term
Forecasting penalty
Legacy Allowed Revenue
Revenue raised outside CDCM - EDCM and Certain Interconnector Revenue
Latest forecast of CDCM Revenue
EOL

    my $labelset = Labelset( list => \@inputlines );

    my $inputs = Dataset(
        name       => 'Value',
        rows       => $labelset,
        rowFormats => [
            map {
                    /raised outside/    ? '0hard'
                  : /conversion factor/ ? '0.000hard'
                  :                       '0.0hard';
            } @inputlines
        ],
        data => [
            map { /Latest forecast of CDCM Revenue/ ? undef : ''; } @inputlines
        ],
    );

    my $calculations = new SpreadsheetModel::Custom(
        name          => 'Calculations (£/year)',
        defaultFormat => '0soft',
        rows          => $labelset,
        custom        => [],
        arithmetic    => '',
        objectType    => 'Mixed calculations',
        wsPrepare     => sub {
            my ( $self, $wb, $ws, $format, $formula ) = @_;
            my ( $inputCellMaker, $calcCellMaker );
            my $unavailable = $wb->getFormat('unavailable');
            sub {
                my ( $x, $y ) = @_;

                # Danger - hardcoding
                return '', $unavailable if $y == 23;

                unless ($inputCellMaker) {
                    my ( $sh, $ro, $co ) = $inputs->wsWrite( $wb, $ws );
                    $inputCellMaker =
                      sub { xl_rowcol_to_cell( $ro + $_[0], $co, $_[1] ); };
                }
                unless ($calcCellMaker) {
                    my ( $sh, $ro, $co ) = $self->wsWrite( $wb, $ws );
                    $calcCellMaker =
                      sub { xl_rowcol_to_cell( $ro + $_[0], $co, $_[1] ); }
                }

                # Danger - hardcoding and insanity
                return
                    '=-1e6*'
                  . $inputCellMaker->($y) . '*'
                  . $inputCellMaker->( 23, 1 ), $format
                  if $y == 12 || $y == 14;
                return
                    '=1e6*'
                  . $inputCellMaker->($y) . '*'
                  . $inputCellMaker->( 23, 1 ), $format
                  if $y < 23;
                return
                    '=SUM('
                  . $calcCellMaker->( 0,  1 ) . ':'
                  . $calcCellMaker->( 27, 1 )
                  . ')', $format
                  if $y == 28;
                return '=-1*' . $inputCellMaker->($y),  $format if $y == 27;
                return '=1e6*' . $inputCellMaker->($y), $format;
            };
        },
    );

    Columnset(
        name     => 'CDCM target revenue',
        number   => 1001,
        appendTo => $model->{inputTables},
        dataset  => $model->{dataset},
        columns  => [
            Constant(
                name => 'Padding 1',
                rows => $labelset,
                data => [],
            ),
            Constant(
                name => 'Padding 2',
                rows => $labelset,
                data => [],
            ),
            Constant(
                name => 'Padding 3',
                rows => $labelset,
                data => [],
            ),
            $inputs,
            $calculations,
        ],
    );

    my $specialRowset =
      Labelset( list => [ $labelset->{list}[ $#{ $labelset->{list} } ] ] );

    my $target = new SpreadsheetModel::Custom(    # Danger - hardcoding
        name          => 'Target revenue (£/year)',
        custom        => ['=A1'],
        arithmetic    => '=A1',
        defaultFormat => '0copy',
        arguments     => {
            A1 => $calculations,
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0],
                  qr/\bA1\b/ =>
                  xl_rowcol_to_cell( $rowh->{A1} + 28, $colh->{A1} );
            };
        },
    );

    if ( $model->{edcmTables} ) {
        $model->{edcmTables}[0][4] =
          new SpreadsheetModel::Custom(    # Danger - hardcoding
            name => 'The amount of money that the DNO wants to raise from use'
              . ' of system charges (£/year)',
            custom        => ['=A1+A2'],
            arithmetic    => '=A1+A2',
            defaultFormat => '0soft',
            arguments     => {
                A1 => $target,
                A2 => $inputs,
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    '', $format, $formula->[0],
                      qr/\bA1\b/ =>
                      xl_rowcol_to_cell( $rowh->{A1}, $colh->{A1} ),
                      qr/\bA2\b/ => xl_rowcol_to_cell(
                        $rowh->{A2} + 27, # hard-coded reference to EDCM revenue
                        $colh->{A2}
                      );
                };
            },
          );
        $model->{edcmTables}[0][6] =
          new SpreadsheetModel::Custom(    # Danger - hardcoding
            name =>
              'Target revenue for domestic demand fixed charge adder (£/year)',
            custom        => ['=A1'],
            arithmetic    => '=A1',
            defaultFormat => '0copy',
            arguments     => {
                A1 => $calculations,
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    '', $format, $formula->[0],
                      qr/\bA1\b/ =>
                      xl_rowcol_to_cell( $rowh->{A1} + 9, $colh->{A1} );
                };
            },
          );
        $model->{edcmTables}[0][7] =
          new SpreadsheetModel::Custom(    # Danger - hardcoding
            name =>
              'Target revenue for metered demand fixed charge adder (£/year)',
            custom        => ['=A1'],
            arithmetic    => '=A1',
            defaultFormat => '0copy',
            arguments     => {
                A1 => $calculations,
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    '', $format, $formula->[0],
                      qr/\bA1\b/ =>
                      xl_rowcol_to_cell( $rowh->{A1} + 10, $colh->{A1} );
                };
            },
          );
    }

    $model->{table1001_array} = [
        $target,
        new SpreadsheetModel::Custom(    # Danger - hardcoding
            name =>
              'Target revenue for domestic demand fixed charge adder (£/year)',
            custom        => ['=A1'],
            arithmetic    => '=A1',
            defaultFormat => '0copy',
            arguments     => {
                A1 => $calculations,
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    '', $format, $formula->[0],
                      qr/\bA1\b/ =>
                      xl_rowcol_to_cell( $rowh->{A1} + 9, $colh->{A1} );
                };
            },
        ),
        new SpreadsheetModel::Custom(    # Danger - hardcoding
            name =>
              'Target revenue for metered demand fixed charge adder (£/year)',
            custom        => ['=A1'],
            arithmetic    => '=A1',
            defaultFormat => '0copy',
            arguments     => {
                A1 => $calculations,
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    '', $format, $formula->[0],
                      qr/\bA1\b/ =>
                      xl_rowcol_to_cell( $rowh->{A1} + 10, $colh->{A1} );
                };
            },
        ),
    ];

    @{ $model->{table1001_array} };

}

1;
