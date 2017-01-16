package Financial::EquityReturns;

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
use SpreadsheetModel::CalcBlock;

sub new {
    my ( $class, %hash ) = @_;
    $hash{$_} || die __PACKAGE__ . " needs a $_ attribute"
      foreach qw(model income balance cashflow);
    bless \%hash, $class;
}

sub cashToFromInvestors {
    my ( $self, $periods ) = @_;
    $self->{cashToFromInvestors}{ 0 + $periods } ||= CalcBlock(
        name => $periods->decorate(
            'Cash flow to/from equity investors' . ( $self->{suffix} || '' )
        ),
        items => [
            A1 => $self->{cashflow}->investors($periods),
            A2 => $self->{balance}->equityInitialAndRaised($periods),
            A3 => {
                name          => 'Cash flow to/from equity investors',
                defaultFormat => '0soft',
                arithmetic    => '=A1-A2',
            },
            A4 => {
                name          => $periods->decorate('Net equity raised (£)'),
                defaultFormat => '0soft',
                arithmetic    => '=0-MIN(A3,0)',
            },
            A5 => {
                name          => $periods->decorate('Net distributions (£)'),
                defaultFormat => '0soft',
                arithmetic    => '=MAX(A3,0)',
            },
        ]
    );
}

sub npv {
    my ( $self, $periods ) = @_;
    return $self->{npv}{ 0 + $periods } if $self->{npv}{ 0 + $periods };
    my $npvset = Labelset(
        name          => 'Items without names',
        defaultFormat => 'thitem',
        list          => [ 1 .. $self->{model}{npvLines} || 3 ]
    );
    my $names = Dataset(
        name          => 'NPV option name',
        defaultFormat => 'texthard',
        rows          => $npvset,
        data          => [ map { ''; } @{ $npvset->{list} } ]
    );
    my $discountRate = Dataset(
        name          => 'NPV discount rate',
        defaultFormat => '%hard',
        rows          => $npvset,
        data          => [ map { 0.05 * $_; } 0 .. $#{ $npvset->{list} } ]
    );
    Columnset(
        name     => 'Net present value parameters',
        number   => 1480,
        columns  => [ $names, $discountRate, ],
        dataset  => $self->{model}{dataset},
        appendTo => $self->{model}{inputTables},
    );
    my $flow = $self->cashToFromInvestors($periods)->{A3};
    $self->{npv}{ 0 + $periods } = SpreadsheetModel::Custom->new(
        name          => 'Equity net present value',
        defaultFormat => '0soft',
        rows          => $npvset,
        legendText    => $names,
        cols          => $periods->labelset,
        custom        => ['=NPV(A1,A2:A3)'],
        arithmetic    => '=NPV(A1,A2 (past only))',
        arguments     => { A1 => $discountRate, A2 => $flow, A3 => $flow, },
        wsPrepare     => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0],
                  qr/\bA1\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A1} + $y,
                    $colh->{A1}, 0, 1, ),
                  qr/\bA2\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A2}, $colh->{A2}, 1, 1, ),
                  qr/\bA3\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A2}, $colh->{A2} + $x, 1, );
            };
        },
    );
}

sub chart_npv {
    my ( $self, $periods ) = @_;
    SpreadsheetModel::Chart->new(
        name => 'NPV',
        type => 'line',
        $periods->{priorPeriod} ? ( ignore_left => 1 ) : (),
        instructions => [
            add_series => $self->npv($periods),
            set_x_axis => [
                num_font => { size => 16 },
                interval_unit => 1 + int( @{ $periods->labelset->{list} } / 6 ),
            ],
            set_y_axis => [ num_font => { size => 16 }, ],
            set_legend => [ font     => { size => 16 }, ],
            set_title  => [
                name => 'Equity NPV',
            ],
        ],
    );
}

sub equityInternalRateOfReturn {
    my ( $self, $periods ) = @_;
    return $self->{equityInternalRateOfReturn}{ 0 + $periods }
      if $self->{equityInternalRateOfReturn}{ 0 + $periods };
    my $block  = $self->cashToFromInvestors($periods);
    my $number = Arithmetic(
        name          => 'Equity IRR',
        defaultFormat => '%soft',
        arithmetic    => '=IRR(A1_A2)',
        arguments     => {
            A1_A2 => $block->{A3},
        }
    );
    my $text = Arithmetic(
        name          => 'Graph title',
        defaultFormat => 'textsoft',
        arithmetic    => '="Equity IRR = "&TEXT(A1,"0.0%")',
        arguments     => {
            A1 => $number,
        },
    );
    my @cols = ( $number, $text );
    Columnset(
        name    => 'Internal rate of return on equity',
        columns => \@cols
    );
    $self->{equityInternalRateOfReturn}{ 0 + $periods } =
      [ @cols, @{$block}{qw(A4 A5)} ];
}

sub chart_equity_dividends {
    my ( $self, $periods ) = @_;
    SpreadsheetModel::Chart->new(
        name         => 'IRR',
        type         => 'column',
        instructions => [
            add_series => [
                1
                ? $self->equityInternalRateOfReturn($periods)->[2]
                : $self->{balance}->equityInitialAndRaised($periods),
                overlap => 100,
                pattern => {
                    pattern  => 'percent_10',    # 'horizontal_brick',
                    fg_color => 'yellow',
                    bg_color => 'red',
                },
            ],
            add_series => [
                1
                ? $self->equityInternalRateOfReturn($periods)->[3]
                : $self->{cashflow}->investors($periods),
                gap  => 0,
                fill => { color => 'black' },
            ],
            set_legend => [ position => 'top', font => { size => 16 }, ],
            set_x_axis => [
                num_font  => { size => 16 },
                name_font => { size => 16 },
                interval_unit => 1 + int( @{ $periods->labelset->{list} } / 6 ),
            ],
            set_y_axis =>
              [ num_font => { size => 16 }, name_font => { size => 16 }, ],
            set_title => [
                name_formula => $self->equityInternalRateOfReturn($periods)->[1]
            ],
        ],
    );
}

1;
