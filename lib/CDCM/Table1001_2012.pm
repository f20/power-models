﻿package CDCM;

# Copyright 2012-2014 Franck Latrémolière, Reckon LLP and others.
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

sub table1001_2012 {

    my ($model) = @_;

    my @labels = (
        'Base Demand Revenue Before Inflation (A1)',
        'RPI Indexation Factor (A2)',
        'Merger Adjustment (A3)',
        'Base Demand Revenue (A)' . "\r" . '[A = A1*A2 – A3]',
        'Pass-Through Business Rates (B1)',
        'Pass-Through Licence Fees (B2)',
        'Pass-Through Transmission Exit (B3)',
        'Pass-Through Price Control Reopener (B4)',
        'Pass-Through Others (B5)',
        'Allowed Pass-Through Items (B)' . "\r"
          . '[B = B1 + B2 + B3 + B4 + B5]',
        'Losses Incentive #1 (C1)',
        'Losses Incentive #2 (C1)',
        'Losses Incentive #3 (C1)',
        'Losses Incentive #4 (C1)',
        'Quality of Service Incentive Adjustment (C2)',
        'Transmission Connection Point Charges Incentive Adjustment (C3)',
        'Innovation Funding Incentive Adjustment (C4)',
        'Incentive Revenue for Distributed Generation (C5)',
        'Connection Guaranteed Standards Systems & Processes penalty (C6)',
        'Low Carbon Network Fund #1 (C7)',
        'Low Carbon Network Fund #2 (C7)',
        'Low Carbon Network Fund #3 (C7)',
        'Incentive Revenue and Other Adjustments (C)' . "\r"
          . '[C = C1 + C2 + C3 + C4 + C5 + C6 + C7]',
        'Correction Factor (D)',
        'Tax Trigger Mechanism Adjustment (E)',
        'Total Allowed Revenue (F)' . "\r" . '[F = A + B + C + D + E]',
        'Other 1. Excluded services - '
          . 'Top-up, standby, and enhanced system security (G1) (see note 1)',
        'Other 2. Excluded services - '
          . 'Revenue protection services (G2) (see note 1)',
        'Other 3. Excluded services - Miscellaneous (G3) (see note 1)',
        'Other 4. (G4)',
        'Other 5. (G5)',
        'Total Other Revenue to be Recovered by Use of System Charges (G)'
          . "\r"
          . '[G = G1 + G2 + G3 + G4 + G5]',
        'Total Revenue for Use of System Charges (H)' . "\r" . '[H = F + G]',
        '1. Revenue raised outside CDCM - '
          . 'EDCM and Certain Interconnector Revenue (I1)',
        '2. Voluntary under-recovery (I2)',
        '3. Revenue raised outside CDCM (I3)',
        '4. Revenue raised outside CDCM (I4)',
        'Total Revenue to be raised outside the CDCM (I)' . "\r"
          . '[I = I1 + I2 + I3 + I4]',
        'Latest Forecast of CDCM Revenue (J)' . "\r" . '[J = H – I]',
    );
    my @descriptions =
      map {
        my ($d) =
            s/\s*\([A-Z]\)\r\[(.*)\]//s             ? $1
          : s/\s*\(([A-Z][0-9]?)\) \((see .*?)\)$// ? "$1 ($2)"
          : s/\s*\(([A-Z][0-9]?)\)$//               ? $1
          :                                           '';
        $d =~ s/ ([=+–]) /$1/g if length $d > 15;
        $d;
      } @labels;
    my @term = (
        qw(PUt PIADt MGt BRt RBt LFt TBt UNCt),
        'MPTt, HBt, IEDt',
        qw(PTt UILt PCOLt –COLt PPLt IQt ITt IFIt IGt),
        'CGSRAt, CGSSPt, AUMt',
        qw(LCN1t LCN2t LCN3t),
        '',
        qw(–Kt CTRAt ARt ES4 ES5 ES7),
        map { '' } 1 .. 10
    );
    my @crc =
      map { $_ > 0 ? "CRC$_" : ''; }
      qw(3 3 3 3 4 4 4 4 4 3 7 7 7 7 8 9 10 11 12 13 13 13 0 3 3 0 15 15 15),
      -20 .. -10;

    my @fakeLines = split /\n/, <<EOL;
Base Demand Revenue before inflation
Annual Iteration adjustment before inflation
RPI True-up before inflation
Price index adjustment (RPI index)
Base demand revenue
Pass-Through Licence Fees
Pass-Through Business Rates
Pass-Through Transmission Connection Point Charges
Pass-through Smart Meter Communication Licence Costs
Pass-through Smart Meter IT Costs
Pass-through Ring Fence Costs
Pass-Through Others
Allowed Pass-Through Items
Broad Measure of Customer Service incentive
Quality of Service incentive
Connections Engagement incentive
Time to Connect incentive
Losses Discretionary Reward incentive
Network Innovation Allowance
Low Carbon Network Fund - Tier 1 unrecoverable
Low Carbon Network Fund - Tier 2 & Discretionary Funding
Connection Guaranteed Standards Systems & Processes penalty
Residual Losses and Growth Incentive - Losses
Residual Losses and Growth Incentive - Growth
Incentive Revenue and Other Adjustments
Correction Factor
Total allowed Revenue
Other 1. Excluded services - Top-up, standby, and enhanced system security
Other 2. Excluded services - Revenue protection services
Other 3. Excluded services - Miscellaneous
Other 4. Please describe if used
Other 5. Please describe if used
Total other revenue recovered by Use of System Charges
Total Revenue for Use of System Charges
1. Revenue raised outside CDCM - EDCM and Certain Interconnector Revenue
2. Revenue raised outside CDCM - Voluntary under-recovery
3. Revenue raised outside CDCM - Please describe if used
4. Revenue raised outside CDCM - Please describe if used
Total Revenue to be raised outside the CDCM
Latest forecast of CDCM Revenue
EOL

    my $labelset = Labelset( list => \@labels, fakeExtraList => \@fakeLines, );

    my $textnocolourc = [ base => 'textnocolour', align => 'center' ];
    my $textnocolourb = [ base => 'textnocolour', bold  => 1, ];
    my $textnocolourbc =
      [ base => 'textnocolour', bold => 1, align => 'center' ];

    my %avoidDoublePush;

    my $moneyInputFormat =
      $model->{targetRevenue} =~ /million/i
      ? 'millionhard'
      : '0hard';
    my $inputs = Dataset(
        name       => 'Value',
        rows       => $labelset,
        rowFormats => [
            map { /^A2/ ? '0.000hard' : /=/ ? undef : $moneyInputFormat; }
              @descriptions
        ],
        data => [ map { /=/ ? undef : ''; } @descriptions ],
    );

    my $handSubtotal;
    my $subtotals = new SpreadsheetModel::Custom(
        name          => 'Revenue elements and subtotals (£/year)',
        defaultFormat => $model->{targetRevenue} =~ /million/i
        ? 'millionsoft'
        : '0soft',
        rows       => $labelset,
        custom     => [ '=A1', '=0-A1', '=A1*(A2-1)', ],
        arithmetic => '',
        objectType => 'Mixed calculations',
        wsPrepare  => sub {
            my ( $self, $wb, $ws, $format, $formula ) = @_;
            unless ( exists $avoidDoublePush{$wb} ) {
                push @{ $self->{location}{postWriteCalls}{$wb} }, sub {
                    my ( $me, $wbMe, $wsMe, $rowrefMe, $colMe ) = @_;
                    $wsMe->write_string(
                        $$rowrefMe += 1,
                        $colMe - 1,
                        'Note 1: Revenues associated '
                          . 'with excluded services should only be included insofar '
                          . 'as they are charged as Use of System Charges.',
                        $wb->getFormat('text')
                    );
                };
                undef $avoidDoublePush{$wb};
            }
            my $boldFormat = $wb->getFormat(
                [
                    base => $model->{targetRevenue} =~ /million/i
                    ? 'millionsoft'
                    : '0soft',
                    bold => 1
                ]
            );
            my ( $shi, $roi, $coi, $sh, $ro, $co );
            sub {
                my ( $x, $y ) = @_;

                $handSubtotal = [] if !$y;

                ( $shi, $roi, $coi ) = $inputs->wsWrite( $wb, $ws )
                  unless defined $roi;

           # The following trick only works if the object is within a Columnset.
                ( $sh, $ro, $co ) = $self->wsWrite( $wb, $ws )
                  unless defined $ro;

                if ( $descriptions[$y] =~ /=/ ) {
                    my $startRowOffset =
                        $descriptions[$y] =~ /^B/ ? 4
                      : $descriptions[$y] =~ /^C/ ? 10
                      : $descriptions[$y] =~ /^G/ ? 26
                      : $descriptions[$y] =~ /^I/ ? 33
                      :                             0;
                    return '', $boldFormat,
                      '='
                      . (
                        join '+',
                        map { xl_rowcol_to_cell( $ro + $_, $co ) }
                          grep { $_ >= $startRowOffset; } @$handSubtotal
                      );

                    # NB: SUM(A,B,C,...) does not scale well
                }
                else {
                    push @$handSubtotal, $y if $handSubtotal;
                    $descriptions[$y] =~ /^A2/
                      ? (
                        '', $format, $formula->[2],
                        qr/\bA1\b/ => xl_rowcol_to_cell( $roi,     $coi ),
                        qr/\bA2\b/ => xl_rowcol_to_cell( $roi + 1, $coi ),
                      )
                      : $descriptions[$y] =~ /^(A3|I)/ ? (
                        '', $format, $formula->[1],
                        qr/\bA1\b/ => xl_rowcol_to_cell( $roi + $y, $coi ),
                      )
                      : (
                        '', $format, $formula->[0],
                        qr/\bA1\b/ => xl_rowcol_to_cell( $roi + $y, $coi ),
                      );
                }
            };
        }
    );

    my $rowFormatsc = [ map { /=/ ? $textnocolourbc : undef; } @descriptions ];

    Columnset(
        name     => 'CDCM target revenue (monetary amounts in £)',
        number   => 1001,
        appendTo => $model->{inputTables},
        dataset  => $model->{dataset},
        columns  => [
            Constant(
                name          => 'Further description',
                rows          => $labelset,
                defaultFormat => 'textnocolour',
                rowFormats    =>
                  [ map { /=/ ? $textnocolourb : undef; } @descriptions ],
                data => \@descriptions,
            ),
            Dataset(
                name               => 'Term',
                rows               => $labelset,
                defaultFormat      => $textnocolourc,
                rowFormats         => $rowFormatsc,
                data               => \@term,
                usePlaceholderData => 1,
            ),
            Dataset(
                name               => 'CRC',
                rows               => $labelset,
                defaultFormat      => $textnocolourc,
                rowFormats         => $rowFormatsc,
                data               => \@crc,
                usePlaceholderData => 1,
            ),
            $inputs,
            $subtotals,
        ]
    );

    my $specialRowset =
      Labelset( list => [ $labelset->{list}[ $#{ $labelset->{list} } ] ] );

    my $target = new SpreadsheetModel::Custom(
        name          => 'Target CDCM revenue (£/year)',
        defaultFormat => '0soft',
        custom        => [
            join( '+',
                '=A100*A101-A102',
                ( map { "A$_" } 104 .. 108, 110 .. 121, 123, 124, 126 .. 130 ) )
              . '-A133-A134-A135-A136'
        ],
        arithmetic => '= derived from A100',
        rows       => $specialRowset,
        arguments  => { map { ( "A$_" => $inputs ); } 100 .. 136, },
        wsPrepare  => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0], map {
                    my $c = 'A' . ( 100 + $_ );
                    qr/\b$c\b/ =>
                      xl_rowcol_to_cell( $rowh->{A100} + $_, $colh->{A100} );
                } 0 .. 36;
            };
        },
    );

    $model->{edcmTables}[0][4] = new SpreadsheetModel::Custom(
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
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0],
                  qr/\bA1\b/ => xl_rowcol_to_cell( $rowh->{A1}, $colh->{A1} ),
                  qr/\bA2\b/ => xl_rowcol_to_cell(
                    $rowh->{A2} + 33,    # hard-coded reference to EDCM revenue
                    $colh->{A2}
                  );
            };
        },
    ) if $model->{edcmTables};

    Columnset(
        name    => 'Target CDCM revenue',
        columns => [
            $target,
            Arithmetic(
                name          => 'Check (should be zero)',
                defaultFormat => '0soft',
                arguments     => { A1 => $target, A2 => $subtotals, },
                arithmetic    => '=A1-A2'
            )
        ]
    );

    $target;

}

1;
