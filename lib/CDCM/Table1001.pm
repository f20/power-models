package CDCM;

=head Copyright licence and disclaimer

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
use Spreadsheet::WriteExcel::Utility;

sub table1001 {

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
      $model->{targetRevenue} =~ /DCP132longlabels/i ? ( map { '' } @labels )
      : (
        map {
            my ($d) =
                s/\s*\([A-Z]\)\r\[(.*)\]//s             ? $1
              : s/\s*\(([A-Z][0-9]?)\) \((see .*?)\)$// ? "$1 ($2)"
              : s/\s*\(([A-Z][0-9]?)\)$//               ? $1
              :                                           '';
            $d =~ s/ ([=+–]) /$1/g if length $d > 15;
            $d;
        } @labels
      );
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

    my $labelset = Labelset( list => \@labels );
    my $textnocolourc = [ base => 'textnocolour', align => 'center' ];
    my $textnocolourb = [ base => 'textnocolour', bold  => 1, ];
    my $textnocolourbc =
      [ base => 'textnocolour', bold => 1, align => 'center' ];

    if ( $model->{targetRevenue} =~ /DCP132longlabels/i ) {

        my %avoidDoublePush;

        my $numbers = new SpreadsheetModel::Custom(
            name          => 'Value',
            defaultFormat => $model->{targetRevenue} =~ /million/i
            ? 'millionhard'
            : '0hard',
            rows   => $labelset,
            custom => [
                '=IV1*IV2-IV3',  '=IV1-IV2',
                '=SUM(IV1:IV2)', '=IV1+IV2+IV3+IV4+IV5',
                '=IV1+IV2',
            ],
            arithmetic => '',
            objectType => 'Mixed inputs and calculations',
            wsPrepare  => sub {
                my ( $self, $wb, $ws, $format, $formula ) = @_;
                unless ( exists $avoidDoublePush{$wb} ) {
                    push @{ $self->{location}{postWriteCalls}{$wb} }, sub {
                        $ws->write_string(
                            $ws->{nextFree}++,
                            0,
                            'Note 1: Cost categories associated '
                              . 'with excluded services should only be populated '
                              . 'if the Company recovers the costs of providing '
                              . 'these services from Use of System Charges.',
                            $wb->getFormat('text')
                        );
                    };
                    undef $avoidDoublePush{$wb};
                }
                my ( $sh, $ro, $co );
                my $softFormat = $wb->getFormat(
                    [
                        base => $model->{targetRevenue} =~ /million/i
                        ? 'millionsoft'
                        : '0soft',
                        bold => 1
                    ]
                );
                my $data;
                if ( $model->{dataset} ) {
                    if ( my $d1001 = $model->{dataset}{1001} ) {
                        $data = $d1001->[4];
                    }
                }
                my $dataEntryMaker = sub {
                    my ($format) = @_;
                    sub {
                        my ( $x, $y ) = @_;
                        my $v;
                        if ($data) {
                            local $_ = $labelset->{list}[$y];
                            s/.*\n//s;
                            s/[^A-Za-z0-9 -]/ /g;
                            s/- / /g;
                            s/ +/ /g;
                            s/^ //;
                            s/ $//;
                            s/ see note 1//i;
                            s/( [A-Z][0-9]?)+$//;
                            my $k = $_;
                            ($v) =
                              map {
                                substr( $_, 0, length $k ) eq $k
                                  ? $data->{$_}
                                  : ();
                              } keys %$data;
                        }
                        defined $v ? $v : '#VALUE!', $format;
                    };
                };
                my $dataEntry = $dataEntryMaker->($format);
                my $calcA     = sub {
                    '', $softFormat, $formula->[0],
                      IV1 => xl_rowcol_to_cell( $ro,     $co ),
                      IV2 => xl_rowcol_to_cell( $ro + 1, $co ),
                      IV3 => xl_rowcol_to_cell( $ro + 2, $co );
                };
                my $calcB = sub {
                    '', $softFormat, $formula->[2],
                      IV1 => xl_rowcol_to_cell( $ro + 4, $co ),
                      IV2 => xl_rowcol_to_cell( $ro + 8, $co );
                };
                my $calcC = sub {
                    '', $softFormat, $formula->[2],
                      IV1 => xl_rowcol_to_cell( $ro + 10, $co ),
                      IV2 => xl_rowcol_to_cell( $ro + 21, $co );
                };
                my $calcF = sub {
                    '', $softFormat, $formula->[3],
                      IV1 => xl_rowcol_to_cell( $ro + 3,  $co ),
                      IV2 => xl_rowcol_to_cell( $ro + 9,  $co ),
                      IV3 => xl_rowcol_to_cell( $ro + 22, $co ),
                      IV4 => xl_rowcol_to_cell( $ro + 23, $co ),
                      IV5 => xl_rowcol_to_cell( $ro + 24, $co );
                };
                my $calcG = sub {
                    '', $softFormat, $formula->[2],
                      IV1 => xl_rowcol_to_cell( $ro + 26, $co ),
                      IV2 => xl_rowcol_to_cell( $ro + 30, $co );
                };
                my $calcH = sub {
                    '', $softFormat, $formula->[4],
                      IV1 => xl_rowcol_to_cell( $ro + 25, $co ),
                      IV2 => xl_rowcol_to_cell( $ro + 31, $co );
                };
                my $calcI = sub {
                    '', $softFormat, $formula->[2],
                      IV1 => xl_rowcol_to_cell( $ro + 33, $co ),
                      IV2 => xl_rowcol_to_cell( $ro + 36, $co );
                };
                my $calcJ = sub {
                    '', $softFormat, $formula->[1],
                      IV1 => xl_rowcol_to_cell( $ro + 32, $co ),
                      IV2 => xl_rowcol_to_cell( $ro + 37, $co );
                };
                my @responseArray = (
                    $dataEntry,
                    $dataEntryMaker->( $wb->getFormat('0.000hard') ),
                    $dataEntry, $calcA,
                    $dataEntry, $dataEntry,
                    $dataEntry, $dataEntry,
                    $dataEntry, $calcB,
                    $dataEntry, $dataEntry,
                    $dataEntry, $dataEntry,
                    $dataEntry, $dataEntry,
                    $dataEntry, $dataEntry,
                    $dataEntry, $dataEntry,
                    $dataEntry, $dataEntry,
                    $calcC,     $dataEntry,
                    $dataEntry, $calcF,
                    $dataEntry, $dataEntry,
                    $dataEntry, $dataEntry,
                    $dataEntry, $calcG,
                    $calcH,     $dataEntry,
                    $dataEntry, $dataEntry,
                    $dataEntry, $calcI,
                    $calcJ,
                );
                sub {
                    my ( $x, $y ) = @_;

           # The following trick only works if the object is within a Columnset.
                    ( $sh, $ro, $co ) = $self->wsWrite( $wb, $ws )
                      unless defined $ro;

                    goto &{ $responseArray[$y] };
                };
            }
        );

        my $rowFormatsc = [ map { /=/ ? $textnocolourbc : undef; } @labels ];

        $model->{inputTable1001} ||= [];

        Columnset(
            name     => 'CDCM target revenue',
            number   => 1001,
            appendTo => $model->{inputTable1001},
            dataset  => $model->{dataset},
            columns  => [
                Dataset(
                    name          => 'Further description',
                    rows          => $labelset,
                    defaultFormat => 'textnocolour',
                    rowFormats =>
                      [ map { /=/ ? $textnocolourb : undef; } @labels ],
                    data            => \@descriptions,
                    useIllustrative => 1,
                ),
                Dataset(
                    name          => 'Term',
                    rows          => $labelset,
                    defaultFormat => $textnocolourc,
                    rowFormats    => $rowFormatsc,
                    data          => \@term,
                ),
                Dataset(
                    name          => 'CRC',
                    rows          => $labelset,
                    defaultFormat => $textnocolourc,
                    rowFormats    => $rowFormatsc,
                    data          => \@crc,
                ),
                $numbers
            ]
        );

        return Stack(
            name          => 'Target CDCM revenue (£/year)',
            defaultFormat => '0copy',
            rows          => Labelset(
                list => [ $labelset->{list}[ $#{ $labelset->{list} } ] ]
            ),
            sources => [$numbers],
        );

    }

    {

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
            rows   => $labelset,
            custom => [
                '=IV1',         '=0-IV1',
                '=IV1*(IV2-1)', $handSubtotal ? () : '=SUBTOTAL(9,IV3:IV4)'
            ],
            arithmetic => '',
            objectType => 'Mixed calculations',
            wsPrepare  => sub {
                my ( $self, $wb, $ws, $format, $formula ) = @_;
                unless ( exists $avoidDoublePush{$wb} ) {
                    push @{ $self->{location}{postWriteCalls}{$wb} }, sub {
                        $ws->write_string(
                            $ws->{nextFree}++,
                            0,
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

                    $handSubtotal = []
                      if !$y && $model->{targetRevenue} =~ /nosubtotal/i;

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
                        $handSubtotal
                          ? (
                            '',
                            $boldFormat,
                            '='
                              . (
                                join '+',
                                map { xl_rowcol_to_cell( $ro + $_, $co ) }
                                  grep { $_ >= $startRowOffset; }
                                  @$handSubtotal
                              )
                          )
                          : (
                            '',
                            $boldFormat,
                            $formula->[3],
                            IV3 =>
                              xl_rowcol_to_cell( $ro + $startRowOffset, $co ),
                            IV4 => xl_rowcol_to_cell( $ro + $y - 1, $co ),
                          );    # NB: SUM(A,B,C,...) does not scale well
                    }
                    else {
                        push @$handSubtotal, $y if $handSubtotal;
                        $descriptions[$y] =~ /^A2/
                          ? (
                            '', $format, $formula->[2],
                            IV1 => xl_rowcol_to_cell( $roi,     $coi ),
                            IV2 => xl_rowcol_to_cell( $roi + 1, $coi ),
                          )
                          : $descriptions[$y] =~ /^(A3|I)/ ? (
                            '', $format, $formula->[1],
                            IV1 => xl_rowcol_to_cell( $roi + $y, $coi ),
                          )
                          : (
                            '', $format, $formula->[0],
                            IV1 => xl_rowcol_to_cell( $roi + $y, $coi ),
                          );
                    }
                };
            }
        );

        my $rowFormatsc =
          [ map { /=/ ? $textnocolourbc : undef; } @descriptions ];

        $model->{table1001} = Columnset(
            name     => 'CDCM target revenue',
            number   => 1001,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            columns  => [
                Dataset(
                    name          => 'Further description',
                    rows          => $labelset,
                    defaultFormat => 'textnocolour',
                    rowFormats =>
                      [ map { /=/ ? $textnocolourb : undef; } @descriptions ],
                    data            => \@descriptions,
                    useIllustrative => 1,
                ),
                Dataset(
                    name          => 'Term',
                    rows          => $labelset,
                    defaultFormat => $textnocolourc,
                    rowFormats    => $rowFormatsc,
                    data          => \@term,
                ),
                Dataset(
                    name          => 'CRC',
                    rows          => $labelset,
                    defaultFormat => $textnocolourc,
                    rowFormats    => $rowFormatsc,
                    data          => \@crc,
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
                join(
                    '+',
                    '=IV100*IV101-IV102',
                    (
                        map { "IV$_" } 104 .. 108,
                        110 .. 121,
                        123, 124, 126 .. 130
                    )
                  )
                  . '-IV133-IV134-IV135-IV136'
            ],
            arithmetic => '= derived from IV100',
            rows       => $specialRowset,
            objectType => 'Special calculation',
            arguments  => {
                map { ( "IV$_" => $inputs ); } 100 .. 102,
                104 .. 108,
                110 .. 121,
                123, 124,
                126 .. 130,
                133 .. 136,
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    '', $format, $formula->[0], map {
                        'IV'
                          . ( 100 + $_ ) =>
                          xl_rowcol_to_cell( $rowh->{IV100} + $_,
                            $colh->{IV100} );
                    } 0 .. 36;
                };
            },
        );

        $model->{edcmTables}[0][4] = new SpreadsheetModel::Custom(
            name => 'The amount of money that the DNO wants to raise from use'
              . ' of system charges, less transmission exit (£/year)',
            defaultFormat => '0hard',
            custom        => ['=IV1+IV2-IV3'],
            arithmetic    => '=IV1+IV2-IV3',
            objectType    => 'Special calculation',
            arguments     => {
                IV1 => $target,
                IV2 => $inputs,
                IV3 => $model->{edcmTables}[0][5]{sources}[0],
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    '', $format, $formula->[0],
                      IV1 => xl_rowcol_to_cell( $rowh->{IV1}, $colh->{IV1} ),
                      IV2 =>
                      xl_rowcol_to_cell( $rowh->{IV2} + 33, $colh->{IV2} ),
                      IV3 => xl_rowcol_to_cell( $rowh->{IV3}, $colh->{IV3} );
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
                    arguments     => { IV1 => $target, IV2 => $subtotals, },
                    arithmetic    => '=IV1-IV2'
                )
            ]
        );

        $target;

    }

}

1;
