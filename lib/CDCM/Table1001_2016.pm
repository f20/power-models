package CDCM;

=head Copyright licence and disclaimer

Copyright 2012-2016 Franck Latrémolière, Reckon LLP and others.

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

sub table1001_2016 {

    my ($model) = @_;

    my @lines = map { [ split /\|/, $_, -1 ] } split /\n/, <<EOL;
Base Demand Revenue before inflation|A1|PU|CRC2A
Annual Iteration adjustment before inflation|A2|MOD|CRC2A
RPI True-up before inflation|A3|TRU|CRC2A
Price index adjustment (RPI index)|A4|RPIF|CRC2A
Base demand revenue|A = (A1 + A2 + A3) * A4|BR|CRC2A
Pass-Through Licence Fees|B1|LF|CRC2B
Pass-Through Business Rates|B2|RB|CRC2B
Pass-Through Transmission Connection Point Charges|B3|TB|CRC2B
Pass-through Smart Meter Communication Licence Costs|B4|SMC|CRC2B
Pass-through Smart Meter IT Costs|B5|SMIT|CRC2B
Pass-through Ring Fence Costs|B6|RF|CRC2B
Pass-Through Others|B7|HB, SEC, UNC|CRC2B
Allowed Pass-Through Items|B = Sum of B1 to B7|PT|CRC2B
Broad Measure of Customer Service incentive|C1|BM|CRC2C
Quality of Service incentive|C2|IQ|CRC2D
Connections Engagement incentive|C3|ICE|CRC2E
Time to Connect incentive|C4|TTC|CRC2F
Losses Discretionary Reward incentive|C5|LDR|CRC2G
Network Innovation Allowance|C6|NIA|CRC2H
Low Carbon Network Fund - Tier 1 unrecoverable|C7|LCN1|CRC2J
Low Carbon Network Fund - Tier 2 & Discretionary Funding|C7|LCN2|CRC2J
Connection Guaranteed Standards Systems & Processes penalty|C8|AUM, CGSRA|CRC2K-L
Residual Losses and Growth Incentive - Losses|C9|PPL|CRC2M
Residual Losses and Growth Incentive - Growth|C9|GTA|CRC2M
Incentive Revenue and Other Adjustments|C = Sum of C1 to C9
Correction Factor|D|-K|CRC2A
Total allowed Revenue|E = A + B + C + D|AR|CRC2A
Other 1. Excluded services - Top-up, standby, and enhanced system security|F1 (see note 1)|DRS4|CRC5C
Other 2. Excluded services - Revenue protection services|F2 (see note 1)|DRS5|CRC5C
Other 3. Excluded services - Miscellaneous|F3 (see note 1)|DRS9|CRC5C
Other 4. Please describe if used|F4|Please describe|if used
Other 5. Please describe if used|F5|Please describe|if used
Total other revenue recovered by Use of System Charges|F = Sum of F1 to F5
Total Revenue for Use of System Charges|G = E + F
1. Revenue raised outside CDCM - EDCM and Certain Interconnector Revenue|H1
2. Revenue raised outside CDCM - Voluntary under-recovery|H2
3. Revenue raised outside CDCM - Please describe if used|H3|Please describe|if used
4. Revenue raised outside CDCM - Please describe if used|H4|Please describe|if used
Total Revenue to be raised outside the CDCM|H = Sum of H1 to H4
Latest forecast of CDCM Revenue|I = G - H
EOL

    my $labelset = Labelset( list => [ map { $_->[0] } @lines ] );
    my $textnocolourc = [ base => 'textnocolour', align => 'center' ];
    my $textnocolourb = [ base => 'textnocolour', bold  => 1, ];
    my $textnocolourbc =
      [ base => 'textnocolour', bold => 1, align => 'center' ];
    my $moneyInputFormat =
      $model->{targetRevenue} =~ /million/i
      ? 'millionhard'
      : '0hard';

    my $inputs = Dataset(
        name       => 'Value',
        rows       => $labelset,
        rowFormats => [
            map {
                    $_->[1] =~ /^A4/ ? '0.000hard'
                  : $_->[1] =~ /=/   ? undef
                  :                    $moneyInputFormat;
            } @lines
        ],
        data => [ map { $_->[1] =~ /=/ ? undef : ''; } @lines ],
    );

    my %avoidDoublePush;

    my $calculations = new SpreadsheetModel::Custom(
        name          => 'Calculations (£/year)',
        defaultFormat => [
            base => $model->{targetRevenue} =~ /million/i
            ? 'millionsoft'
            : '0soft',
            bold => 1
        ],
        rows       => $labelset,
        custom     => [],
        arithmetic => '',
        objectType => 'Mixed calculations',
        wsPrepare  => sub {
            my ( $self, $wb, $ws, $format, $formula ) = @_;
            unless ( exists $avoidDoublePush{$wb} ) {
                push @{ $self->{location}{postWriteCalls}{$wb} }, sub {
                    my ( $me, $wbMe, $wsMe, $rowrefMe, $colMe ) = @_;
                    $wsMe->write_string(
                        $$rowrefMe += 2,
                        $colMe - 1,
                        'Note 1: Cost categories associated with excluded '
                          . 'services should only be populated if the Company '
                          . 'recovers the costs of providing these services '
                          . 'from Use of System Charges',
                        $wb->getFormat('text')
                    );
                };
                undef $avoidDoublePush{$wb};
            }
            my ( $inputCell, $calcCell );
            my $unavailable = $wb->getFormat('unavailable');
            sub {
                my ( $x, $y ) = @_;
                local $_ = $lines[$y][1];
                return '', $unavailable unless /=/;
                unless ($inputCell) {
                    my ( $sh, $ro, $co ) = $inputs->wsWrite( $wb, $ws );
                    $inputCell = sub { xl_rowcol_to_cell( $ro + $_[0], $co ); };
                }
                unless ($calcCell) {
                    my ( $sh, $ro, $co ) = $self->wsWrite( $wb, $ws );
                    $calcCell = sub { xl_rowcol_to_cell( $ro + $_[0], $co ); }
                }
                return
                    '=('
                  . $inputCell->(0) . '+'
                  . $inputCell->(1) . '+'
                  . $inputCell->(2) . ')*'
                  . $inputCell->(3), $format
                  if /^A/;
                return '=SUM(' . $inputCell->(5) . ':' . $inputCell->(11) . ')',
                  $format
                  if /^B/;
                return
                    '=SUM('
                  . $inputCell->(13) . ':'
                  . $inputCell->(23)
                  . ')', $format
                  if /^C/;
                return
                    '='
                  . $calcCell->(4) . '+'
                  . $calcCell->(12) . '+'
                  . $calcCell->(24) . '+'
                  . $inputCell->(25), $format
                  if /^E/;
                return
                    '=SUM('
                  . $inputCell->(27) . ':'
                  . $inputCell->(31)
                  . ')', $format
                  if /^F/;
                return '=' . $calcCell->(26) . '+' . $calcCell->(32), $format
                  if /^G/;
                return
                    '=SUM('
                  . $inputCell->(34) . ':'
                  . $inputCell->(37)
                  . ')', $format
                  if /^H/;
                return '=' . $calcCell->(33) . '-' . $calcCell->(38), $format
                  if /^I/;
                die "Unexpected: $_";
            };
        },
    );

    my $rowFormatsc =
      [ map { $_->[1] =~ /=/ ? $textnocolourbc : undef; } @lines ];

    $model->{table1001_2016} = Columnset(
        name     => 'CDCM target revenue (£ unless otherwise stated)',
        number   => 1001,
        appendTo => $model->{inputTables},
        dataset  => $model->{dataset},
        columns  => [
            Constant(
                name          => 'Further description',
                rows          => $labelset,
                defaultFormat => 'textnocolour',
                rowFormats =>
                  [ map { $_->[1] =~ /=/ ? $textnocolourb : undef; } @lines ],
                data => [ map { $_->[1] } @lines ],
            ),
            Dataset(
                name               => 'Term',
                rows               => $labelset,
                defaultFormat      => $textnocolourc,
                rowFormats         => $rowFormatsc,
                data               => [ map { $_->[2] } @lines ],
                usePlaceholderData => 1,
            ),
            Dataset(
                name               => 'CRC',
                rows               => $labelset,
                defaultFormat      => $textnocolourc,
                rowFormats         => $rowFormatsc,
                data               => [ map { $_->[3] } @lines ],
                usePlaceholderData => 1,
            ),
            $inputs,
            $calculations,
        ]
    );

    my $specialRowset =
      Labelset( list => [ $labelset->{list}[ $#{ $labelset->{list} } ] ] );

    my $target = new SpreadsheetModel::Custom(
        name          => 'Target CDCM revenue (£/year)',
        defaultFormat => '0soft',
        custom        => [
            join( '+',
                '=(A100+A101+A102)*A103',
                ( map { "A$_" } 105 .. 111, 113 .. 123, 125, 127 .. 131 ) )
              . '-A134-A135-A136-A137'
        ],
        arithmetic => '= derived from A100',
        rows       => $specialRowset,
        arguments  => { map { ( "A$_" => $inputs ); } 100 .. 137, },
        wsPrepare  => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0], map {
                    my $c = 'A' . ( 100 + $_ );
                    qr/\b$c\b/ =>
                      xl_rowcol_to_cell( $rowh->{A100} + $_, $colh->{A100} );
                } 0 .. 37;
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
                  qr/\bA2\b/ =>
                  xl_rowcol_to_cell( $rowh->{A2} + 33, $colh->{A2} );
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
                arguments     => { A1 => $target, A2 => $calculations, },
                arithmetic    => '=A1-A2'
            )
        ]
    );

    $target;

}

1;
