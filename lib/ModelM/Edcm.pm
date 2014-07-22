package ModelM;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others.
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

sub discountEdcm {
    my ( $model, $alloc, $direct, ) = @_;
    my $edcmLevelset = Labelset(
        name => 'EDCM method M network levels',
        list => [qw(EHV/HV EHV 132kV/EHV 132kV)]
    );
    my ($meavPercentagesEdcm) = $model->meavPercentagesEdcm($edcmLevelset);
    my $ehv132 = Labelset(
        name => 'EHV and 132kV network level',
        list => [ grep { /EHV/ } @{ $alloc->{cols}{list} } ]
    );
    my $allLevelset = Labelset(
        name => 'All network levels',
        list => [
            map { /EHV/ ? @{ $edcmLevelset->{list} } : $_; }
              @{ $alloc->{cols}{list} }
        ]
    );
    my $allocAll = Stack(
        name          => 'Extended allocation',
        defaultFormat => '%copy',
        cols          => $allLevelset,
        sources       => [
            Arithmetic(
                name          => 'Allocation between EHV network levels',
                defaultFormat => '%soft',
                arithmetic    => '=IV1*IV2',
                arguments     => {
                    IV1 => $meavPercentagesEdcm,
                    IV2 => Stack(
                        name    => 'Allocation to EHV network levels',
                        cols    => $ehv132,
                        sources => [$alloc],
                    ),
                },
            ),
            $alloc,
        ],
    );
    my $discountLevelset = Labelset( list => [ split /\n/, <<EOL] );
LDNO LV: LV demand
LDNO HV: LV demand
LDNO HV: LV Sub demand
LDNO HV: HV demand
LDNO HVplus: LV demand
LDNO HVplus: LV Sub dem | LV gen
LDNO HVplus: HV dem | LV Sub gen
LDNO HVplus: HV generation
LDNO EHV: LV demand
LDNO EHV: LV Sub dem | LV gen
LDNO EHV: HV dem | LV Sub gen
LDNO EHV: HV generation
LDNO 132kV/EHV: LV demand
LDNO 132kV/EHV: LV Sub dem | LV gen
LDNO 132kV/EHV: HV dem | LV Sub gen
LDNO 132kV/EHV: HV generation
LDNO 132kV: LV demand
LDNO 132kV: LV Sub dem | LV gen
LDNO 132kV: HV dem | LV Sub gen
LDNO 132kV: HV generation
LDNO 0000: LV demand
LDNO 0000: LV Sub dem | LV gen
LDNO 0000: HV dem | LV Sub gen
LDNO 0000: HV generation
EOL
    my $atwBypassed = SumProduct(
        name => 'Proportion of costs not covered by all-the-way tariff',
        defaultFormat => '%soft',
        vector        => $allocAll,
        matrix        => Constant(
            name          => 'Network levels not covered by all-the-way tariff',
            defaultFormat => '0con',
            rows          => $discountLevelset,
            cols          => $allLevelset,
            byrow         => 1,
            data          => [ map { [ split /\s+/ ] } split /\n/, <<EOT], ) );


1 1
1 1 1

1 1
1 1 1
1 1 1 1

1 1
1 1 1
1 1 1 1

1 1
1 1 1
1 1 1 1

1 1
1 1 1
1 1 1 1

1 1
1 1 1
1 1 1 1
EOT
    my $dnoBypassMatrix = Constant(
        name          => 'Network levels not covered by DNO network',
        defaultFormat => '0con',
        rows          => $discountLevelset,
        cols          => $allLevelset,
        byrow         => 1,
        data          => [ map { [ split /\s+/ ] } split /\n/, <<EOT], );
1 split
1 1 1 split
1 1 1 split
1 1 1 split
1 1 1 1
1 1 1 1
1 1 1 1
1 1 1 1
1 1 1 1 1 split
1 1 1 1 1 split
1 1 1 1 1 split
1 1 1 1 1 split
1 1 1 1 1 1
1 1 1 1 1 1
1 1 1 1 1 1
1 1 1 1 1 1
1 1 1 1 1 1 1 split
1 1 1 1 1 1 1 split
1 1 1 1 1 1 1 split
1 1 1 1 1 1 1 split
1 1 1 1 1 1 1 1
1 1 1 1 1 1 1 1
1 1 1 1 1 1 1 1
1 1 1 1 1 1 1 1
EOT

    if (
        grep {
            grep { /split/ }
              @$_
        } @{ $dnoBypassMatrix->{data} }
      )
    {
        my @splits      = $model->splits;
        my $splitDirect = SpreadsheetModel::Custom->new(
            name          => 'Splitting factors',
            defaultFormat => '%soft',
            cols => Labelset( list => [ 'LV mains', 'HV', 'EHV', '132kV' ] ),
            custom => [ '=1-IV2*IV9', '=1-IV3*IV9', '=1-IV9' ],
            arguments =>
              { IV2 => $splits[0], IV3 => $splits[1], IV9 => $direct, },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    return '', $format, $formula->[0], map {
                        $_ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{$_}, $colh->{$_}, )
                    } @$pha if !$x;
                    return '', $format, $formula->[1], map {
                        $_ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{$_}, $colh->{$_} + ( /IV9/ ? 2 : 0 ),
                          )
                    } @$pha if $x == 1;
                    '', $format, $formula->[2], map {
                        $_ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{$_}, $colh->{$_} + ( /IV9/ ? 3 : 0 ),
                          )
                    } @$pha;
                };
            },
        );
        my @sources = $dnoBypassMatrix;
        foreach my $level ( 'LV mains', 'HV', 'EHV', '132kV' ) {
            my ($col) = grep { $allLevelset->{list}[$_] eq $level }
              0 .. $#{ $allLevelset->{list} };
            my $set = Labelset(
                list => [
                    @{ $dnoBypassMatrix->{rows}{list} }[
                      grep {
                               $dnoBypassMatrix->{data}[$_][$col]
                            && $dnoBypassMatrix->{data}[$_][$col] =~ /split/
                      } 0 .. $#{ $dnoBypassMatrix->{rows}{list} }
                    ]
                ]
            );
            unshift @sources,
              Arithmetic(
                name          => "Splitting factor $level",
                rows          => $set,
                cols          => Labelset( list => [$level] ),
                defaultFormat => '%copy',
                arithmetic    => '=IV1',
                arguments     => { IV1 => $splitDirect },
              );
        }
        $dnoBypassMatrix = Stack(
            name    => 'Network levels not covered by DNO network',
            rows    => $discountLevelset,
            cols    => $allLevelset,
            sources => \@sources,
        );
    }

    my $dnoBypassed = SumProduct(
        name          => 'Proportion of costs not covered by DNO network',
        defaultFormat => '%soft',
        vector        => $allocAll,
        matrix        => $dnoBypassMatrix,
    );
    push @{ $model->{calcTables} }, $allocAll, $atwBypassed, $dnoBypassed;
    my $discounts = Arithmetic(
        name          => 'Discount factors',
        defaultFormat => '%soft',
        arithmetic    => '=1-MAX(0,(1-IV1)/(1-IV2))',
        arguments     => {
            IV2 => $atwBypassed,
            IV1 => $dnoBypassed,
        },
    );
    push @{ $model->{impactTables} }, $discounts;
}

1;
