package ModelM;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others.
Copyright 2012-2017 Franck Latrémolière, Reckon LLP and others.

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
DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTI
AL DAMAGES
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
                arithmetic    => '=A1*A2',
                arguments     => {
                    A1 => $meavPercentagesEdcm,
                    A2 => Stack(
                        name    => 'Allocation to EHV network levels',
                        cols    => $ehv132,
                        sources => [$alloc],
                    ),
                },
            ),
            $alloc,
        ],
    );
    my $discountLevelset = $model->{edcm} =~ /only/i
      ? Labelset( list => [ split /\n/, <<EOL] )
$model->{qno} HVplus: LV demand
$model->{qno} HVplus: LV Sub dem | LV gen
$model->{qno} HVplus: HV dem | LV Sub gen
$model->{qno} HVplus: HV generation
$model->{qno} EHV: LV demand
$model->{qno} EHV: LV Sub dem | LV gen
$model->{qno} EHV: HV dem | LV Sub gen
$model->{qno} EHV: HV generation
$model->{qno} 132kV/EHV: LV demand
$model->{qno} 132kV/EHV: LV Sub dem | LV gen
$model->{qno} 132kV/EHV: HV dem | LV Sub gen
$model->{qno} 132kV/EHV: HV generation
$model->{qno} 132kV: LV demand
$model->{qno} 132kV: LV Sub dem | LV gen
$model->{qno} 132kV: HV dem | LV Sub gen
$model->{qno} 132kV: HV generation
$model->{qno} 0000: LV demand
$model->{qno} 0000: LV Sub dem | LV gen
$model->{qno} 0000: HV dem | LV Sub gen
$model->{qno} 0000: HV generation
EOL
      : Labelset( list => [ split /\n/, <<EOL] );
$model->{qno} LV: LV demand
$model->{qno} HV: LV demand
$model->{qno} HV: LV Sub demand
$model->{qno} HV: HV demand
$model->{qno} HVplus: LV demand
$model->{qno} HVplus: LV Sub dem | LV gen
$model->{qno} HVplus: HV dem | LV Sub gen
$model->{qno} HVplus: HV generation
$model->{qno} EHV: LV demand
$model->{qno} EHV: LV Sub dem | LV gen
$model->{qno} EHV: HV dem | LV Sub gen
$model->{qno} EHV: HV generation
$model->{qno} 132kV/EHV: LV demand
$model->{qno} 132kV/EHV: LV Sub dem | LV gen
$model->{qno} 132kV/EHV: HV dem | LV Sub gen
$model->{qno} 132kV/EHV: HV generation
$model->{qno} 132kV: LV demand
$model->{qno} 132kV: LV Sub dem | LV gen
$model->{qno} 132kV: HV dem | LV Sub gen
$model->{qno} 132kV: HV generation
$model->{qno} 0000: LV demand
$model->{qno} 0000: LV Sub dem | LV gen
$model->{qno} 0000: HV dem | LV Sub gen
$model->{qno} 0000: HV generation
EOL
    my $dataAtwBypass = [ map { [ split /\s+/ ] } split /\n/, <<EOT];


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
    my $dataDnoBypass = [ map { [ split /\s+/ ] } split /\n/, <<EOT];
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

    unless ( $model->{dcp095} ) {
        shift @$_ foreach @$dataAtwBypass, @$dataDnoBypass;
    }
    if ( $model->{edcm} =~ /only/i ) {
        splice @$_, 0, 4 foreach $dataAtwBypass, $dataDnoBypass;
    }
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
            data          => $dataAtwBypass,
        ),
    );
    my $dnoBypassMatrix = Constant(
        name          => 'Network levels not covered by DNO network',
        defaultFormat => '0con',
        rows          => $discountLevelset,
        cols          => $allLevelset,
        byrow         => 1,
        data          => $dataDnoBypass,
    );

    my $nameOfLvLevel = $model->{dcp095} ? 'LV mains' : 'LV';

    if (
        grep {
            grep { /split/ }
              @$_
        } @{ $dnoBypassMatrix->{data} }
      )
    {
        my $splitDirect = SpreadsheetModel::Custom->new(
            name          => 'Splitting factors',
            defaultFormat => '%soft',
            cols =>
              Labelset( list => [ $nameOfLvLevel, 'HV', 'EHV', '132kV' ] ),
            custom     => [ '=1-A2*A9', '=1-A3*A9', '=1-A9' ],
            arithmetic => 'Special calculation',
            arguments  => {
                A2 => $model->lvSplit,
                A3 => $model->hvSplit,
                A9 => $direct,
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    return '', $format, $formula->[0], map {
                        qr/\b$_\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{$_}, $colh->{$_}, )
                    } @$pha if !$x;
                    return '', $format, $formula->[1], map {
                        qr/\b$_\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{$_}, $colh->{$_} + ( /A9/ ? 2 : 0 ),
                          )
                    } @$pha if $x == 1;
                    '', $format, $formula->[2], map {
                        qr/\b$_\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{$_}, $colh->{$_} + ( /A9/ ? 3 : 0 ),
                          )
                    } @$pha;
                };
            },
        );
        my @sources = $dnoBypassMatrix;
        foreach
          my $level ( $model->{edcm} =~ /only/i ? () : ( $nameOfLvLevel, 'HV' ),
            'EHV', '132kV' )
        {
            my ($col) = grep { $allLevelset->{list}[$_] eq $level }
              0 .. $#{ $allLevelset->{list} };
            die "$level not found in @{$allLevelset->{list}}"
              unless defined $col;
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
                arithmetic    => '=A1',
                arguments     => { A1 => $splitDirect },
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

    my $discounts = Arithmetic(
        name          => $model->{qno} . ' discounts (EDCM)',
        defaultFormat => '%soft',
        arithmetic    => '=1-MAX(0,(1-A1)/(1-A2))',
        arguments     => {
            A2 => $atwBypassed,
            A1 => $dnoBypassed,
        },
    );

    if ( $model->{not1181layout} ) {
        push @{ $model->{objects}{calcSheets} },
          [ $model->{suffix}, $atwBypassed, $dnoBypassed ];
        $discounts = Columnset(
            name    => $model->{qno} . ' discounts (EDCM)',
            columns => [
                $discounts,
                map {
                    my $digits = /([0-9])/ ? $1 : 6;
                    SpreadsheetModel::Checksum->new(
                        name => $_,
                        /table|recursive|model/i ? ( recursive => 1 ) : (),
                        digits  => $digits,
                        columns => [$discounts],
                        factors => [10000]
                    );
                  } split /;\s*/,
                $model->{checksums}
            ],
        ) if $model->{checksums};
    }
    else {
        push @{ $model->{objects}{calcSheets} },
          [ $model->{suffix}, $discounts ];
        my $ldnoLevelset = Labelset( list => [ split /\n/, <<EOL] );
Boundary 0000
Boundary 132kV
Boundary 132kV/EHV
Boundary EHV
Boundary HVplus
EOL
        my @cdcmLevels = split /\n/, <<EOL;
LV demand
LV Sub demand or LV generation
HV demand or LV Sub generation
HV generation
EOL
        my @columns;
        my $offset = $model->{edcm} =~ /only/i ? 0 : 4;
        foreach ( my $i = 0 ; $i < 4 ; ++$i ) {
            my $iForClosure = $i;
            $columns[$i] = SpreadsheetModel::Custom->new(
                name          => $cdcmLevels[$i],
                rows          => $ldnoLevelset,
                defaultFormat => '%copy',
                arithmetic    => '=A1',
                custom        => ['=A1'],
                arguments     => { A1 => $discounts, },
                wsPrepare     => sub {
                    my ( $self, $wb, $ws, $format, $formula, $pha, $rowh,
                        $colh ) = @_;
                    sub {
                        my ( $x, $y ) = @_;
                        '', $format, $formula->[0], map {
                            qr/\b$_\b/ =>
                              Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                                $rowh->{$_} +
                                  $offset + $iForClosure +
                                  4 * ( 4 - $y ),
                                $colh->{$_},
                              )
                        } @$pha;
                    };
                },
            );
        }
        push @columns, map {
            my $digits = /([0-9])/ ? $1 : 6;
            SpreadsheetModel::Checksum->new(
                name => $_,
                /table|recursive|model/i ? ( recursive => 1 ) : (),
                digits  => $digits,
                columns => [@columns],
                factors => [ map { 10000 } 1 .. 4 ]
            );
        } split /;\s*/, $model->{checksums} if $model->{checksums};
        $discounts = Columnset(
            name => $model->{qno} . ' discounts (EDCM) ⇒1181. For EDCM model',
            columns => \@columns,
        );
        push @{ $model->{objects}{table1181sources} }, @columns;
    }
    push @{ $model->{objects}{resultsTables} },
      $model->{objects}{table1181columnset} = $discounts;
    $discounts;
}

1;
