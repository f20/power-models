package CDCM;

=head Copyright licence and disclaimer

Copyright 2016 Franck Latrémolière, Reckon LLP and others.

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

sub modelG {

    my ( $model, $nonExcludedComponents, $daysAfter, $volumeDataPcd,
        $allEndUsers, @utaTables, )
      = @_;

    my @utaPoundsOriginal = map {
        $model->applyVolumesToTariff(
            $nonExcludedComponents,
            $_,
            $volumeDataPcd,
            $daysAfter,
            'Baseline revenues from '
              . lcfirst $_->{name}
              . ' charges (£/year)'
        );
    } @utaTables;
    my @originalTotals = map {
        GroupBy(
            name          => $_->objectShortName,
            defaultFormat => '0soft',
            source        => $_,
        );
    } @utaPoundsOriginal;
    push @{ $model->{modelgTables} },
      Columnset(
        name    => 'Unrounded revenue analysis (baseline)',
        columns => \@utaPoundsOriginal,
      ),
      Columnset(
        name    => 'Unrounded revenue analysis (baseline totals)',
        columns => \@originalTotals,
      );

    my %volumeData      = %{ $model->{pcd}{volumeData} };
    my $ldnoGenLabelset = Labelset(
        list => [
            grep { /ldno.*gener/i; }
              @{ $model->{pcd}{allTariffsByEndUser}{list} }
        ]
    );
    $volumeData{'Fixed charge p/MPAN/day'} = Stack(
        name          => 'MPANs excluding LDNO generation',
        rows          => $volumeData{'Fixed charge p/MPAN/day'}{rows},
        defaultFormat => '0copy',
        sources       => [
            Constant(
                name          => '0 for LDNO generation',
                defaultFormat => '0con',
                rows          => $ldnoGenLabelset,
                data => [ [ map { 0 } @{ $ldnoGenLabelset->{list} } ] ],
            ),
            $volumeData{'Fixed charge p/MPAN/day'},
        ],
    );

    my @utaPoundsBeforeReordering = map {
        $model->applyVolumesToTariff( $nonExcludedComponents, $_, \%volumeData,
            $daysAfter,
            'Revenues from ' . lcfirst $_->{name} . ' charges (£/year)' );
    } @utaTables;
    push @{ $model->{modelgTables} },
      Columnset(
        name    => 'Unrounded revenue analysis',
        columns => \@utaPoundsBeforeReordering,
      );

    my @groups;
    foreach ( @{ $model->{pcd}{allTariffsByEndUser}{groups} } ) {
        if (/related mpan/i) {
            my $last = pop @groups;
            my $name = "$last";
            $name =~ s/.*\n//s;
            push @groups,
              Labelset(
                name => "$name and related MPAN tariffs",
                list => [ @{ $last->{list} }, @{ $_->{list} }, ]
              );
        }
        else {
            push @groups, $_;
        }
    }
    my $regroupedTariffset = Labelset( groups => \@groups );
    my $groupset = Labelset( list => $regroupedTariffset->{groups} );
    my @utaPounds =
      map { Stack( sources => [$_], rows => $regroupedTariffset, ); }
      @utaPoundsBeforeReordering;
    my $units = Arithmetic(
        name          => 'Total MWh',
        defaultFormat => '0soft',
        rows          => $regroupedTariffset,
        arithmetic    => '='
          . join( '+', map { "A$_" } 1 .. $model->{maxUnitRates} ),
        arguments => {
            map { ( "A$_" => $volumeData{"Unit rate $_ p/kWh"} ); }
              1 .. $model->{maxUnitRates}
        },
    );
    Columnset(
        name    => 'Unrounded revenue analysis (with reordered tariff list)',
        columns => [ @utaPounds, $units, ],
    );

    my @groupedRevenueElements = map {
        GroupBy(
            name          => 'Grouped ' . lcfirst( $_->objectShortName ),
            defaultFormat => '0soft',
            rows          => $groupset,
            source        => $_,
          )
    } @utaPounds;
    my $groupedUnits = GroupBy(
        name          => 'Grouped units (MWh)',
        defaultFormat => '0soft',
        rows          => $groupset,
        source        => $units,
    );
    Columnset(
        name    => 'Unrounded revenue analysis (by tariff group)',
        columns => [ @groupedRevenueElements, $groupedUnits ],
    );

    my $ppuDiscounts =
      $model->{pcdByTariff}
      ? Dataset(
        name       => 'LDNO discounts (p/kWh)',
        number     => 1039,
        appendTo   => $model->{inputTables},
        dataset    => $model->{dataset},
        rows       => $regroupedTariffset,
        validation => {
            validate      => 'decimal',
            criteria      => '>=',
            value         => 0,
            input_title   => 'Discount:',
            input_message => 'p/kWh',
            error_title   => 'Invalid discount',
            error_message => 'The discount must be'
              . ' a non-negative percentage value.',
        },
        data => [
            map { /gener/i ? undef : /^LDNO/ ? 1 : 0; }
              @{ $regroupedTariffset->{list} },
        ],
      )
      : SumProduct(
        name          => 'Discount for each tariff (except for fixed charges)',
        defaultFormat => '%softnz',
        matrix        => Stack(
            name          => 'Discount map (re-grouped)',
            defaultFormat => '0copy',
            sources       => [ $model->{pcd}{discount}{matrix} ],
            cols          => $model->{pcd}{discount}{matrix}{cols},
            rows          => $regroupedTariffset,
        ),
        vector => Dataset(
            name       => 'LDNO discounts (p/kWh)',
            number     => 1039,
            appendTo   => $model->{inputTables},
            dataset    => $model->{dataset},
            cols       => $model->{pcd}{discount}{matrix}{cols},
            validation => {
                validate      => 'decimal',
                criteria      => '>=',
                value         => 0,
                input_title   => 'Discount:',
                input_message => 'p/kWh',
                error_title   => 'Invalid discount',
                error_message => 'The discount must be'
                  . ' a non-negative percentage value.',
            },
            data => [
                map {
                        /^no/i         ? undef
                      : /LDNO LV: LV/i ? 1
                      : /LDNO HV: LV/i ? 2
                      :                  1.5;
                } @{ $model->{pcd}{discount}{matrix}{cols}{list} },
            ],
        ),
      );

    my $calcErrors = sub {

        my (@scalingFactors) = @_;

        my $ppu = Arithmetic(
            name       => 'Average p/kWh',
            arithmetic => '=IF(A91,('
              . join( '+',
                map { $_ == 1 ? 'A1' : "A$_*A93$_" }
                  1 .. @groupedRevenueElements )
              . ')/A92*0.1,0)',
            arguments => {
                A91 => $groupedUnits,
                A92 => $groupedUnits,
                (
                    map { ( "A$_" => $groupedRevenueElements[ $_ - 1 ] ); }
                      1 .. @groupedRevenueElements
                ),
                (
                    map { ( "A93$_" => $scalingFactors[ $_ - 2 ] ); }
                      2 .. @groupedRevenueElements
                ),
            },
        );
        my $chargeablePercentage = Arithmetic(
            name          => 'Chargeable percentage',
            defaultFormat => '%soft',
            arithmetic    => '=IF(A21,1-A1/A22,0)',
            arguments     => {
                A1  => $ppuDiscounts,
                A21 => $ppu,
                A22 => $ppu,
            },
        );
        my @discountedCharges = map {
            SumProduct(
                name          => $_->objectShortName,
                defaultFormat => '0soft',
                matrix        => $chargeablePercentage,
                vector        => $_,
            );
        } @utaPounds;

        push @{ $model->{modelgTables} },
          Columnset(
            name    => 'Total discounted revenue by charge category',
            columns => \@discountedCharges,
          );

        my @errors = map {
            Arithmetic(
                name          => "Error $_",
                defaultFormat => '0softpm',
                arithmetic    => '=A1*A2-A3',
                arguments     => {
                    A1 => $discountedCharges[$_],
                    A2 => $scalingFactors[ $_ - 1 ],
                    A3 => $originalTotals[$_],
                },
            );
        } 1 .. ( $#discountedCharges - 1 );
        push @errors,
          Arithmetic(
            name => 'Error ' . ( @discountedCharges - 1 ),
            defaultFormat => '0softpm',
            arithmetic =>
              join( '+', '=A1*A2-A3+A10-A30', map { "A9$_" } 0 .. $#errors ),
            arguments => {
                A1  => $discountedCharges[$#discountedCharges],
                A2  => $scalingFactors[ $#discountedCharges - 1 ],
                A3  => $originalTotals[$#discountedCharges],
                A10 => $discountedCharges[0],
                A30 => $originalTotals[0],
                map { ( "A9$_" => $errors[$_] ); } 0 .. $#errors,
            },
          );
        @errors;
    };

    # hard-coded three dimensional

    my @runs;
    foreach ( my $run = 1 ; $run < 5 ; ++$run ) {
        my @scalingFactors = map {
            Constant(
                name => 'Scaling factor for '
                  . lcfirst( $utaPounds[$_]->objectShortName ),
                data => [ [ $_ < $run ? .99 : 1 ] ],
            );
        } 1 .. $#utaPounds;
        my @errors = $calcErrors->(@scalingFactors);
        push @runs, [ \@scalingFactors, \@errors, ];
        push @{ $model->{modelgTables} },
          Columnset(
            name    => "Scaling factors for run $run",
            columns => \@scalingFactors,
          ),
          Columnset(
            name    => "Error values from run $run",
            columns => \@errors,
          );
    }

    my $matrixLabelset = Labelset( list => [qw(X Y Z)] );
    my $iterate = sub {
        my $derivatives = SpreadsheetModel::Custom->new(
            name   => 'First derivatives (£ million)',
            rows   => $matrixLabelset,
            cols   => $matrixLabelset,
            custom => [
                '=1e-6*(B9-B8)/(B3-B2)',   '=1e-6*(C9-C8)/(B3-B2)',
                '=1e-6*(D9-D8)/(B3-B2)',   '=1e-6*(B10-B9)/(C4-C3)',
                '=1e-6*(C10-C9)/(C4-C3)',  '=1e-6*(D10-D9)/(C4-C3)',
                '=1e-6*(B11-B10)/(D5-D4)', '=1e-6*(C11-C10)/(D5-D4)',
                '=1e-6*(D11-D10)/(D5-D4)',
            ],
            arithmetic => '= Special calculation',
            arguments  => {
                (
                    map {
                        (
                            "B$_" => $runs[ $_ - 2 ][0][0],
                            "C$_" => $runs[ $_ - 2 ][0][1],
                            "D$_" => $runs[ $_ - 2 ][0][2],
                        );
                    } 2 .. 5
                ),
                (
                    map {
                        (
                            "B$_" => $runs[ $_ - 8 ][1][0],
                            "C$_" => $runs[ $_ - 8 ][1][1],
                            "D$_" => $runs[ $_ - 8 ][1][2],
                        );
                    } 8 .. 11
                ),
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    '', $format, $formula->[ 3 * $y + $x ], map {
                        qr/\b$_\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{$_}, $colh->{$_} )
                    } @$pha;
                };
            },
        );
        my $codeterminants = SpreadsheetModel::Custom->new(
            name   => 'Co-determinants',
            rows   => $matrixLabelset,
            cols   => $matrixLabelset,
            custom => [
                '=C15*D16-D15*C16', '=D15*B16-B15*D16',
                '=B15*C16-C15*B16', '=D14*C16-C14*D16',
                '=B14*D16-D14*B16', '=C14*B16-B14*C16',
                '=C14*D15-D14*C15', '=D14*B15-B14*D15',
                '=B14*C15-C14*B15',
            ],
            arithmetic => '= Special calculation',
            arguments  => {
                A1 => $derivatives,
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                my @mappings = map {
                    my $letter = qw(B C D) [$_];
                    my $x = $colh->{A1} + $_;
                    map {
                        qr/\b$letter$_\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A1} + $_ - 14, $x, );
                    } 14 .. 16;
                } 0 .. 2;
                sub {
                    my ( $x, $y ) = @_;
                    '', $format, $formula->[ 3 * $y + $x ], @mappings;
                };
            },
        );
        my $determinant = SpreadsheetModel::Custom->new(
            name      => 'Determinant',
            custom    => ['=SUMPRODUCT(B14:D14,B19:D19)'],
            arguments => {
                B14 => $derivatives,
                D14 => $derivatives,
                B19 => $codeterminants,
                D19 => $codeterminants,
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    '', $format, $formula->[0], map {
                        qr/\b$_\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{$_},
                            $colh->{$_} +
                              ( /^B/ ? 0 : /^C/ ? 1 : /^D/ ? 2 : die ),
                          )
                    } @$pha;
                };
            },
        );
        map {
            my $offset = $_;
            SpreadsheetModel::Custom->new(
                name => 'New scaling factor ' . ( 1 + $offset ),
                custom    => ['=B1-1e-6*(B11*B19+C11*C19+D11*D19)/B24'],
                arguments => {
                    B1  => $runs[3][0][$offset],
                    B24 => $determinant,
                    B11 => $runs[3][1][0],
                    C11 => $runs[3][1][1],
                    D11 => $runs[3][1][2],
                    B19 => $codeterminants,
                    C19 => $codeterminants,
                    D19 => $codeterminants,
                },
                wsPrepare => sub {
                    my ( $self, $wb, $ws, $format, $formula, $pha, $rowh,
                        $colh ) = @_;
                    sub {
                        my ( $x, $y ) = @_;
                        '', $format, $formula->[0], map {
                            qr/\b$_\b/ =>
                              Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                                $rowh->{$_} + ( /19$/ ? $offset : 0 ),
                                $colh->{$_} + ( /^C19$/ ? 1 : /^D19$/ ? 2 : 0 ),
                              )
                        } @$pha;
                    };
                },
            );
        } 0 .. 2;
    };

    my $runOffset = 0;

    my @nextScalingFactors = $iterate->();
    splice @runs, 0, 3;
    $runOffset += 3;
    foreach ( my $run = 2 ; $run < 5 ; ++$run ) {
        my @scalingFactors = $run == 4 ? @nextScalingFactors : map {
            Stack(
                sources => [
                      $_ < $run
                    ? $nextScalingFactors[ $_ - 1 ]
                    : $runs[0][0][ $_ - 1 ]
                ],
            );
        } 1 .. $#utaPounds;
        my @errors = $calcErrors->(@scalingFactors);
        push @runs, [ \@scalingFactors, \@errors, ];
        my $runa = $run + $runOffset;
        push @{ $model->{modelgTables} },
          Columnset(
            name    => "Scaling factors for run $runa",
            columns => \@scalingFactors,
          ),
          Columnset(
            name    => "Error values from run $runa",
            columns => \@errors,
          );
    }

    my @finalScalingFactors = $iterate->();
    push @{ $model->{modelgTables} },
      Columnset(
        name    => 'Final scaling factors',
        columns => \@finalScalingFactors,
      );
    push @{ $model->{modelgTables} },
      my $ppu = Arithmetic(
        name       => 'All-the-way p/kWh',
        arithmetic => '=IF(A91,('
          . join( '+',
            map { $_ == 1 ? 'A1' : "A$_*A93$_" } 1 .. @groupedRevenueElements )
          . ')/A92*0.1,0)',
        arguments => {
            A91 => $groupedUnits,
            A92 => $groupedUnits,
            (
                map { ( "A$_" => $groupedRevenueElements[ $_ - 1 ] ); }
                  1 .. @groupedRevenueElements
            ),
            (
                map { ( "A93$_" => $finalScalingFactors[ $_ - 2 ] ); }
                  2 .. @groupedRevenueElements
            ),
        },
      );

    push @{ $model->{modelgTables} },
      my $discounts = Arithmetic(
        name          => 'LDNO discounts',
        defaultFormat => '%soft',
        arithmetic    => '=IF(A21,A1/A22,0)',
        arguments     => {
            A1  => $ppuDiscounts,
            A21 => $ppu,
            A22 => $ppu,
        },
      );

    push @{ $model->{modelgTables2} },
      Stack(
        name    => 'LDNO discounts ⇒1038. For CDCM',
        rows    => $model->{pcd}{allTariffsByEndUser},
        sources => [$discounts],
      );

    push @{ $model->{modelgTables} }, my $ppuReference = Arithmetic(
        name       => 'All-the-way reference p/kWh values',
        arithmetic => '=A1',
        arguments  => { A1 => $ppu },
        rows       => Labelset(
            groups => [
                map {
                    Labelset(
                        name => "$_",
                        list => [ grep { !/^LDNO/i } @{ $_->{list} } ]
                    );
                } @groups
            ]
        ),
    );

    push @{ $model->{modelgTables2} },
      Stack(
        name    => 'All-the-way reference p/kWh values ⇒1185. For EDCM model',
        rows    => $allEndUsers,
        sources => [$ppuReference],
      );

    $model;

}

1;
