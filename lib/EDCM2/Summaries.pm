package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2012 Energy Networks Association Limited and others.
Copyright 2013-2017 Franck Latrémolière, Reckon LLP and others.

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

sub summaries {

    my (
        $model,                        $activeCoincidence935,
        $activeUnits,                  $actualRedDemandRate,
        $daysInYear,                   $exportCapacityChargeable,
        $fixedDchargeTrue,             $fixedDchargeTrueRound,
        $fixedGchargeTrue,             $genCreditRound,
        $hoursInPurple,                $importCapacity,
        $importCapacity935,            $importCapacityScaledRound,
        $netexportCapacityChargeRound, $previousChargeExport,
        $previousChargeImport,         $purpleRateFcpLricRound,
        $tariffDaysInYearNot,          $tariffHoursInPurpleNot,
        $tariffs,                      @tariffColumns,
    ) = @_;

    my $format0withLine =
      $model->{vertical}
      ? '0soft'
      : [ base => '0soft', left => 5, left_color => 8 ];

    my @revenueBitsD = (

        Arithmetic(
            name          => 'Capacity charge for demand (£/year)',
            defaultFormat => $format0withLine,
            arithmetic    => '=0.01*A9*A8*A1',
            arguments     => {
                A1 => $importCapacityScaledRound,
                A9 => $daysInYear,
                A7 => $tariffDaysInYearNot,
                A8 => $importCapacity,
            }
        ),

        Arithmetic(
            name => "$model->{TimebandName} charge for demand (£/year)",
            defaultFormat => '0softnz',
            arithmetic    => '=0.01*(A9-A7)*A1*A6*(A91/(A92-A71))*A8',
            arguments     => {
                A1  => $purpleRateFcpLricRound,
                A9  => $hoursInPurple,
                A7  => $tariffHoursInPurpleNot,
                A6  => $importCapacity,
                A8  => $activeCoincidence935,
                A91 => $daysInYear,
                A92 => $daysInYear,
                A71 => $tariffDaysInYearNot
            }
        ),

        Arithmetic(
            name          => 'Fixed charge for demand (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => '=0.01*(A9-A7)*A1',
            arguments     => {
                A1 => $fixedDchargeTrueRound,
                A9 => $daysInYear,
                A7 => $tariffDaysInYearNot,
            }
        ),

    );

    my @revenueBitsG = (

        Arithmetic(
            name => 'Net capacity charge (or credit) for generation (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => '=0.01*A9*A8*A1',
            arguments     => {
                A1 => $netexportCapacityChargeRound,
                A9 => $daysInYear,
                A8 => $exportCapacityChargeable,
            }
        ),

        Arithmetic(
            name          => 'Fixed charge for generation (£/year)',
            defaultFormat => '0softnz',
            arithmetic    => '=0.01*(A9-A7)*A1',
            arguments     => {
                A1 => $fixedGchargeTrue,
                A9 => $daysInYear,
                A7 => $tariffDaysInYearNot,
            }
        ),

        Arithmetic(
            name          => "$model->{TimebandName} credit (£/year)",
            defaultFormat => '0softnz',
            arithmetic    => '=0.01*A1*A6',
            arguments     => {
                A1 => $genCreditRound,
                A6 => $activeUnits,
            }
        ),

    );

    my $rev1d = Stack( sources => [$previousChargeImport] );

    my $rev1g = Stack( sources => [$previousChargeExport] );

    my $rev2d = Arithmetic(
        name          => 'Total for demand (£/year)',
        defaultFormat => $format0withLine,
        arithmetic    => '=' . join( '+', map { "A$_" } 1 .. @revenueBitsD ),
        arguments =>
          { map { ( "A$_" => $revenueBitsD[ $_ - 1 ] ) } 1 .. @revenueBitsD },
    );

    my $rev2g = Arithmetic(
        name          => 'Total for generation (£/year)',
        defaultFormat => $format0withLine,
        arithmetic    => '=' . join( '+', map { "A$_" } 1 .. @revenueBitsG ),
        arguments =>
          { map { ( "A$_" => $revenueBitsG[ $_ - 1 ] ) } 1 .. @revenueBitsG },
    );

    my $change1d = Arithmetic(
        name          => 'Change (demand) (£/year)',
        arithmetic    => '=A1-A4',
        defaultFormat => '0softpm',
        arguments     => { A1 => $rev2d, A4 => $rev1d }
    );

    my $change1g = Arithmetic(
        name          => 'Change (generation) (£/year)',
        arithmetic    => '=A1-A4',
        defaultFormat => '0softpm',
        arguments     => { A1 => $rev2g, A4 => $rev1g }
    );

    my $change2d = Arithmetic(
        name          => 'Change (demand) (%)',
        arithmetic    => '=IF(A1,A3/A4-1,"")',
        defaultFormat => '%softpm',
        arguments     => { A1 => $rev1d, A3 => $rev2d, A4 => $rev1d }
    );

    my $change2g = Arithmetic(
        name          => 'Change (generation) (%)',
        arithmetic    => '=IF(A1,A3/A4-1,"")',
        defaultFormat => '%softpm',
        arguments     => { A1 => $rev1g, A3 => $rev2g, A4 => $rev1g }
    );

    my $soleUseAssetChargeUnround = Arithmetic(
        name          => 'Fixed charge for demand (unrounded) (£/year)',
        defaultFormat => $format0withLine,
        arithmetic    => '=0.01*(A9-A7)*A1',
        arguments     => {
            A1 => $fixedDchargeTrue,
            A9 => $daysInYear,
            A7 => $tariffDaysInYearNot,
        }
    );

    $model->{summaryInformationColumns}[0] = $soleUseAssetChargeUnround;

    0 and my $purpleUnits = Arithmetic(
        name          => "$model->{TimebandName} units (kWh)",
        defaultFormat => '0softnz',
        arithmetic    => '=A1*(A3-A7)*A5',
        arguments     => {
            A3 => $hoursInPurple,
            A7 => $tariffHoursInPurpleNot,
            A1 => $activeCoincidence935,
            A5 => $importCapacity935,
        }
    );

    my $check = Arithmetic(
        name          => 'Check (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => join(
            '', '=A1',
            map {
                $model->{summaryInformationColumns}[$_]
                  ? ( "-A" . ( 20 + $_ ) )
                  : ()
            } 0 .. $#{ $model->{summaryInformationColumns} }
        ),
        arguments => {
            A1 => $rev2d,
            map {
                $model->{summaryInformationColumns}[$_]
                  ? ( "A" . ( 20 + $_ ),
                    $model->{summaryInformationColumns}[$_] )
                  : ()
            } 0 .. $#{ $model->{summaryInformationColumns} }
        }
    );

    if (   $model->{layout} && $model->{layout} =~ /matrix/i
        || $model->{summaries} =~ /no4601/i )
    {
        my @copyTariffs = map { Stack( sources => [$_] ) } @tariffColumns;
        SpreadsheetModel::MatrixSheet->new(
            $model->{tariff1Row}
            ? (
                dataRow            => $model->{tariff1Row},
                captionDecorations => [qw(algae purple slime)],
              )
            : (),
          )->addDatasetGroup(
            name    => 'Tariff name',
            columns => [ $copyTariffs[0] ],
          )->addDatasetGroup(
            name    => 'Import tariff',
            columns => [ @copyTariffs[ 1 .. 4 ] ],
          )->addDatasetGroup(
            name    => 'Export tariff',
            columns => [ @copyTariffs[ 5 .. 8 ] ],
          )->addDatasetGroup(
            name    => 'Import charges',
            columns => \@revenueBitsD,
          )->addDatasetGroup(
            name    => 'Export charges',
            columns => \@revenueBitsG,
          )->addDatasetGroup(
            name    => 'Change in import charges',
            columns => [ $rev2d, $rev1d, $change1d, $change2d, ],
          )->addDatasetGroup(
            name    => 'Change in export charges',
            columns => [ $rev2g, $rev1g, $change1g, $change2g, ],
          )->addDatasetGroup(
            name => 'Analysis of import charges',
            columns =>
              [ grep { $_ } @{ $model->{summaryInformationColumns} }, $check, ],
          );
        push @{ $model->{revenueTables} }, @copyTariffs, @revenueBitsD,
          @revenueBitsG, $change2d, $change2g, $check;
    }
    elsif ( $model->{vertical} ) {
        push @{ $model->{revenueTables} },
          Columnset(
            name    => 'Import charges',
            columns => [
                Stack( sources => [ $tariffColumns[0] ] ),
                @revenueBitsD, $rev2d, $rev1d, $change1d, $change2d,
            ],
          ),
          Columnset(
            name    => 'Export charges',
            columns => [
                Stack( sources => [ $tariffColumns[0] ] ),
                @revenueBitsG, $rev2g, $rev1g, $change1g, $change2g,
            ],
          ),
          Columnset(
            name    => 'Import charges based on sole-use assets',
            columns => [
                Stack( sources => [ $tariffColumns[0] ] ),
                grep { $_ } $model->{summaryInformationColumns}[0],
            ],
          ),
          Columnset(
            name    => 'Import charges based on non-sole-use notional assets',
            columns => [
                Stack( sources => [ $tariffColumns[0] ] ),
                grep { $_ } @{ $model->{summaryInformationColumns} }[ 2, 4, 7 ],
            ],
          ),
          Columnset(
            name    => 'Import charges based on capacity and consumption',
            columns => [
                Stack( sources => [ $tariffColumns[0] ] ),
                grep { $_ } @{ $model->{summaryInformationColumns} }[ 1, 3, 6 ],
            ],
          ),
          Columnset(
            name    => 'Other elements of import charges',
            columns => [
                Stack( sources => [ $tariffColumns[0] ] ),
                grep { $_ } $model->{summaryInformationColumns}[5],
                $check,
            ],
          );
    }
    else {
        push @{ $model->{revenueTables} },
          Columnset(
            name    => 'Horizontal information',
            columns => [
                ( map { Stack( sources => [$_] ) } @tariffColumns ),
                @revenueBitsD,
                @revenueBitsG,
                $rev2d,
                $rev1d,
                $change1d,
                $change2d,
                $rev2g,
                $rev1g,
                $change1g,
                $change2g,
                grep { $_ } @{ $model->{summaryInformationColumns} },
                $check,
            ],
          );
    }

    my $totalForDemandAllTariffs = GroupBy(
        source        => $rev2d,
        name          => 'Total for demand across all tariffs (£/year)',
        defaultFormat => '0softnz'
    );

    my $totalForGenerationAllTariffs = GroupBy(
        source        => $rev2g,
        name          => 'Total for generation across all tariffs (£/year)',
        defaultFormat => '0softnz'
    );

    if (   $model->{ldnoRevTotal}
        || $model->{summaries} && $model->{summaries} =~ /total/i )
    {

        push @{ $model->{ $model->{layout} ? 'TotalsTables' : 'revenueTables' }
          },
          my $totalAllTariffs = Columnset(
            name    => 'Total for all tariffs (£/year)',
            columns => [
                $model->{layout} ? () : Constant(
                    name          => 'This column is not used',
                    defaultFormat => '0con',
                    data          => [ [''] ]
                ),
                $totalForDemandAllTariffs,
                $totalForGenerationAllTariffs,
                Arithmetic(
                    name          => 'Total for all tariffs (£/year)',
                    defaultFormat => '0softnz',
                    arithmetic    => '=A1+A2',
                    arguments     => {
                        A1 => $totalForDemandAllTariffs,
                        A2 => $totalForGenerationAllTariffs,
                    }
                )
            ]
          );

        push @{ $model->{TotalsTables} },
          Columnset(
            name    => 'Total EDCM revenue (£/year)',
            columns => [
                Arithmetic(
                    name => 'All EDCM tariffs '
                      . 'including discounted CDCM tariffs (£/year)',
                    defaultFormat => '0softnz',
                    arithmetic    => '=A1+A2+A3',
                    arguments     => {
                        A1 => $totalForDemandAllTariffs,
                        A2 => $totalForGenerationAllTariffs,
                        A3 => $model->{ldnoRevTotal},
                    }
                )
            ]
          ) if $model->{ldnoRevTotal};

    }

    push @{ $model->{revenueTables} },
      $model->impactFinancialSummary( $tariffs, \@tariffColumns,
        $actualRedDemandRate, \@revenueBitsD, @revenueBitsG, $rev2g )
      if $model->{transparencyImpact};

}

1;
