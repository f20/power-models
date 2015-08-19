package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others.

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

sub drmSizer {

    my (
        $model,              $modelLife,     $annuityRate,
        $powerFactorInModel, $coreLevels,    $coreExitLevels,
        $assetDrmLevels,     $drmExitLevels, $lineLossFactorsNetwork,
        $rerouteingMatrix
    ) = @_;

    my $exitLaf = Arithmetic(
        name => 'Loss adjustment factor'
          . ' to transmission'
          . ' for network level exit',
        rows       => $coreExitLevels,
        cols       => 0,
        arithmetic => '=A1',
        arguments  => { A1 => $lineLossFactorsNetwork }
    );

    my $entryLaf = new SpreadsheetModel::Custom(
        name => 'Loss adjustment factor'
          . ' to transmission'
          . ' for network level entry',
        rows       => $coreExitLevels,
        custom     => ['=A1'],
        arithmetic => '= A1',
        objectType => 'Special copy',
        arguments  => { A1 => $exitLaf },
        wsPrepare  => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            my $unavailable = $wb->getFormat('unavailable');
            sub {
                my ( $x, $y ) = @_;
                return '', $unavailable unless $y;
                '', $format, $formula->[0], map {
                    qr/\b$_\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + $y - 1,
                        $colh->{$_} + $x,
                        0, 0
                      )
                } @$pha;
            };
        }
    );

    my $diversityInLevel = Dataset(
        name =>
          Label( 'Diversity allowance between top and bottom of network level',
          ),
        validation => {
            validate      => 'decimal',
            criteria      => 'between',
            minimum       => 0,
            maximum       => 4,
            error_message => 'Must be' . ' a non-negative percentage value.',
        },
        lines => <<'EOL',
Source: operational data analysis and/or network model.
The diversity figure against GSP is the diversity between GSP Group (the whole system) and individual GSPs.
The diversity figure against 132kV is the diversity between GSPs (the top of the 132kV network) and 132kV/EHV bulk supply points (the bottom of the 132kV network). 
The diversity figure against EHV is the diversity between 132kV/EHV bulk supply points (the top of the EHV network) and EHV/HV primary substations (the bottom of the EHV network). 
The diversity figure against HV is the diversity between EHV/HV primary substations (the top of the HV network) and HV/LV substations (the bottom of the HV network). 
EOL
        number        => 1017,
        appendTo      => $model->{inputTables},
        dataset       => $model->{dataset},
        defaultFormat => '%hard',
        rows          => $coreExitLevels,
        data          => [ 0.05, 0.0, undef, .13, undef, .25 ]
    );

    my $exitGspCoincidence = new SpreadsheetModel::Custom(
        name          => 'Coincidence to GSP peak at level exit',
        defaultFormat => '%softnz',
        rows          => $coreExitLevels,
        custom        => [ '=1/(1+A2)', '=A1/(1+A2)' ],
        arithmetic    => '=previous/(1+A2)',
        arguments     => { A2 => $diversityInLevel },
        wsPrepare     => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            my $unavailable = $wb->getFormat('unavailable');
            push @$pha, 'A1';
            sub {
                my ( $x, $y ) = @_;
                return '', $unavailable unless $y;
                $rowh->{A1} ||= $self->{$wb}{row} - 1;
                $colh->{A1} ||= $self->{$wb}{col};
                '', $format, $formula->[ $y > 1 ? 1 : 0 ], map {
                    qr/\b$_\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + $y,
                        $colh->{$_} + $x,
                        0, 0
                      )
                } @$pha;
            };
        }
    );

    my $exitCoincidence = new SpreadsheetModel::Custom(
        name          => 'Coincidence to system peak at level exit',
        defaultFormat => '%softnz',
        rows          => $coreExitLevels,
        custom        => [ '=1/(1+A2)', '=A1/(1+A2)' ],
        arithmetic    => '=previous/(1+A2)',
        arguments     => { A2 => $diversityInLevel },
        wsPrepare     => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;

            # $self->{arguments}{A1} = $self;
            push @$pha, 'A1';
            sub {
                my ( $x, $y ) = @_;
                $rowh->{A1} ||= $self->{$wb}{row} - 1;
                $colh->{A1} ||= $self->{$wb}{col};
                '', $format, $formula->[ $y ? 1 : 0 ], map {
                    qr/\b$_\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + $y,
                        $colh->{$_} + $x,
                        0, 0
                      )
                } @$pha;
            };
        }
    );

    my $diversityAllowances = new SpreadsheetModel::Custom(
        name          => 'Diversity allowance between level exit and GSP Group',
        defaultFormat => '%softnz',
        rows          => $coreExitLevels,
        custom        => ['=1/A1-1'],
        arithmetic    => '=1/A1-1',
        arguments     => { A1 => $exitCoincidence },
        wsPrepare     => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            my $unavailable = $wb->getFormat('unavailable');
            my $last        = $#{ $self->{rows}{list} };
            sub {
                my ( $x, $y ) = @_;
                return '', $unavailable if $y == $last;
                '', $format, $formula->[0], map {
                    qr/\b$_\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + $y,
                        $colh->{$_} + $x,
                        0, 0
                      )
                } @$pha;
            };
        }
    );

=head Development note

primaryLaf assumes sizing at the substation below,
but it might be more sensible to size HV networks by primaries.

=cut

    my $drmOptions = $model->{drm} || '';

    my $primaryLaf =
      $drmOptions =~ /top/i
      ? undef
      : new SpreadsheetModel::Custom(
        name => 'Loss adjustment factor'
          . ' to transmission'
          . ' for relevant substation primary',
        defaultFormat => '0.000copy',
        rows          => $coreExitLevels,
        custom        => [ '=A1', '=A2' ],
        arithmetic    => '= A1 or A2',
        objectType    => 'Special copy',
        arguments     => { A1 => $exitLaf, A2 => $entryLaf },
        wsPrepare     => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            my $unavailable = $wb->getFormat('unavailable');
            sub {
                my ( $x, $y ) = @_;
                return '', $unavailable unless $y;
                my $lev = $coreExitLevels->{list}[$y];
                '', $format, $formula->[ $lev =~ m#^LV|/#i ? 1 : 0 ], map {
                    qr/\b$_\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + $y - ( $lev =~ m#^LV#i ? 1 : 0 ),
                        $colh->{$_} + $x,
                        0, 0
                      )
                } @$pha;
            };
        }
      );

    push @{ $model->{networkModel} },
      Columnset(
        name => 'Loss adjustment factors',
        columns =>
          [ $exitLaf, $entryLaf, $drmOptions =~ /top/i ? () : $primaryLaf, ]
      ),
      Columnset(
        name => 'Diversity calculations',
        1 ? () : ( number => 1017 ),
        columns => [
            1 ? () : $diversityInLevel, $exitGspCoincidence,
            $exitCoincidence, $diversityAllowances,
        ]
      );

    my (
        $elementUnscaledCount, $elementNameplate, $elementUtilisation,
        $elementUnitCost,      $unscaledMw
    );

    unless ( $drmOptions =~ /ext|top/i ) {

        my $networkElements = Labelset(
            name   => 'Network elements',
            groups => [
                Labelset(
                    name => '132kV',
                    list => [
                        '132kV double circuit (to supply urban BSP)',
                        '132kV double circuit (to supply rural BSP)',
                    ]
                ),
                Labelset(
                    name => '132kV/EHV',
                    list => [
                        '132kV/33kV bulk supply point (urban)',
                        '132kV/33kV bulk supply point (rural)',
                    ]
                ),
                Labelset(
                    name => 'EHV',
                    list => [
                        '33kV double circuit (urban A)',
                        '33kV double circuit (urban B)',
                        '33kV double circuit (rural C)',
                        '33kV double circuit (rural D)'
                    ]
                ),
                Labelset(
                    name => 'EHV/HV',
                    list => [
                        '33kV/11kV primary substation (urban A)',
                        '33kV/11kV primary substation (urban B)',
                        '33kV/11kV primary substation (rural C)',
                        '33kV/11kV primary substation (rural D)',
                    ]
                ),
                Labelset(
                    name => 'HV',
                    list => [
                        '11kV circuit set (urban)',
                        '11kV circuit set (suburban)',
                        '11kV circuit set (rural)',
                        '11kV circuit set (remote)',
                    ]
                ),
                Labelset(
                    name => 'HV/LV',
                    list => [
                        '11kV/415V distribution substation (urban)',
                        '11kV/415V distribution substation (suburban)',
                        '11kV/415V distribution substation (rural)',
                        '11kV/415V distribution substation (remote)',
                    ]
                ),
                Labelset(
                    name => 'LV circuits',
                    list => [
                        'LV circuit set (urban)',
                        'LV circuit set (suburban)',
                        'LV circuit set (rural)',
                        'LV circuit set (remote)',
                    ]
                )
            ]
        );

        $assetDrmLevels = Labelset(
            name    => 'Network levels',
            list    => $networkElements->{groups},
            accepts => [$assetDrmLevels]
        );

        $elementNameplate = Dataset(
            name       => 'Aggregate nameplate capacity (MVA)',
            validation => {
                validate => 'decimal',
                criteria => 'between',
                minimum  => 0,
                maximum  => 999_999.999,
            },
            rows => $networkElements,
            data => [
                undef, 240, 160,   undef, 240, 160,   undef, 80,
                60,    48,  40,    undef, 80,  60,    48,    40,
                undef, 1,   .7,    .5,    .2,  undef, 1,     .7,
                .5,    .2,  undef, 1,     .7,  .5,    .2,
            ]
        );

        $elementUtilisation = Dataset(
            name       => 'Utilisation factor',
            validation => {
                validate => 'decimal',
                criteria => 'between',
                minimum  => 0,
                maximum  => 1,
            },
            rows          => $networkElements,
            defaultFormat => '%hard',
            data          => [
                undef, .5, .5,    undef, .5, .5,    undef, .4,
                .4,    .4, .4,    undef, .4, .4,    .4,    .4,
                undef, .7, .7,    .7,    .7, undef, .7,    .7,
                .7,    .7, undef, .7,    .7, .7,    .7,
            ]
        );

        $elementUnscaledCount = Dataset(
            name       => 'Number of elements before rescaling',
            validation => {
                validate => 'decimal',
                criteria => 'between',
                minimum  => 0,
                maximum  => 999_999.999,
            },
            rows          => $networkElements,
            defaultFormat => '0hardnz',
            data          => [
                undef, 3,   2,     undef, 3,   2,     undef, 10,
                6,     3,   2,     undef, 10,  6,     3,     2,
                undef, 600, 240,   72,    24,  undef, 600,   240,
                72,    24,  undef, 600,   240, 72,    24,
            ]
        );

        $elementUnitCost = Dataset(
            name       => 'Network element asset cost (£)',
            validation => {
                validate => 'decimal',
                criteria => 'between',
                minimum  => 0,
                maximum  => 999_999_999.999,
            },
            rows          => $networkElements,
            defaultFormat => '0hardnz',
            data          => [
                undef,     8_000_000, 2_000_000, undef,
                4e6,       2e6,       undef,     4e6,
                3.2e6,     2.8e6,     2.5e6,     undef,
                1_500_000, 1_000_000, 800_000,   700_000,
                undef,     75_000,    50_000,    30_000,
                20_000,    undef,     30_000,    20_000,
                15_000,    7_000,     undef,     140_000,
                100_000,   80_000,    40_000,
            ]
        );

        push @{ $model->{networkModel} },
          Columnset(
            name    => 'Network model elements',
            columns => [
                $elementUnscaledCount, $elementNameplate,
                $elementUtilisation,   $elementUnitCost,
            ]
          );

        $_ = Stack( sources => [$_] )
          foreach $elementUnscaledCount, $elementNameplate, $elementUtilisation,
          $elementUnitCost;

        my $elementMw = Arithmetic(
            name       => 'Maximum demand by element (MW)',
            arithmetic => '=A1*A2*A3',
            arguments  => {
                A3 => $powerFactorInModel,
                A1 => $elementNameplate,
                A2 => $elementUtilisation
            }
        );

        my $elementUnscaledMw = Arithmetic(
            name       => 'Maximum load before rescaling (MW)',
            arithmetic => '=A1*A2*A3*A4',
            arguments  => {
                A3 => $powerFactorInModel,
                A1 => $elementNameplate,
                A2 => $elementUtilisation,
                A4 => $elementUnscaledCount
            }
        );

        push @{ $model->{networkModel} },
          Columnset(
            name    => 'Network model capacities before rescaling',
            columns => [
                $elementUnscaledCount, $elementNameplate,
                $elementUtilisation,   $elementUnscaledMw
            ]
          );

        $unscaledMw = GroupBy(
            name   => 'Aggregate maximum entry capacity before rescaling (MW)',
            rows   => $assetDrmLevels,
            source => $elementUnscaledMw,
        );

    }

    my $drmSize = $drmOptions =~ /([0-9]+)/ ? $1 : 0;

    push @{ $model->{optionLines} },
      $drmSize
      ? "Network model: $drmSize MW"
      . (
          $drmOptions =~ /sub/i ? ' at each substation level'
        : $drmOptions =~ /gsp/i ? ' at time of GSP peak'
        :                         ' at time of GSP Group peak'
      )
      : 'Network model: £/kW cost';

    my $wantedMw =
      $drmOptions =~ /sub/i
      ? Dataset(
        name       => 'Network model capacity (MW)',
        validation => {
            validate => 'decimal',
            criteria => 'between',
            minimum  => 0.001,
            maximum  => 999_999.999,
        },
        defaultFormat => '0hardnz',
        number        => 1019,
        appendTo      => $model->{inputTables},
        dataset       => $model->{dataset},
        data          => [$drmSize]
      )
      : $drmOptions =~ /gsp/i ? Arithmetic(
        name => 'Network model total maximum demand' . ' at substation (MW)',
        arithmetic => '=A2/A1',
        rows       => $coreLevels,
        arguments  => {
            A2 => Dataset(
                name => $drmOptions =~ /top/i
                ? 'Network model GSP peak demand (MW)'
                : 'Network model substation demand'
                  . ' at time of GSP peak (MW)',
                validation => {
                    validate => 'decimal',
                    criteria => 'between',
                    minimum  => 0.001,
                    maximum  => 999_999.999,
                },
                defaultFormat => '0hardnz',
                number        => 1019,
                appendTo      => $model->{inputTables},
                dataset       => $model->{dataset},
                data          => [$drmSize]
            ),
            A1 => $exitGspCoincidence
        }
      )
      : Arithmetic(
        name => 'Network model total maximum demand' . ' at substation (MW)',
        arithmetic => '=A2/A1',
        rows       => $coreLevels,
        arguments  => {
            A2 => Dataset(
                name => 'Network model substation demand'
                  . ' at time of system peak (MW)',
                defaultFormat => '0hardnz',
                validation    => {
                    validate => 'decimal',
                    criteria => 'between',
                    minimum  => 0,
                    maximum  => 999_999.999,
                },
                number   => 1019,
                appendTo => $model->{inputTables},
                dataset  => $model->{dataset},
                data     => [$drmSize]
            ),
            A1 => $exitCoincidence
        }
      );

    my $sizingMw =
      $unscaledMw
      ? Arithmetic(
        name       => 'Network model aggregate capacity (MW)',
        arithmetic => '=IF(A3,A2,A1)',
        arguments  => {
            A1 => $unscaledMw,
            A3 => $wantedMw,
            A2 => $wantedMw,
        }
      )
      : $wantedMw;

    my $modelSml = Arithmetic(
        name => 'Network model contribution to system maximum load'
          . ' measured at network level exit (MW)',
        rows       => $coreLevels,
        arithmetic => $drmOptions =~ /top/i
        ? '=A1*A2/A3'
        : '=A1*A2/A3*A4',
        arguments => {
            A1 => $sizingMw,
            A2 => $exitCoincidence,
            A3 => $exitLaf,
            $drmOptions =~ /top/i ? () : ( A4 => $primaryLaf )
        }
    );

    push @{ $model->{networkModel} },
      Columnset(
        name    => 'Rescaling of network model',
        columns => [
            $unscaledMw           ? $unscaledMw : (),
            $drmOptions =~ /sys/i ? $wantedMw   : (),
            $sizingMw, $modelSml
        ]
      ) unless $drmOptions =~ /ext|top/i;

    $modelSml = SumProduct(
        name => 'GSP simultaneous maximum load'
          . ' assumed through each network level (MW)',
        matrix => $modelSml,
        vector => $rerouteingMatrix,
    ) if $rerouteingMatrix;

    my $modelGrossAssetsByLevel = Dataset(
        name =>
          Label( 'Gross assets £', 'Gross asset cost by network level (£)' ),
        validation => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
        number => $model->{drmNumber} || 1020,
        appendTo      => $model->{inputTables},
        dataset       => $model->{dataset},
        defaultFormat => '0hard',
        rows          => $assetDrmLevels->{accepts}[0] || $assetDrmLevels,
        data          => [
            0, 0 * 5e6, 125e6 + 5e6,
            25e6, $rerouteingMatrix ? 0 : (),
            175e6, 25e6, 150e6
        ]
    );

    unless ( $drmOptions =~ /ext|top/i ) {

        my $elementScaledCount = Arithmetic(
            name          => 'Rescaled element count',
            defaultFormat => '0softnz',
            arithmetic    => '=IF(A5,CEILING(A6*A2/A3,1),A1)',
            arguments     => {
                A5 => $wantedMw,
                A6 => $sizingMw,
                A1 => $elementUnscaledCount,
                A2 => $elementUnscaledCount,
                A3 => $unscaledMw
            }
        );

        my $elementScaledMw = Arithmetic(
            name       => 'Rescaled maximum demand (MW)',
            arithmetic => '=A1*A2*A3*A4',
            arguments  => {
                A3 => $powerFactorInModel,
                A1 => $elementNameplate,
                A2 => $elementUtilisation,
                A4 => $elementScaledCount
            }
        );

        my $elementScaledCost = Arithmetic(
            name          => 'Asset cost (£)',
            defaultFormat => '0softnz',
            arithmetic    => '=A1*A2',
            arguments     => {
                A1 => $elementScaledCount,
                A2 => $elementUnitCost,
            }
        );

        push @{ $model->{networkModel} },
          Columnset(
            name    => 'Rescaled model elements and costs',
            columns => [
                $elementScaledCount, $elementScaledMw,
                $elementUnitCost,    $elementScaledCost
            ]
          );

        $modelGrossAssetsByLevel = GroupBy(
            name => Label(
                'Gross assets £', 'Gross asset cost by network level (£)'
            ),
            defaultFormat => '0softnz',
            rows          => $assetDrmLevels,
            cols          => 0,
            source        => $elementScaledCost,
        );

    }

    my $modelCostToSml = Arithmetic(
        arithmetic => '=IF(A5,0.001*A1*A4/A2,0)',
        rows       => $assetDrmLevels,
        arguments  => {
            A1 => $modelGrossAssetsByLevel,
            A2 => $modelSml,
            A4 => $annuityRate,
            A5 => $modelSml,
        },
        name => Label(
            'Model £/kW SML',
            'Network model annuity by simultaneous'
              . ' maximum load for each network level (£/kW/year)'
        )
    );

    0 and push @{ $model->{networkModel} }, $modelGrossAssetsByLevel;

    0
      and push @{ $model->{networkModel} },
      Columnset(
        name => 'Network model gross asset costs (£)',
        1 ? () : ( number => 1020 ),
        columns => [
            $drmOptions =~ /ext|top/i
              && !$rerouteingMatrix ? ( $sizingMw, $modelSml ) : (),
            1 ? () : $modelGrossAssetsByLevel,
            $modelCostToSml,
        ]
      );

    1 and push @{ $model->{networkModel} }, $modelCostToSml;

    $assetDrmLevels, $modelCostToSml, $diversityAllowances,
      $modelGrossAssetsByLevel, $modelSml;

}

1;
