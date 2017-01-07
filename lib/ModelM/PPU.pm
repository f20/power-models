package ModelM;

=head Copyright licence and disclaimer

Copyright 2016-2017 Franck Latrémolière, Reckon LLP and others.

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

sub ppuCalcCdcm {

    my ( $model, $discounts, $ppu, $ppuNotSplit, ) = @_;

    push @{ $model->{objects}{resultsTables} }, Columnset(
        name          => 'Discount p/kWh ⇒1039. For Model G',
        singleRowName => 'LDNO discount p/kWh',
        columns       => [
            map {
                if (/No discount/) {
                    Constant( name => 'No discount', data => [ [] ], );
                }
                else {
                    my $offset = /: HV/i ? 2 : /: LV Sub/i ? 1 : 0;
                    ++$offset if $offset && $model->{dcp095};
                    push @{ $model->{objects}{table1039sources} },
                      my $col = SpreadsheetModel::Custom->new(
                        name =>
                          SpreadsheetModel::Object::_shortName( $_->{name} ),
                        cols => Labelset(
                            list => [
                                SpreadsheetModel::Object::_shortName(
                                    $_->{name}
                                )
                            ]
                        ),
                        custom     => ['=A1*SUM(A2:A3,A4)'],
                        arithmetic => 'Special calculation',
                        arguments  => {
                            A1 => $_,
                            A2 => $ppu,
                            A3 => $ppu,
                            A4 => $ppuNotSplit,
                        },
                        wsPrepare => sub {
                            my ( $self, $wb, $ws, $format, $formula, $pha,
                                $rowh, $colh )
                              = @_;
                            sub {
                                my ( $x, $y ) = @_;
                                '', $format, $formula->[0], map {
                                    qr/\b$_\b/ =>
                                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                                        $rowh->{$_},
                                        $colh->{$_} + (
                                              /A3/ ? $ppu->lastCol
                                            : /A2/ ? $offset
                                            : 0
                                        ),
                                      )
                                } @$pha;
                            };
                        },
                      );
                    $col;
                }
            } @{ $discounts->{columns} }
        ],
    );

}

sub ppuCalcEdcm {

    my ( $model, $discounts, $ppu, $ppuNotSplit, ) = @_;

    push @{ $model->{objects}{resultsTables} }, Columnset(
        name    => 'Discount p/kWh ⇒1184. For EDCM model',
        columns => [
            map {
                my $offset =
                    /^HV (?:sub|gen)/i ? 3
                  : /^HV/i             ? 2
                  : /LV (?:sub|gen)/i  ? 1
                  :                      0;
                ++$offset if $offset && $model->{dcp095};
                SpreadsheetModel::Custom->new(
                    name       => $_->{name},
                    custom     => ['=A1*SUM(A2:A3,A4)'],
                    arithmetic => 'Special calculation',
                    rows       => $_->{rows},
                    arguments  => {
                        A1 => $_,
                        A2 => $ppu,
                        A3 => $ppu,
                        A4 => $ppuNotSplit,
                    },
                    wsPrepare => sub {
                        my ( $self, $wb, $ws, $format, $formula, $pha, $rowh,
                            $colh )
                          = @_;
                        sub {
                            my ( $x, $y ) = @_;
                            '', $format, $formula->[0], map {
                                qr/\b$_\b/ =>
                                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                                    $rowh->{$_} + ( /A1/ ? $y : 0 ),
                                    $colh->{$_} + (
                                          /A3/ ? $ppu->lastCol
                                        : /A2/ ? $offset
                                        : 0
                                    ),
                                  )
                            } @$pha;
                        };
                    },
                );
            } @{ $discounts->{columns} }
        ],
    );

}

1;
