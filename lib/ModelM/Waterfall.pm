package ModelM::MultiModel;

=head Copyright licence and disclaimer

Copyright 2017 Franck Latrémolière, Reckon LLP and others.

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
use SpreadsheetModel::Chart;
use SpreadsheetModel::Custom;

sub waterfallCharts {

    my ( $me, $titlePrefix, @datasets, ) = @_;
    my ( @tables, @charts );

    foreach my $d (@datasets) {

        my $boundaryPrefix =
          $d->objectShortName =~ /Boundary (\S+)/i ? "LDNO $1: " : '';

        for my $col ( 0 .. $d->lastCol ) {

            my $columnName = $d->{cols}{list}[$col];
            next if $columnName =~ /no discount/i;
            my $itemName = $boundaryPrefix . $columnName;

            my $value = SpreadsheetModel::Custom->new(
                name          => 'Baseline value',
                defaultFormat => '%soft',
                rows          => $d->{rows},
                custom        => [ '=A1', '=MIN(A1,A2)' ]
                ,    # =NA() as a formula does not work
                arithmetic => '=A3 or N/A or MIN(previous A1, A2)',
                arguments  => { A1 => $d, A2 => $d, A3 => $d, },
                wsPrepare  => sub {
                    my ( $self, $wb, $ws, $format, $formula, $pha, $rowh,
                        $colh ) = @_;
                    sub {
                        my ( $x, $y ) = @_;
                        return '', $format, $formula->[0],
                          qr/\bA1\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A1}, $colh->{A1} + $col )
                          unless $y;
                        return '', $format, $formula->[1],
                          qr/\bA1\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A1} + $y - 1,
                            $colh->{A1} + $col
                          ),
                          qr/\bA2\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A2} + $y,
                            $colh->{A2} + $col )
                          if $y == $#{ $d->{rows}{list} };
                        '=NA()', $format;
                    };
                },
            );

            my $padding = SpreadsheetModel::Custom->new(
                name          => 'Padding',
                defaultFormat => '%soft',
                rows          => $d->{rows},
                custom        => [ '=MIN(A1,A2)', ],
                arithmetic    => '=MIN(previous A1,A2) or N/A',
                arguments     => { A1 => $d, A2 => $d, },
                wsPrepare     => sub {
                    my ( $self, $wb, $ws, $format, $formula, $pha, $rowh,
                        $colh ) = @_;
                    sub {
                        my ( $x, $y ) = @_;
                        return '=NA()', $format
                          if !$y || $y == $#{ $d->{rows}{list} };
                        '', $format, $formula->[0],
                          qr/\bA1\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A1} + $y - 1,
                            $colh->{A1} + $col
                          ),
                          qr/\bA2\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A2} + $y,
                            $colh->{A2} + $col );
                    };
                },
            );

            my $increase = SpreadsheetModel::Custom->new(
                name          => 'Increase',
                defaultFormat => '%soft',
                rows          => $d->{rows},
                custom        => [ '=MAX(0,A2-A1)', ],
                arithmetic    => '=MAX(0,A2-previous A1)',
                arguments     => { A1 => $d, A2 => $d, },
                wsPrepare     => sub {
                    my ( $self, $wb, $ws, $format, $formula, $pha, $rowh,
                        $colh ) = @_;
                    sub {
                        my ( $x, $y ) = @_;
                        return '=NA()', $format unless $y;
                        '', $format, $formula->[0],
                          qr/\bA1\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A1} + $y - 1,
                            $colh->{A1} + $col
                          ),
                          qr/\bA2\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A2} + $y,
                            $colh->{A2} + $col );
                    };
                },
            );

            my $decrease = SpreadsheetModel::Custom->new(
                name          => 'Decrease',
                defaultFormat => '%soft',
                rows          => $d->{rows},
                custom        => [ '=MAX(0,A1-A2)', ],
                arithmetic    => '=MAX(0,previous A1-A2)',
                arguments     => { A1 => $d, A2 => $d, },
                wsPrepare     => sub {
                    my ( $self, $wb, $ws, $format, $formula, $pha, $rowh,
                        $colh ) = @_;
                    sub {
                        my ( $x, $y ) = @_;
                        return '=NA()', $format unless $y;
                        '', $format, $formula->[0],
                          qr/\bA1\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A1} + $y - 1,
                            $colh->{A1} + $col
                          ),
                          qr/\bA2\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{A2} + $y,
                            $colh->{A2} + $col );
                    };
                },
            );

            push @tables,
              Columnset(
                name    => "Waterfall calculations for $itemName",
                columns => [ $value, $padding, $increase, $decrease, ],
              );

            push @charts,
              SpreadsheetModel::Chart->new(
                name => $me->{waterfalls} =~ /standalone/i
                ? 'Chart ' . ( 1 + @charts )
                : $titlePrefix . $itemName,
                type         => 'bar',
                subtype      => 'stacked',
                width        => 1200,
                instructions => [
                    $me->{waterfalls} =~ /standalone/i
                    ? ( set_title => [ name => $titlePrefix . $itemName, ] )
                    : (),
                    add_series => [
                        $value,
                        overlap  => 100,
                        gap      => 8,
                        gradient => {
                            colors => [ '#FFFFFF', '#999999' ],
                            angle  => 0,
                        }
                    ],
                    add_series => [ $padding, fill => { none => 1 }, ],
                    add_series => [
                        $increase,
                        gradient => {
                            colors => [ '#FFFFFF', '#0066CC' ],
                            angle  => 0,
                        },
                    ],
                    add_series => [
                        $decrease,
                        gradient => {
                            colors => [ '#FF6633', '#FFFFFF' ],
                            angle  => 0,
                        },
                    ],
                    set_x_axis => [
                        num_format => '0%',
                        num_font   => { size => 16 },
                        min        => 0,
                        max        => 1,
                        major_unit => .1,
                    ],
                    set_y_axis => [
                        reverse  => 1,
                        num_font => { size => 16 },
                    ],
                    set_legend => [ position => 'none' ],
                ],
              );
        }

    }

    \@tables, \@charts;

}

1;
