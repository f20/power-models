package ModelM;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others.
Copyright 2012-2013 Franck Latrémolière, Reckon LLP and others.

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

sub adjust117 {

    my ( $model, $meavPercentages, $preAllocated, ) = @_;

    my $dcp117negative = Stack(
        name => 'DCP 117: negative number being removed',
        rows => Labelset(
            list => [
                    'Load related new connections & '
                  . 'reinforcement (net of contributions)'
            ]
        ),
        cols          => Labelset( list => ['LV'] ),
        sources       => [$preAllocated],
        defaultFormat => '0copynz',
    );

    my $dcp117 = SpreadsheetModel::Custom->new(
        name          => 'DCP 117: shares of reallocation',
        cols          => $preAllocated->{cols},
        defaultFormat => '%softnz',
        custom        => [
            '=A11/SUM(A21:A22)', $model->{dcp117} =~ /A/i
            ? ()
            : '=SUM(A11:A12)/SUM(A21:A22)',
        ],
        arguments => {
            A11 => $meavPercentages,
            A21 => $meavPercentages,
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            my $unavailable = $wb->getFormat('unavailable');
            $model->{dcp117} =~ /A/i
              ? sub {
                my ( $x, $y ) = @_;
                return -1, $format if $x == 0;
                return '', $format, $formula->[0],
                  qr/\bA11\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A11}, $colh->{A11} + 1,
                    0, 1 ),
                  qr/\bA21\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A11}, $colh->{A11} + 1,
                    0, 1 ),
                  qr/\bA22\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A11}, $colh->{A11} + 3,
                    0, 1 ),
                  if $x == 1;
                return '', $format, $formula->[0],
                  qr/\bA11\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A11}, $colh->{A11} + 2,
                    0, 1 ),
                  qr/\bA21\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A11}, $colh->{A11} + 1,
                    0, 1 ),
                  qr/\bA22\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A11}, $colh->{A11} + 3,
                    0, 1 ),
                  if $x == 2;
                return '', $format, $formula->[0],
                  qr/\bA11\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A11}, $colh->{A11} + 3,
                    0, 1 ),
                  qr/\bA21\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A11}, $colh->{A11} + 1,
                    0, 1 ),
                  qr/\bA22\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A11}, $colh->{A11} + 3,
                    0, 1 ),
                  if $x == 3;
                return '', $unavailable;
              }
              : sub {
                my ( $x, $y ) = @_;
                return -1, $format if $x == 0;
                return 0,  $format if $x == 1;
                return '', $format, $formula->[1],
                  qr/\bA11\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A11}, $colh->{A11} + 1,
                    0, 1 ),
                  qr/\bA12\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A11}, $colh->{A11} + 2,
                    0, 1 ),
                  qr/\bA21\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A11}, $colh->{A11} + 1,
                    0, 1 ),
                  qr/\bA22\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A11}, $colh->{A11} + 3,
                    0, 1 ),
                  if $x == 2;
                return '', $format, $formula->[0],
                  qr/\bA11\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A11}, $colh->{A11} + 3,
                    0, 1 ),
                  qr/\bA21\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A11}, $colh->{A11} + 1,
                    0, 1 ),
                  qr/\bA22\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A11}, $colh->{A11} + 3,
                    0, 1 ),
                  if $x == 3;
                return '', $unavailable;
              };
        },
    );

    Stack(
        name    => 'Allocated costs after DCP 117 adjustments',
        rows    => $preAllocated->{rows},
        cols    => $preAllocated->{cols},
        sources => [
            Arithmetic(
                name => 'Load related new connections & '
                  . 'reinforcement (net of contributions)'
                  . ' after DCP 117',
                rows          => $dcp117negative->{rows},
                cols          => $preAllocated->{cols},
                defaultFormat => '0softnz',
                arithmetic    => '=A1+A2*A3',
                arguments     => {
                    A1 => $preAllocated,
                    A2 => $dcp117,
                    A3 => $dcp117negative,
                },
            ),
            $preAllocated
        ],
    );

}

1;
