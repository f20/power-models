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

sub ajust117 {

    my ( $model, $meavPercentages, $preAllocated, ) = @_;

    if ( $model->{dcp117} =~ /2014/ ) {

        $preAllocated = Stack(
            name => 'Table 1330 allocated costs, after DCP 117 adjustments',
            defaultFormat => '0copy',
            rows          => $preAllocated->{rows},
            cols          => $preAllocated->{cols},
            sources       => [
                Dataset(
                    name => 'Net new connections and reinforcement costs (£)',
                    rows =>
                      Labelset( list => [ $preAllocated->{rows}{list}[0] ] ),
                    cols          => $preAllocated->{cols},
                    defaultFormat => '0hard',
                    data     => [ map { '' } @{ $preAllocated->{cols}{list} } ],
                    number   => 1329,
                    dataset  => $model->{dataset},
                    appendTo => $model->{inputTables},
                ),
                $preAllocated
            ],
        );

    }
    elsif ( $model->{dcp117} =~ /half[ -]?baked/i ) {

        $preAllocated = Stack(
            name    => 'Table 1330 allocated costs, after DCP 117 adjustments',
            rows    => $preAllocated->{rows},
            cols    => $preAllocated->{cols},
            sources => [
                Constant(
                    name => 'DCP 117: remove negative number',
                    rows => Labelset(
                        list => [
                                'Load related new connections & '
                              . 'reinforcement (net of contributions)'
                        ]
                    ),
                    cols => Labelset( list => ['LV'] ),
                    data => [ [0] ],
                    defaultFormat => '0connz',
                ),
                $preAllocated
            ],
        );

    }
    else {

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
        my $dcp117 = new SpreadsheetModel::Custom(
            name          => 'DCP 117: shares of reallocation',
            cols          => $preAllocated->{cols},
            defaultFormat => '%softnz',
            custom        => [
                '=IV11/SUM(IV21:IV22)',
                $model->{dcp117} =~ /A/i
                ? ()
                : '=SUM(IV11:IV12)/SUM(IV21:IV22)',
            ],
            arguments => {
                IV11 => $meavPercentages,
                IV21 => $meavPercentages,
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                my $unavailable = $wb->getFormat('unavailable');
                $model->{dcp117} =~ /A/i
                  ? sub {
                    my ( $x, $y ) = @_;
                    return -1, $format if $x == 0;
                    return '', $format, $formula->[0],
                      IV11 =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{IV11}, $colh->{IV11} + 1,
                        0, 1 ),
                      IV21 =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{IV11}, $colh->{IV11} + 1,
                        0, 1 ),
                      IV22 =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{IV11}, $colh->{IV11} + 3,
                        0, 1 ),
                      if $x == 1;
                    return '', $format, $formula->[0],
                      IV11 =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{IV11}, $colh->{IV11} + 2,
                        0, 1 ),
                      IV21 =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{IV11}, $colh->{IV11} + 1,
                        0, 1 ),
                      IV22 =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{IV11}, $colh->{IV11} + 3,
                        0, 1 ),
                      if $x == 2;
                    return '', $format, $formula->[0],
                      IV11 =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{IV11}, $colh->{IV11} + 3,
                        0, 1 ),
                      IV21 =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{IV11}, $colh->{IV11} + 1,
                        0, 1 ),
                      IV22 =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{IV11}, $colh->{IV11} + 3,
                        0, 1 ),
                      if $x == 3;
                    return '', $unavailable;
                  }
                  : sub {
                    my ( $x, $y ) = @_;
                    return -1, $format if $x == 0;
                    return 0,  $format if $x == 1;
                    return '', $format, $formula->[1],
                      IV11 =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{IV11}, $colh->{IV11} + 1,
                        0, 1 ),
                      IV12 =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{IV11}, $colh->{IV11} + 2,
                        0, 1 ),
                      IV21 =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{IV11}, $colh->{IV11} + 1,
                        0, 1 ),
                      IV22 =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{IV11}, $colh->{IV11} + 3,
                        0, 1 ),
                      if $x == 2;
                    return '', $format, $formula->[0],
                      IV11 =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{IV11}, $colh->{IV11} + 3,
                        0, 1 ),
                      IV21 =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{IV11}, $colh->{IV11} + 1,
                        0, 1 ),
                      IV22 =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{IV11}, $colh->{IV11} + 3,
                        0, 1 ),
                      if $x == 3;
                    return '', $unavailable;
                  };
            },
        );
        $preAllocated = Stack(
            name    => 'Table 1330 allocated costs, after DCP 117 adjustments',
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
                    arithmetic    => '=IV1+IV2*IV3',
                    arguments     => {
                        IV1 => $preAllocated,
                        IV2 => $dcp117,
                        IV3 => $dcp117negative,
                    },
                ),
                $preAllocated
            ],
        );

    }

    $preAllocated;

}

1;
