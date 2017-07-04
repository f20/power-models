package CDCM;

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

# Includes rounding, finishing, tariff matrices and revenue matrices.

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Chart;
use SpreadsheetModel::Chartset;

sub colourListPastel {
    (
        '#ccffff',    # Assets 132kV
        '#ccffe6',    # Assets 132kV/EHV
        '#ccffcc',    # Assets EHV
        '#e6e6cc',    # Assets EHV/HV
        '#e6e6e6',    # Assets 132kV/HV
        '#ffcccc',    # Assets HV
        '#ffe6cc',    # Assets HV/LV
        '#ffffcc',    # Assets LV circuits
        '#ffe6e6',    # Assets LV customer
        '#ffcce6',    # Assets HV customer
        '#ccccff',    # Transmission exit
        '#ccffff',    # Operating 132kV
        '#ccffe6',    # Operating 132kV/EHV
        '#ccffcc',    # Operating EHV
        '#e6e6cc',    # Operating EHV/HV
        '#e6e6e6',    # Operating 132kV/HV
        '#ffcccc',    # Operating HV
        '#ffe6cc',    # Operating HV/LV
        '#ffffcc',    # Operating LV circuits
        '#ffe6e6',    # Operating LV customer
        '#ffcce6',    # Operating HV customer
        '#ffccff',    # Matching
        '#ffffff',    # Rounding
    );
}

sub colourListSaturated {
    (
        '#00ffff',    # Assets 132kV
        '#00ff80',    # Assets 132kV/EHV
        '#00ff00',    # Assets EHV
        '#808000',    # Assets EHV/HV
        '#808080',    # Assets 132kV/HV
        '#ff0000',    # Assets HV
        '#ff8000',    # Assets HV/LV
        '#ffff00',    # Assets LV circuits
        '#ff8080',    # Assets LV customer
        '#ff0080',    # Assets HV customer
        '#0000ff',    # Transmission exit
        '#00ffff',    # Operating 132kV
        '#00ff80',    # Operating 132kV/EHV
        '#00ff00',    # Operating EHV
        '#808000',    # Operating EHV/HV
        '#808080',    # Operating 132kV/HV
        '#ff0000',    # Operating HV
        '#ff8000',    # Operating HV/LV
        '#ffff00',    # Operating LV circuits
        '#ff8080',    # Operating LV customer
        '#ff0080',    # Operating HV customer
        '#ff00ff',    # Matching
        '#ffffff',    # Rounding
    );
}

sub colourListAdjusted {
    (
        '#009999',    # Assets 132kV
        '#00994d',    # Assets 132kV/EHV
        '#009900',    # Assets EHV
        '#808000',    # Assets EHV/HV
        '#804d4d',    # Assets 132kV/HV
        '#ff0000',    # Assets HV
        '#cc8000',    # Assets HV/LV
        '#999900',    # Assets LV circuits
        '#cc4d80',    # Assets LV customer
        '#ff0080',    # Assets HV customer
        '#0000ff',    # Transmission exit
        '#009999',    # Operating 132kV
        '#00994d',    # Operating 132kV/EHV
        '#009900',    # Operating EHV
        '#808000',    # Operating EHV/HV
        '#804d4d',    # Operating 132kV/HV
        '#ff0000',    # Operating HV
        '#cc8000',    # Operating HV/LV
        '#999900',    # Operating LV circuits
        '#cc4d80',    # Operating LV customer
        '#ff0080',    # Operating HV customer
        '#ff00ff',    # Matching
        '#ffffff',    # Rounding
    );
}

sub colourList {
    (
        '#66cccc',    # Assets 132kV
        '#66cc99',    # Assets 132kV/EHV
        '#66cc66',    # Assets EHV
        '#b3b366',    # Assets EHV/HV
        '#b3b3b3',    # Assets 132kV/HV
        '#ff6666',    # Assets HV
        '#e6b366',    # Assets HV/LV
        '#cccc66',    # Assets LV circuits
        '#e699b3',    # Assets LV customer
        '#ff66b3',    # Assets HV customer
        '#0000ff',    # Transmission exit
        '#009999',    # Operating 132kV
        '#00994d',    # Operating 132kV/EHV
        '#009900',    # Operating EHV
        '#808000',    # Operating EHV/HV
        '#804d4d',    # Operating 132kV/HV
        '#ff0000',    # Operating HV
        '#cc8000',    # Operating HV/LV
        '#999900',    # Operating LV circuits
        '#cc4d80',    # Operating LV customer
        '#ff0080',    # Operating HV customer
        '#ff00ff',    # Matching
        '#000000',    # Rounding
    );
}

sub matrixCharts {
    my ( $model, $title, @columns, ) = @_;
    my @colourList      = $model->colourList;
    my @nameFormulaCols = map {
        my $name  = $_->objectShortName;
        my $units = '';
        $units = $1 if $name =~ s/(\S+\/\S+)$//;
        my $format = $units =~ /\/day/i ? '0.00' : '0.000';
        $_->{nameFormula} = Arithmetic(
            name       => '',
            arithmetic => qq%="$name\n"&TEXT(SUM(A1_A2),"$format")&"$units"%,
            arguments  => { A1_A2 => $_ }
        );
    } @columns;
    Columnset( name => '', columns => \@nameFormulaCols, );
    SpreadsheetModel::Chartset->new(
        name => "$title: pie charts (hover over each slice to see details)",
        rows => $columns[0]{rows},
        rowFormats => [
            map {
                [
                    base         => 'th',
                    bg_color     => 1 ? 'white' : 'black',
                    num_format   => '@',
                    color        => $_,
                    border       => 7,
                    border_color => '#999999',
                ];
            } @colourList,
        ],
        charts => [
            map {
                SpreadsheetModel::Chart->new(
                    type         => 'pie',
                    instructions => [
                        set_title => [
                            name_formula => $_->{nameFormula},
                            1 ? ()
                            : (
                                name_font => {
                                    bold  => 1,
                                    color => '#99b3cc',
                                }
                            ),
                        ],
                        1 ? ()
                        : ( set_chartarea => [ fill => { color => 'black' } ] ),
                        set_legend => [ position => 'none' ],
                        add_series => [
                            $_,
                            line   => { none => 1 },
                            points => [
                                map { +{ fill => { color => $_, } }; }
                                  @colourList
                            ],
                        ],
                    ],
                  )
            } @columns,
        ],
    );
}

1;
