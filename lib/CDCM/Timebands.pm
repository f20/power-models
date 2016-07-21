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
use SpreadsheetModel::ColumnsetFilter;
use Spreadsheet::WriteExcel::Utility;

sub timebandDetails {
    my ($model) = @_;
    my $rows = Labelset(
        list          => [ 1 .. 24 ],
        defaultFormat => 'thitem',
    );
    my @columns =
      map {
        my ( $name, $format ) = split /\t/;
        Dataset(
            name          => $name,
            defaultFormat => $format,
            rows          => $rows,
            /Name/
            ? (
                validation => {
                    validate => 'list',
                    value    => [qw(Red Amber Green Black Yellow Super-red)],
                }
              )
            : /Weekdays/ ? (
                validation => {
                    validate => 'list',
                    value    => [qw(A MF SS)],
                }
              )
            : /month/ ? (
                validation => {
                    validate => 'date',
                    criteria => '>',
                    value    => 0,
                }
              )
            : /time/ ? (
                validation => {
                    validate => 'time',
                    criteria => 'between',
                    minimum  => 0,
                    maximum  => 1,
                }
              )
            : (),
        );
      } split /\n/, <<EOL;
Name of time band	texthardnowrap
Weekdays\r(A, MF or SS)	puretexthard
From month/day	monthdayhard
To month/day	monthdayhard
From time	timehard
To time	timehard
EOL
    push @{ $model->{informationTables} },
      SpreadsheetModel::ColumnsetFilter->new(
        name    => 'Specification of time bands',
        number  => 1210,
        columns => \@columns,
        dataset => $model->{dataset},
      ),
      SpreadsheetModel::ColumnsetFilter->new(
        name     => 'Time band description',
        rows     => $rows,
        noFilter => 1,
        columns  => [
            Arithmetic(
                name => 'Machine-readable form',
                arithmetic =>
                  '=IF(A1="","",A91&"|"&A2&"|"&IF(A4,TEXT(A94,"d|m"),"|")&'
                  . '"|"&TEXT(A5,"[hh]:mm:ss")&"|"&IF(A3,TEXT(A93,"d|m"),"|")&"|"&TEXT(A6,"[hh]:mm:ss"))',
                defaultFormat => 'code',
                arguments     => {
                    map {
                        (
                            "A$_"  => $columns[ $_ - 1 ],
                            "A9$_" => $columns[ $_ - 1 ]
                        );
                    } 1 .. 6
                },
            ),
            Arithmetic(
                name          => 'Human-readable form',
                defaultFormat => 'colnoteleft',
                arithmetic => '=IF(A1>"",A91&" time band applies "&IF(A2="MF",'
                  . '"Monday to Friday, ",'
                  . 'IF(A92="SS","Saturdays and Sundays, ",""))&'
                  . 'IF(A3,TEXT(A93,"mmmm d")&" to "&TEXT(A4,"mmmm d")&'
                  . '", ","")&IF(A5,TEXT(A95,"[hh]:mm")&" to "&TEXT(A6,"[hh]:mm"),""),"")',
                arguments => {
                    map {
                        (
                            "A$_"  => $columns[ $_ - 1 ],
                            "A9$_" => $columns[ $_ - 1 ]
                        );
                    } 1 .. 6
                },
            ),
            map {
                Constant(
                    name          => '',
                    rows          => $rows,
                    defaultFormat => 'colnoteleft',
                    data          => [ map { '' } @{ $rows->{list} } ]
                );
            } 2 .. 5
        ],
        postWriteCalls => {
            obj => [
                sub {
                    my ( $self, $wb, $ws ) = @_;
                    my ( $ws1, $ro1, $co1 ) = $columns[0]->wsWrite( $wb, $ws );
                    my ( $ws2, $ro2, $co2 ) =
                      $self->{columns}[1]->wsWrite( $wb, $ws );
                    my @range =
                      ( $ro2, $co2, $ro2 + $#{ $rows->{list} }, $co2 + 4 );
                    my $start = '='
                      . ( $ws1 == $ws2 ? '' : "'" . $ws1->name . "'!" )
                      . xl_rowcol_to_cell( $ro1, $co1, 0, 1 );
                    $ws2->conditional_formatting(
                        @range,
                        {
                            type     => 'formula',
                            criteria => $start . '="Red"',
                            format   => $wb->getFormat(
                                [ bg_color => 10, color => 9, bold => 1 ]
                            ),
                        }
                    );
                    $ws2->conditional_formatting(
                        @range,
                        {
                            type     => 'formula',
                            criteria => $start . '="Amber"',
                            format   => $wb->getFormat(
                                [ bg_color => 52, color => 8, bold => 1 ]
                            ),
                        }
                    );
                    $ws2->conditional_formatting(
                        @range,
                        {
                            type     => 'formula',
                            criteria => $start . '="Green"',
                            format   => $wb->getFormat(
                                [ bg_color => 17, color => 9, bold => 1 ]
                            ),
                        }
                    );
                    $ws2->conditional_formatting(
                        @range,
                        {
                            type     => 'formula',
                            criteria => $start . '="Black"',
                            format   => $wb->getFormat(
                                [ bg_color => 8, color => 9, bold => 1 ]
                            ),
                        }
                    );
                    $ws2->conditional_formatting(
                        @range,
                        {
                            type     => 'formula',
                            criteria => $start . '="Yellow"',
                            format   => $wb->getFormat(
                                [ bg_color => 43, color => 8, bold => 1 ]
                            ),
                        }
                    );
                    $ws2->conditional_formatting(
                        @range,
                        {
                            type     => 'formula',
                            criteria => $start . '="Super-red"',
                            format   => $wb->getFormat(
                                [
                                    bg_color => 10,
                                    fg_color => 8,
                                    pattern  => 9,
                                    color    => 9,
                                    bold     => 1
                                ]
                            ),
                        }
                    );
                },
            ],
        },
      );
}

1;
