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

sub timebandDetails {
    my ($model) = @_;
    my $rows = Labelset(
        list          => [ 1 .. 24 ],
        defaultFormat => 'thitem',
    );
    push @{ $model->{informationTables} },
      Notes(
        name       => '',
        rowFormats => ['caption'],
        lines      => ['Specification of time bands'],
      ),
      SpreadsheetModel::ColumnsetFilter->new(
        name    => '',
        columns => [
            map {
                my ( $name, $format ) = split /\t/;
                Dataset(
                    name          => $name,
                    defaultFormat => $format,
                    rows          => $rows,
                    /name/
                    ? (
                        validation => {
                            validate => 'list',
                            value =>
                              [qw(Red Amber Green Black Yellow Super-red)],
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
              } split /\n/,
            <<EOL
Time band name	texthardnowrap
Weekdays\r(A, MF or SS)	puretexthard
From month/day	monthdayhard
To month/day	monthdayhard
From time	timehard
To time	timehard
EOL
        ],
        dataset => $model->{dataset},
      );
}

1;
