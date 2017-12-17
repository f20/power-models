package SpreadsheetModel::WaterfallChart;

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

use base 'SpreadsheetModel::Chart';

sub check {

    my ($self) = @_;
    return 'No padding in waterfall chart'
      unless ref $self->{padding};

    $self->{type}    = 'bar';
    $self->{subtype} = 'stacked';

    my $lastId = $self->{padding}->lastRow;
    my $lc     = $self->{padding}->lastCol;
    $lastId = $lc if $lc > $lastId;
    $self->{height} ||= 192 + 24 * $lastId;
    push @{ $self->{instructions} },
      add_series => [
        $self->{grey_rightwards},
        overlap  => 100,
        gap      => 8,
        gradient => {
            colors => [ '#CCCCCC', '#666666' ],
            angle  => 0,
        },
      ],
      add_series => [
        $self->{grey_leftwards},
        gradient => {
            colors => [ '#CCCCCC', '#666666' ],
            angle  => 180,
        },
      ],
      add_series => [ $self->{padding}, fill => { none => 1 }, ],
      add_series => [
        $self->{blue_rightwards},
        gradient => {
            colors => [ '#C0E0FF', '#0066CC' ],
            angle  => 0,
        },
      ],
      add_series => [
        $self->{blue_leftwards},
        gradient => {
            colors => [ '#C0E0FF', '#0066CC' ],
            angle  => 180,
        },
      ],
      add_series => [
        $self->{orange_rightwards},
        gradient => {
            colors => [ '#FFCFBF', '#FF6633' ],
            angle  => 0,
        },
      ],
      add_series => [
        $self->{orange_leftwards},
        gradient => {
            colors => [ '#FFCFBF', '#FF6633' ],
            angle  => 180,
        },
      ],
      set_y_axis => [
        reverse  => 1,
        num_font => { size => 16 },
      ],
      set_legend => [ position => 'none' ];

    return $self->SUPER::check;

}

1;
