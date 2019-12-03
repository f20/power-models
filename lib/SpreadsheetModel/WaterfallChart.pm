package SpreadsheetModel::WaterfallChart;

# Copyright 2017-2018 Franck Latrémolière, Reckon LLP and others.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice,
# this list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
# this list of conditions and the following disclaimer in the documentation
# and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
# EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
# THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

use warnings;
use strict;
use utf8;

use base 'SpreadsheetModel::Chart';

sub check {

    my ($self) = @_;
    return 'No waterfall chart without padding data!'
      unless ref $self->{padding};

    $self->{type}    = 'bar';
    $self->{subtype} = 'stacked';
    $self->{scaling_factor} ||= 1;

    my $lastId = $self->{padding}->lastRow;
    my $lc     = $self->{padding}->lastCol;
    $lastId = $lc if $lc > $lastId;
    $self->{height} ||=
      $self->{scaling_factor} * ( $lastId > 14 ? 180 + 24 * $lastId : 480 );

    my @greyBarSeriesSettings = (
        gradient => {
            colors =>
              [ '#999999', '#999999', '#999999', '#999999', '#999999', ],
            transparency => [ 80, 80, 0,  80, 80, ],
            positions    => [ 0,  35, 50, 65, 100, ],
            angle        => 90,
        },
    );
    my @orangeColours = (
        colors       => [ '#FF6633', '#FF6633', ],
        transparency => [ 80,        0, ],
    );
    my @blueColours = (
        colors       => [ '#0066CC', '#0066CC', ],
        transparency => [ 80,        0, ],
    );
    my @overlapGap = ( overlap => 100, gap => 5, );
    push @{ $self->{instructions} },
      add_series => [ $self->{padding}, @overlapGap, fill => { none => 1, }, ],
      add_series =>
      [ $self->{blue_rightwards}, gradient => { @blueColours, angle => 0, }, ],
      add_series =>
      [ $self->{blue_leftwards}, gradient => { @blueColours, angle => 180, }, ],
      add_series => [
        $self->{orange_rightwards},
        gradient => { @orangeColours, angle => 0, },
      ],
      add_series => [
        $self->{orange_leftwards},
        gradient => { @orangeColours, angle => 180, },
      ],
      set_y_axis => [
        reverse  => 1,
        num_font => { size => $self->{scaling_factor} * 12, },
      ],
      set_legend => [ position => 'none' ],
      combine    => [
        type         => 'bar',
        instructions => [
            add_series => [
                $self->{grey_rightwards}, @overlapGap, @greyBarSeriesSettings,
            ],
            add_series =>
              [ $self->{grey_leftwards}, @overlapGap, @greyBarSeriesSettings, ],
        ],
      ];

    return $self->SUPER::check;

}

1;
