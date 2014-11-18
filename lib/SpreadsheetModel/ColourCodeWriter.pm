package SpreadsheetModel::ColourCodeWriter;

=head Copyright licence and disclaimer

Copyright 2008-2014 Franck Latrémolière, Reckon LLP and others.

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

sub new {
    my $class = shift;
    bless [@_], $class;
}

sub wsWrite {
    my ( $colourCode, $wbook, $wsheet ) = @_;
    my $row = $wsheet->{nextFree} || 0;
    $row -= $colourCode->[0] ? 5 : 8;
    $row = 1 if $row < 1;
    $wsheet->write_string(
        ++$row, 2,
        'Colour coding',
        $wbook->getFormat('thc')
    );
    $wsheet->write_string( ++$row, 2, 'Input data',
        $wbook->getFormat('0.000hard') );
    $wsheet->write_string(
        ++$row, 2,
        'Constant value',
        $wbook->getFormat('0.000con')
    ) unless $colourCode->[0];
    $wsheet->write_string(
        ++$row, 2,
        'Formula: calculation',
        $wbook->getFormat('0.000soft')
    );
    $wsheet->write_string(
        ++$row, 2,
        $colourCode->[0] ? 'Data from tariff model' : 'Formula: copy',
        $wbook->getFormat('0.000copy')
    );
    $wsheet->write_string(
        ++$row, 2,
        'Unused cell in input data table',
        $wbook->getFormat('unused')
    ) unless $colourCode->[0];
    $wsheet->write_string(
        ++$row, 2,
        'Unused cell in other table',
        $wbook->getFormat('unavailable')
    ) unless $colourCode->[0];
    $wsheet->write_string(
        ++$row, 2,
        'Unlocked cell for notes',
        $wbook->getFormat('scribbles')
    ) unless $colourCode->[0];
    $wsheet->{nextFree} = $row
      unless $wsheet->{nextFree} && $wsheet->{nextFree} > $row;
}

1;