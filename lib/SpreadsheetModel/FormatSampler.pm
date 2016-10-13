package SpreadsheetModel::FormatSampler;

=head Copyright licence and disclaimer

Copyright 2015 Franck Latrémolière, Reckon LLP and others.

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
use JSON;
use SpreadsheetModel::Shortcuts ':all';

sub new {
    my $class = shift;
    bless [@_], $class;
}

sub wsWrite {
    my ( $sampler, $wbook, $wsheet ) = @_;
    my $includeJsonColumn = $sampler->[0];

    my $row = $wsheet->{nextFree} || -1;
    $wsheet->write_string(
        ++$row, 0,
        'Spreadsheet format sampler',
        $wbook->getFormat('notes')
    );
    my $row0      = $row += 2;
    my $thFormat  = $wbook->getFormat('th');
    my $thcFormat = $wbook->getFormat('thc');

    $wsheet->write_string( $row, 1, 'Positive', $thcFormat );
    $wsheet->write_string( $row, 2, 'Negative', $thcFormat );
    $wsheet->write_string( $row, 3, 'Zero',     $thcFormat );
    $wsheet->write_string( $row, 4, 'Text',     $thcFormat );
    $wsheet->write_string( $row, 5, 'Error',    $thcFormat );
    $wsheet->write_string( $row, 6, 'JSON',     $thcFormat )
      if $includeJsonColumn;
    ++$row;

    foreach ( sort keys %{ $wbook->{formatspec} } ) {
        my $format = $wbook->getFormat($_);
        $wsheet->write_string( $row, 0, $_, $thFormat );
        $wsheet->write( $row, 1, 42,  $format );
        $wsheet->write( $row, 2, -42, $format );
        $wsheet->write( $row, 3, 0,   $format );
        $wsheet->write_string( $row, 4, $_, $format );
        $wsheet->write( $row, 5, '=1/0', $format );
        $wsheet->write_string( $row, 6, to_json( $wbook->{formatspec}{$_} ) )
          if $includeJsonColumn;
        ++$row;
    }
    $wsheet->autofilter( $row0, 0, $row - 1, $includeJsonColumn ? 6 : 5 );
    $wsheet->{nextFree} = $row
      unless $wsheet->{nextFree} && $wsheet->{nextFree} > $row;

    my $rows = Labelset( list => [ map { "MPAN $_"; } 1 .. 3 ] );
    my $data = [ [ 4200004242423 .. 4200004242425 ] ];
    Columnset(
        name    => 'Conditional formatting for MPAN validation',
        columns => [
            Constant(
                name          => 'mpanhard',
                rows          => $rows,
                data          => $data,
                defaultFormat => 'mpanhard',
                conditionalFormatting =>
                  { type => 'MPAN', format => [ bg_color => 10 ], },
            ),
            Constant(
                name          => 'mpancopy',
                rows          => $rows,
                data          => $data,
                defaultFormat => 'mpancopy',
                conditionalFormatting =>
                  { type => 'MPAN', format => [ color => 10 ], },
            ),
        ]
    )->wsWrite( $wbook, $wsheet );

}

1;
