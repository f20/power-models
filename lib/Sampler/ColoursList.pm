package Sampler;

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
use Data::Dumper;
use SpreadsheetModel::Shortcuts ':all';

sub writeColourCode {

    my ( $model, $wbook, $wsheet ) = @_;

    Notes( name => 'Colour coding and options' )->wsWrite( $wbook, $wsheet );
    my $row = $wsheet->{nextFree} || -1;

    $wsheet->write_string( ++$row, 0, 'Header', $wbook->getFormat('th') );
    $wsheet->write_string( $row, 1, '#eeddff',
        $wbook->getFormat( [ base => 'puretexthard', bg_color => '#eeddff' ] )
    );
    $wsheet->write_string( $row, 2, '#cc99ff',
        $wbook->getFormat( [ base => 'puretexthard', bg_color => '#cc99ff' ] )
    );
    $wsheet->write_string( $row, 3, '#ffcc99',
        $wbook->getFormat( [ base => 'puretexthard', bg_color => '#ffcc99' ] )
    );
    $wsheet->write_string( $row, 4, '#ffd700',
        $wbook->getFormat( [ base => 'puretexthard', bg_color => '#ffd700' ] )
    );

    $wsheet->write_string( ++$row, 0, 'Input data',
        $wbook->getFormat('texthard') );
    $wsheet->write_string( $row, 1, '#ccffff',
        $wbook->getFormat( [ base => 'puretexthard', bg_color => '#ccffff' ] )
    );

    $wsheet->write_string(
        ++$row, 0,
        'Constant value',
        $wbook->getFormat('textcon')
    );
    $wsheet->write_string( $row, 1, '#e9e9e9',
        $wbook->getFormat( [ base => 'puretexthard', bg_color => '#e9e9e9' ] )
    );
    $wsheet->write_string( $row, 2, '#c0c0c0',
        $wbook->getFormat( [ base => 'puretexthard', bg_color => '#c0c0c0' ] )
    );

    $wsheet->write_string(
        ++$row, 0,
        'Formula: calculation',
        $wbook->getFormat('textsoft')
    );
    $wsheet->write_string( $row, 1, '#ffffcc',
        $wbook->getFormat( [ base => 'puretexthard', bg_color => '#ffffcc' ] )
    );
    $wsheet->write_string( $row, 2, '#ffff99',
        $wbook->getFormat( [ base => 'puretexthard', bg_color => '#ffff99' ] )
    );

    $wsheet->write_string(
        ++$row, 0,
        'Formula: copy',
        $wbook->getFormat('textcopy')
    );
    $wsheet->write_string( $row, 1, '#ccffcc',
        $wbook->getFormat( [ base => 'puretexthard', bg_color => '#ccffcc' ] )
    );

    $wsheet->write_string(
        ++$row, 0,
        'Unlocked cell for notes',
        $wbook->getFormat('scribbles')
    );
    $wsheet->write_string( $row, 1, '#fbf8ff',
        $wbook->getFormat( [ base => 'puretexthard', bg_color => '#fbf8ff' ] )
    );
    $wsheet->write_string(
        $row, 2,
        '#800080',
        $wbook->getFormat(
            [
                base     => 'puretexthard',
                color    => '#99b3cc',
                bg_color => '#800080'
            ]
        )
    );
    $wsheet->write_string( $row, 3, '#99ccff',
        $wbook->getFormat( [ base => 'puretexthard', bg_color => '#99ccff' ] )
    );

    $wsheet->{nextFree} = $row + 1;

}

sub writeNumFormatColours {
    my ( $model, $wbook, $wsheet ) = @_;
    Notes( name => 'Excel number format colours' )->wsWrite( $wbook, $wsheet );
    my $row = $wsheet->{nextFree} || -1;
    foreach (
        [ '[red]',     '#dd0806' ],
        [ '[cyan]',    '#00abea' ],
        [ '[blue]',    '#0000d4' ],
        [ '[green]',   '#17b715' ],
        [ '[magenta]', '#f20784' ],
        [ '[yellow]',  '#fcf304' ],
        [ '',          '#ff00ff' ]
      )
    {
        $wsheet->write_string(
            ++$row,
            0,
            'invisible text',
            $wbook->getFormat(
                [
                    base       => 'texthard',
                    bg_color   => $_->[1],
                    num_format => "$_->[0]\@"
                ]
            )
        );
        $wsheet->write_string( $row, 1, $_->[0],
            $wbook->getFormat('puretextcon') );
        $wsheet->write_string( $row, 2, $_->[1],
            $wbook->getFormat('puretextcon') );
    }
    $wsheet->{nextFree} = $row + 1;
}

1;
