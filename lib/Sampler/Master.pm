package Sampler;

=head Copyright licence and disclaimer

Copyright 2015-2017 Franck Latrémolière, Reckon LLP and others.

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

sub new {
    my $class = shift;
    bless {@_}, $class;
}

sub worksheetsAndClosures {
    my ( $model, $wbook ) = @_;
    Sampler => sub {
        my ($wsheet) = @_;
        $wsheet->set_column( 0, 5, 20 );
        Notes( name => 'Spreadsheet format sampler' )
          ->wsWrite( $wbook, $wsheet );
        $model->writeFormatList( $wbook, $wsheet ) unless $model->{omitFormats};
        $model->writeColourCode( $wbook, $wsheet )
          if $model->{showColourCode};
        $model->writeNumFormatColours( $wbook, $wsheet )
          if $model->{showNumFormatColours};
        $model->writeMpanFormats( $wbook, $wsheet )
          unless $model->{omitMpanFormats};
    };
}

sub writeFormatList {
    my ( $model, $wbook, $wsheet ) = @_;
    my $row = $wsheet->{nextFree} || -1;
    ++$row;
    my $thFormat  = $wbook->getFormat('th');
    my $thcFormat = $wbook->getFormat('thc');
    $wsheet->write_string( $row, 1, 'Positive', $thcFormat );
    $wsheet->write_string( $row, 2, 'Negative', $thcFormat );
    $wsheet->write_string( $row, 3, 'Zero',     $thcFormat );
    $wsheet->write_string( $row, 4, 'Text',     $thcFormat );
    $wsheet->write_string( $row, 5, 'Error',    $thcFormat );
    ++$row;

    foreach ( sort keys %{ $wbook->{formatspec} } ) {
        my $format = $wbook->getFormat($_);
        $wsheet->write_string( $row, 0, $_, $thFormat );
        $wsheet->write( $row, 1, 42,  $format );
        $wsheet->write( $row, 2, -42, $format );
        $wsheet->write( $row, 3, 0,   $format );
        $wsheet->write_string( $row, 4, $_, $format );
        $wsheet->write( $row, 5, '=1/0', $format );
        ++$row;
    }
    $wsheet->{nextFree} = $row;
}

sub writeMpanFormats {
    my ( $model, $wbook, $wsheet ) = @_;
    my $rows = Labelset(
        defaultFormat => [ base => 'th', num_format => '\M\P\A\N 00', ],
        list          => [ 1 .. 12 ]
    );
    my $data = [ [ 4200004242423 .. 4200004242434 ] ];
    my $entry = Constant(
        name          => 'Entry field',
        rows          => $rows,
        data          => $data,
        defaultFormat => 'mpanhard',
        conditionalFormatting =>
          { type => 'MPAN', format => [ bg_color => 10 ], },
    );
    Columnset(
        name    => 'Conditional formatting for MPAN validation',
        columns => [
            $entry,
            Arithmetic(
                name          => 'Copy field',
                arithmetic    => '=A1',
                arguments     => { A1 => $entry },
                defaultFormat => 'mpancopy',
                conditionalFormatting =>
                  { type => 'MPAN', format => [ color => 10 ], },
            ),
        ]
    )->wsWrite( $wbook, $wsheet );
}

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
