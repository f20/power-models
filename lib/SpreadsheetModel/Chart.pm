package SpreadsheetModel::Chart;

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
use utf8;

require SpreadsheetModel::Object;
our @ISA = qw(SpreadsheetModel::Object);
use Spreadsheet::WriteExcel::Utility;

sub objectType {
    'Chart';
}

sub check {
    my ($self) = @_;
    return
        "Broken chart $self->{name} $self->{debug}: instructions is "
      . ( ref $self->{instructions} )
      . ' but must be ARRAY'
      unless ref $self->{instructions} eq 'ARRAY';
    $self->{height} ||= 288;
    $self->{width}  ||= 480;
    return;
}

sub wsUrl {
    my ( $self, $wb ) = @_;
    return unless $self->{$wb};
    my ( $wo, $ro, $co ) = @{ $self->{$wb} }{qw(worksheet row col)};
    my $ce = xl_rowcol_to_cell( $ro, $co );
    my $wn =
        $wo
      ? $wo->get_name
      : die( join "No worksheet for $self->{name}" );
    "internal:'$wn'!$ce";
}

sub applyInstructions {
    my ( $self, $chart, $wb, $ws, $instructions ) = @_;
    my @instructions = @$instructions;
    while (@instructions) {
        my ( $verb, $args ) = splice @instructions, 0, 2;
        if ( $verb eq 'combine' ) {
            my %a = @$args;
            my $i = delete $a{instructions};
            my $c = $wb->add_chart( %a, embedded => 1 );
            $self->applyInstructions( $c, $wb, $ws, $i );
            $chart->combine($c);
            next;
        }
        if (   $verb eq 'add_series'
            && UNIVERSAL::isa( $args, 'SpreadsheetModel::Dataset' )
            && !$args->{rows}
            && $args->{cols} )
        {
            my ( $w2, $r2, $c2 ) = $args->wsWrite( $wb, $ws, undef, undef, 1 );
            $w2 = "'" . $w2->get_name . "'!";
            my $r3 = $r2 - 1;
            if (
                UNIVERSAL::isa(
                    $args->{location}, 'SpreadsheetModel::CalcBlock'
                )
              )
            {
                $r3 = $args->{location}{items}[0]{$wb}{row} - 1;
            }
            $args = [
                name       => "$args->{name}",
                categories => '='
                  . $w2
                  . xl_rowcol_to_cell( $r3, $c2, 1, 1 ) . ':'
                  . xl_rowcol_to_cell( $r3, $c2 + $args->lastCol, 1, 1 ),
                values => '='
                  . $w2
                  . xl_rowcol_to_cell( $r2, $c2, 1, 1 ) . ':'
                  . xl_rowcol_to_cell( $r2, $c2 + $args->lastCol, 1, 1 ),
            ];
        }
        $chart->$verb(@$args);
    }
}

sub wsWrite {
    my ( $self, $wb, $ws, $row, $col ) = @_;
    return if $self->{$wb};
    my $chart = $wb->add_chart(
        %$self,
        embedded => $ws ? 1 : 0,
        name => $self->objectShortName,
    );
    $self->applyInstructions( $chart, $wb, $ws, $self->{instructions} );

    if ($ws) {
        ( $row, $col ) = ( ( $ws->{nextFree} ||= -1 ) + 1, 0 )
          unless defined $row && defined $col;

        if ( $self->{name} ) {
            $ws->write( $row, $col, "$self->{name}", $wb->getFormat('notes') );
            $ws->set_row( $row, 21 );
            $row += 2;
        }

        $ws->set_row( $row, $self->{height} * 0.75 );
        $ws->insert_chart(
            $row, $col + 1, $chart, 0, 0,
            $self->{width} / 480.0,
            $self->{height} / 288.0
        );
        $row += 2;
        $ws->{nextFree} = $row unless $ws->{nextFree} > $row;
    }

}

1;
