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
        if (    $verb eq 'add_series'
            and UNIVERSAL::isa( $args, 'SpreadsheetModel::Dataset' )
            and !$args->{rows} && $args->{cols}
            || $args->{rows}   && !$args->{cols} )
        {
            push @{ $self->{sourceLines} }, $args
              unless $self->{sourceLines} && grep { $_ == $args }
              @{ $self->{sourceLines} };
            my ( $w2, $r2, $c2 ) = $args->wsWrite( $wb, $ws, undef, undef, 1 );
            $w2 = "'" . $w2->get_name . "'!";
            my $r3 = $r2;
            my $c3 = $c2;
            if ( $args->{cols} ) {
                if (
                    UNIVERSAL::isa(
                        $args->{location}, 'SpreadsheetModel::CalcBlock'
                    )
                  )
                {
                    $r3 = $args->{location}{items}[0]{$wb}{row};
                }
                --$r3;
            }
            else {
                if (
                    UNIVERSAL::isa(
                        $args->{location}, 'SpreadsheetModel::Columnset'
                    )
                  )
                {
                    $c3 = $args->{location}{columns}[0]{$wb}{col};
                }
                --$c3;
            }
            $args = [
                name       => $args->objectShortName,
                categories => '='
                  . $w2
                  . xl_rowcol_to_cell( $r3, $c3, 1, 1 ) . ':'
                  . xl_rowcol_to_cell(
                    $r3 + $args->lastRow,
                    $c3 + $args->lastCol,
                    1, 1
                  ),
                values => '='
                  . $w2
                  . xl_rowcol_to_cell( $r2, $c2, 1, 1 ) . ':'
                  . xl_rowcol_to_cell(
                    $r2 + $args->lastRow,
                    $c2 + $args->lastCol,
                    1, 1
                  ),
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
            ++$row;
        }

        if ( $self->{lines}
            or !( $wb->{noLinks} && $wb->{noLinks} == 1 )
            and $self->{name} && $self->{sourceLines} )
        {
            my $hideFormulas = $wb->{noLinks} && $self->{sourceLines};
            my $textFormat   = $wb->getFormat('text');
            my $linkFormat   = $wb->getFormat('link');
            my $xc           = 0;
            foreach (
                $self->{lines} ? @{ $self->{lines} } : (),
                !( $wb->{noLinks} && $wb->{noLinks} == 1 )
                && $self->{sourceLines} && @{ $self->{sourceLines} }
                ? ( 'Data sources:', @{ $self->{sourceLines} } )
                : ()
              )
            {
                if ( UNIVERSAL::isa( $_, 'SpreadsheetModel::Object' ) ) {
                    my $na = 'x' . ( ++$xc ) . " = $_->{name}";
                    if ( my $url = $_->wsUrl($wb) ) {
                        $ws->set_row( $row, undef, undef, 1, 1 )
                          if $hideFormulas;
                        $ws->write_url( $row++, $col, $url, $na, $linkFormat );
                        (
                            $_->{location}
                              && UNIVERSAL::isa( $_->{location},
                                'SpreadsheetModel::Columnset' )
                            ? $_->{location}
                            : $_
                          )->addForwardLink($self)
                          if $wb->{findForwardLinks};
                    }
                    else {
                        $ws->set_row( $row, undef, undef, 1, 1 )
                          if $hideFormulas;
                        $ws->write_string( $row++, $col, $na, $textFormat );
                    }
                }
                elsif (/^(https?|mailto:)/) {
                    $ws->set_row( $row, undef, undef, 1, 1 ) if $hideFormulas;
                    $ws->write_url( $row++, $col, "$_", "$_", $linkFormat );
                }
                else {
                    $ws->set_row( $row, undef, undef, 1, 1 ) if $hideFormulas;
                    $ws->write_string( $row++, $col, "$_", $textFormat );
                }
            }
            $ws->set_row( $row, undef, undef, undef, 0, 0, 1 )
              if $hideFormulas;
        }

        ++$row;
        $ws->set_row( $row, $self->{height} * 0.75 );
        $ws->insert_chart(
            $row, $col + 1, $chart, 0, 0,
            $self->{width} / 480.0,
            $self->{height} / 288.0
        );
        $row += 2;
        $ws->{nextFree} = $row unless $ws->{nextFree} > $row;
    }
    else {    # Chartsheet
        $chart->protect();
    }

}

1;
