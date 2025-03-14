﻿package SpreadsheetModel::Chart;

# Copyright 2015-2021 Franck Latrémolière and others.
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

require SpreadsheetModel::Object;
our @ISA = qw(SpreadsheetModel::Object);
use Spreadsheet::WriteExcel::Utility;

sub objectType {
    'Chart';
}

sub chartsheetPriority {
    my ($self) = @_;
    $self->{chartsheetPriority} ||= 5;
}

sub check {
    my ($self) = @_;
    return
        "Broken chart $self->{name} $self->{debug}: instructions is "
      . ( ref $self->{instructions} )
      . ' but must be ARRAY'
      unless ref $self->{instructions} eq 'ARRAY';

    $self->{scaling_factor} ||= 1;
    $self->{width}          ||= 600 * $self->{scaling_factor};
    $self->{height}         ||= 400 * $self->{scaling_factor};

    return;
}

sub wsWrite {
    my ( $self, $wb, $ws, $row, $col ) = @_;
    $self->applyInstructions( $self->wsCreate( $wb, $ws, $row, $col ),
        $wb, $ws, $self->{instructions} );
}

sub applyInstructions {
    my ( $self, $chart, $wb, $ws, $instructions ) = @_;
    $chart->set_plotarea( border => { none => 1 } );
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
        if (   'ARRAY' eq ref $args
            && $args->[0]
            && $args->[0] eq 'name_formula'
            && UNIVERSAL::isa( $args->[1], 'SpreadsheetModel::Dataset' )
            && !$args->[1]{rows}
            && !$args->[1]{cols} )
        {
            push @{ $self->{sourceLines} }, $args->[1]
              unless $self->{sourceLines} && grep { $_ == $args->[1] }
              @{ $self->{sourceLines} };
            my ( $w2, $r2, $c2 ) =
              $args->[1]->wsWrite( $wb, $ws, undef, undef, 1 );
            $args =
              [     name_formula => "='"
                  . $w2->get_name . "'!"
                  . xl_rowcol_to_cell( $r2, $c2 ) ];
        }
        if ( $verb eq 'add_series' ) {
            my $series;
            if ( ref $args eq 'ARRAY' ) {
                ( $series, my @newArgs ) = @$args;
                $args = \@newArgs;
            }
            else {
                $series = $args;
                $args   = [];
            }
            next unless $series;
            if ( ref $series eq 'CODE' ) {
                $chart->$verb( $series->( $wb, $ws ), @$args );
                next;
            }
            elsif ( UNIVERSAL::can( $series, 'valuesNameCategories' ) ) {
                $chart->$verb( $series->valuesNameCategories( $wb, $ws ),
                    @$args );
                next;
            }
            elsif ( UNIVERSAL::isa( $series, 'SpreadsheetModel::Dataset' )
                and $series->lastCol || $series->lastRow )
            {
                push @{ $self->{sourceLines} }, $series
                  unless $self->{sourceLines} && grep { $_ == $series }
                  @{ $self->{sourceLines} };
                my ( $w2, $r2, $c2 ) =
                  $series->wsWrite( $wb, $ws, undef, undef, 1 );
                $w2 = "'" . $w2->get_name . "'!";
                my $r3 = $r2;
                my $c3 = $c2;
                if ( $series->lastCol ) {
                    if (
                        UNIVERSAL::isa(
                            $series->{location}, 'SpreadsheetModel::CalcBlock'
                        )
                      )
                    {
                        $r3 = $series->{location}{items}[0]{$wb}{row};
                    }
                    --$r3;
                }
                else {
                    if (
                        UNIVERSAL::isa(
                            $series->{location}, 'SpreadsheetModel::Columnset'
                        )
                      )
                    {
                        $c3 = $series->{location}{columns}[0]{$wb}{col};
                    }
                    --$c3;
                }
                my ( $w4, $r4, $c4 );
                if ( $series->{legendText} ) {
                    ( $w4, $r4, $c4 ) =
                      $series->{legendText}
                      ->wsWrite( $wb, $ws, undef, undef, 1 );
                    $w4 = "'" . $w4->get_name . "'!";
                }
                map {
                    $chart->$verb(
                        $w4
                        ? (
                                name_formula => '='
                              . $w4
                              . xl_rowcol_to_cell(
                                $r4 + $_->[0],
                                $c4 + $_->[1],
                                1, 1
                              )
                          )
                        : ( name => $_->[4] ),
                        categories => '='
                          . $w2
                          . xl_rowcol_to_cell(
                            $r3 + ( $self->{ignore_top}  || 0 ),
                            $c3 + ( $self->{ignore_left} || 0 ),
                            1, 1,
                          )
                          . ':'
                          . xl_rowcol_to_cell(
                            $r3 + $_->[2] - ( $self->{ignore_bottom} || 0 ),
                            $c3 + $_->[3] - ( $self->{ignore_right}  || 0 ),
                            1,
                            1,
                          ),
                        values => '='
                          . $w2
                          . xl_rowcol_to_cell(
                            $r2 + $_->[0] + ( $self->{ignore_top}  || 0 ),
                            $c2 + $_->[1] + ( $self->{ignore_left} || 0 ),
                            1, 1,
                          )
                          . ':'
                          . xl_rowcol_to_cell(
                            $r2 + $_->[2] - ( $self->{ignore_bottom} || 0 ),
                            $c2 + $_->[3] - ( $self->{ignore_right}  || 0 ),
                            1,
                            1,
                          ),
                        @$args
                    );
                  } !$series->lastCol
                  ? [ 0, 0, $series->lastRow, 0, $series->objectShortName ]
                  : !$series->lastRow
                  ? [ 0, 0, 0, $series->lastCol, $series->objectShortName ]
                  : map {
                    [ $_, 0, $_, $series->lastCol, $series->{rows}{list}[$_] ];
                  } $series->{rows}->indices;
                next;
            }
            elsif (ref $series eq 'ARRAY'
                && UNIVERSAL::isa( $series->[0], 'SpreadsheetModel::Dataset' )
                && UNIVERSAL::isa( $series->[1], 'SpreadsheetModel::Dataset' )
                and !$series->[0]{rows}
                && !$series->[1]{rows}
                && $series->[0]{cols}
                && $series->[1]{cols}
                && $series->[0]{cols} == $series->[1]{cols}
                || !$series->[0]{cols}
                && !$series->[1]{cols}
                && $series->[0]{rows}
                && $series->[1]{rows}
                && $series->[0]{rows} == $series->[1]{rows} )
            {    # Scatter plot
                foreach my $d (@$series) {
                    push @{ $self->{sourceLines} }, $d
                      unless $self->{sourceLines} && grep { $_ == $d }
                      @{ $self->{sourceLines} };
                }
                my ( $w2, $r2, $c2 ) =
                  $series->[1]->wsWrite( $wb, $ws, undef, undef, 1 );
                $w2 = "'" . $w2->get_name . "'!";
                my ( $w3, $r3, $c3 ) =
                  $series->[0]->wsWrite( $wb, $ws, undef, undef, 1 );
                $w3 = "'" . $w3->get_name . "'!";
                unshift @$args,
                  name       => $series->[1]->objectShortName,
                  categories => '='
                  . $w3
                  . xl_rowcol_to_cell(
                    $r3 + ( $self->{ignore_top}  || 0 ),
                    $c3 + ( $self->{ignore_left} || 0 ),
                    1, 1,
                  )
                  . ':'
                  . xl_rowcol_to_cell(
                    $r3 + $series->[0]->lastRow -
                      ( $self->{ignore_bottom} || 0 ),
                    $c3 + $series->[0]->lastCol -
                      ( $self->{ignore_right} || 0 ),
                    1,
                    1,
                  ),
                  values => '='
                  . $w2
                  . xl_rowcol_to_cell(
                    $r2 + ( $self->{ignore_top}  || 0 ),
                    $c2 + ( $self->{ignore_left} || 0 ),
                    1, 1,
                  )
                  . ':'
                  . xl_rowcol_to_cell(
                    $r2 + $series->[1]->lastRow -
                      ( $self->{ignore_bottom} || 0 ),
                    $c2 + $series->[1]->lastCol -
                      ( $self->{ignore_right} || 0 ),
                    1,
                    1,
                  );
            }
            else {
                warn "Something has probably gone wrong with @$args";
                next;
            }
        }
        $chart->$verb(@$args);
    }
}

sub wsCreate {

    my ( $self, $wb, $ws, $row, $col ) = @_;
    return $self->{$wb} if $self->{$wb};

    my $chart = $wb->add_chart(
        %$self,
        embedded => $ws ? 1 : 0,
        name     => $self->objectShortName,
    );

    if ( !$ws ) {    # Chartsheet

        # Does not work: problem with the <c:protection> element in the chart?
        $chart->protect() unless $self->{unprotected};

        return $self->{$wb} = $chart;

    }

    ( $row, $col ) = ( ( $ws->{nextFree} ||= -1 ) + 1, 0 )
      unless defined $row && defined $col;

    if ( $self->{name} ) {
        if (   ref $self->{name} eq 'ARRAY'
            && UNIVERSAL::isa( $self->{name}[0], 'SpreadsheetModel::Dataset' )
            && defined $self->{name}[1]
            && defined $self->{name}[2] )
        {
            push @{ $self->{sourceLines} }, $self->{name}[0]
              unless !$self->{sourceLines}
              || grep { $_ == $self->{name}[0] } @{ $self->{sourceLines} };
            my ( $w2, $r2, $c2 ) =
              $self->{name}[0]->wsWrite( $wb, $ws, undef, undef, 1 );
            $ws->write(
                $row, $col,
                "='"
                  . $w2->get_name . "'!"
                  . xl_rowcol_to_cell(
                    $r2 + $self->{name}[1],
                    $c2 + $self->{name}[2]
                  ),
                $wb->getFormat('notes')
            );
        }
        else {
            $ws->write( $row, $col, "$self->{name}", $wb->getFormat('notes') );
        }
        $ws->set_row( $row, $wb->{captionRowHeight} );
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
                          && UNIVERSAL::can( $_->{location}, 'wsWrite' )
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
                $ws->set_row( $row, undef, undef, 1, 1 )
                  if $hideFormulas;
                $ws->write_url( $row++, $col, "$_", "$_", $linkFormat );
            }
            else {
                $ws->set_row( $row, undef, undef, 1, 1 )
                  if $hideFormulas;
                $ws->write_string( $row++, $col, "$_", $textFormat );
            }
        }
        $ws->set_row( $row, undef, undef, undef, 0, 0, 1 )
          if $hideFormulas;
    }

    ++$row;
    $ws->set_row( $row, $self->{height} * 0.75 );
    $ws->insert_chart(
        $row, $col, $chart, 0, 0,
        $self->{width} / 480.0,
        $self->{height} / 288.0
    );
    $row += 2;
    $ws->{nextFree} = $row unless $ws->{nextFree} > $row;

    $self->{$wb} = $chart;

}

1;
