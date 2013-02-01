package SpreadsheetModel::GroupBy;

=head Copyright licence and disclaimer

Copyright 2008-2013 Reckon LLP and others. All rights reserved.

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

require SpreadsheetModel::Dataset;
our @ISA = qw(SpreadsheetModel::Dataset);

use Spreadsheet::WriteExcel::Utility;

sub objectType {
    'Cell summation';
}

sub populateCore {
    my ($self) = @_;
    $self->{core}{$_} = $self->{$_}->getCore foreach qw(source);
}

sub check {

    my ($self) = @_;

    return "No source $self->{debug}" unless $self->{source};

    if (   $self->{cols}
        && !$self->{rows}
        && $self->{source}{rows}
        && $self->{source}{rows}{groups}
        && $self->{cols}{list} == $self->{source}{rows}{groups} )
    {
        $self->{specialTransposed} = 1;
    }
    else {
        return "Mismatch in cols $self->{debug}"
          unless !$self->{cols}

     #   || !$self->{source}{cols} -- we don't do extending as part of a GroupBy
          || $self->{cols} == $self->{source}{cols}
          || $self->{cols}{list} == $self->{source}{cols}{groups};

        return "Mismatch in rows $self->{debug}"
          unless !$self->{rows}

     #   || !$self->{source}{rows} -- we don't do extending as part of a GroupBy
          || $self->{rows} == $self->{source}{rows}
          || $self->{rows}{list} == $self->{source}{rows}{groups};
    }
    push @{ $self->{sourceLines} }, $self->{source};

    $self->{arithmetic} = '=SUM(IV1)';
    $self->{arguments} = { IV1 => $self->{source} };
    $self->SUPER::check;

}

sub wsPrepare {

    my ( $self, $wb, $ws ) = @_;
    my $broken;
    my ( $srcsheet, $srcr, $srcc ) = $self->{source}->wsWrite( $wb, $ws );
    $broken = "UNFEASIBLE LINK to source for $self->{name} $self->{debug}"
      unless $srcsheet;
    $srcsheet = !$srcsheet
      || $srcsheet == $ws ? '' : "'" . $srcsheet->get_name . "'!";

    my $formula = $ws->store_formula("=SUM(${srcsheet}IV1:IV2)");
    my $format = $wb->getFormat( $self->{defaultFormat} || '0.000soft' );

    my ( $xabs, $yabs ) = ( 1, 1 );
    my ( @x1, @x2, @y1, @y2 );

    if ( $self->{specialTransposed} || $self->{cols} == $self->{source}{cols} )
    {
        $xabs = 0;
    }
    elsif ( $self->{cols} ) {
        my $x = 0;
        for ( 0 .. $#{ $self->{cols}{list} } ) {
            if (   $self->{source}{cols}{noCollapse}
                || $#{ $self->{cols}{list}[$_]{list} } )
            {
                $x1[$_] = ++$x;
                $x2[$_] = $x += $#{ $self->{cols}{list}[$_]{list} };
                $x++;
            }
            else {
                $x1[$_] = $x;
                $x2[$_] = $x;
                $x++;
            }
        }
    }
    else {
        $x1[0] = 0;
        $x2[0] = $self->{source}->lastCol;
    }

    if ( $self->{rows} == $self->{source}{rows} ) {
        $yabs = 0;
    }
    elsif ($self->{specialTransposed}
        || $self->{rows} )
    {
        my $y = 0;
        for ( 0 .. $#{ $self->{source}{rows}{groups} } ) {
            if (   $self->{source}{rows}{noCollapse}
                || $#{ $self->{source}{rows}{groups}[$_]{list} } )
            {
                $y1[$_] = ++$y;
                $y2[$_] = $y += $#{ $self->{source}{rows}{groups}[$_]{list} };
                $y++;
            }
            else {
                $y1[$_] = $y;
                $y2[$_] = $y;
                $y++;
            }
        }
    }
    else {
        $y1[0] = 0;
        $y2[0] = $self->{source}->lastRow;
    }

    $self->{specialTransposed}

      ?

      sub {
        die $broken if $broken;
        my ( $y,  $x )  = @_;
        my ( $x1, $x2 ) = $xabs ? ( $x1[$x], $x2[$x] ) : ( $x, $x );
        my ( $y1, $y2 ) = $yabs ? ( $y1[$y], $y2[$y] ) : ( $y, $y );
        '', $format, $formula,
          IV1 => xl_rowcol_to_cell( $srcr + $y1, $srcc + $x1, 1, 1 ),
          IV2 => xl_rowcol_to_cell( $srcr + $y2, $srcc + $x2, 1, 1 );
      }

      :

      sub {
        die $broken if $broken;
        my ( $x,  $y )  = @_;
        my ( $x1, $x2 ) = $xabs ? ( $x1[$x], $x2[$x] ) : ( $x, $x );
        my ( $y1, $y2 ) = $yabs ? ( $y1[$y], $y2[$y] ) : ( $y, $y );
        '', $format, $formula,
          IV1 => xl_rowcol_to_cell( $srcr + $y1, $srcc + $x1, $yabs, $xabs ),
          IV2 => xl_rowcol_to_cell( $srcr + $y2, $srcc + $x2, $yabs, $xabs );
      };

}

1;
