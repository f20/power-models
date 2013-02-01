package SpreadsheetModel::SumProduct;

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
    'Sum-product calculation';
}

sub populateCore {
    my ($self) = @_;
    $self->{core}{$_} = $self->{$_}->getCore foreach qw(matrix vector);
}

sub check {

    my ($self) = @_;

    my ( $matrix, $vector ) = @{$self}{qw(matrix vector)};

    $self->{arithmetic} = '=SUMPRODUCT(IV1, IV2)';
    $self->{arguments} = { IV1 => $self->{matrix}, IV2 => $self->{vector} };
    push @{ $self->{sourceLines} }, $matrix, $vector;

    if (   $self->{rows}
        && defined $matrix->{cols}
        && defined $vector->{cols}
        && $matrix->{cols} == $vector->{cols}
        && $matrix->{rows}
        && $vector->{rows}
        && $matrix->{rows} == $vector->{rows}
        && $matrix->{rows}{groups}
        && $self->{rows}{list} == $matrix->{rows}{groups} )
    {
        $self->{summingByRowGroup} = 1;
        $self->{cols}              = $vector->{cols};
        return $self->SUPER::check;
    }

    if (   defined $self->{rows}
        && $self->{cols}
        && $matrix->{cols}
        && $matrix->{cols}{groups}
        && $self->{cols}{list} == $matrix->{cols}{groups}
        && $vector->{cols}{list}
        && !grep { !ref $_ || $_->{list} != $vector->{cols}{list} }
        @{ $self->{cols}{list} } )
    {
        foreach ( $matrix, $vector ) {

            return <<EOM
Mismatch (SumProduct):
$_->{name}: $_->{rows} x $_->{cols}
 - v -
$self->{name}: $self->{rows} x $self->{cols}
EOM
              if

              $_->{rows}
              and $_->{rows} != $self->{rows}
              and !$self->{rows}
              || !$self->{rows}{groups}
              || $_->{rows}{list} != $self->{rows}{groups}
              and !$self->{rows}
              || !defined $self->{rows}->supersetIndex( $_->{rows} );

        }
        $self->{summingByColumnGroup} = 1;
        return $self->SUPER::check;
    }

    if (   !defined $self->{rows}
        && !defined $self->{cols}
        && $vector->{rows}
        && $vector->{rows} == $matrix->{rows} )
    {
        $self->{rows} = $vector->{cols};
        $self->{cols} = $matrix->{cols};
        return $self->SUPER::check;
    }

    if (   defined $self->{cols}
        && ( !$self->{cols} || !$#{ $self->{cols}{list} } )
        && defined $self->{rows}
        && $vector->{rows}
        && $matrix->{rows}
        && $vector->{rows} == $matrix->{rows}
        && $vector->{cols}
        && $matrix->{cols}
        && $vector->{cols} == $matrix->{cols} )
    {
        die
          unless $self->{rowIndex} =
          $self->{rows}->supersetIndex( $matrix->{rows} );
        return $self->SUPER::check;
    }

    if (   !defined $self->{rows}
        && !defined $self->{cols}
        && $vector->{cols}
        && $vector->{cols} == $matrix->{cols} )
    {
        $self->{cols} = $vector->{rows};
        $self->{rows} = $matrix->{rows};
        return $self->SUPER::check;
    }

    if (  !defined $self->{rows}
        && defined $self->{cols}
        && ( !$self->{cols}   || !$#{ $self->{cols}{list} } )
        && ( !$vector->{rows} || !$#{ $vector->{rows}{list} } )
        && $vector->{cols}
        && $vector->{cols} == $matrix->{cols} )
    {
        $self->{rows} = $matrix->{rows};
        return $self->SUPER::check;
    }

    return <<ERR if !defined $self->{rows} || !defined $self->{cols};
Problem in matrix-vector style SumProduct:
matrix: $matrix->{name} ($matrix->{rows} x $matrix->{cols})
vector: $vector->{name} ($vector->{rows} x $vector->{cols})
ERR

    <<ERR ;
Problem in SumProduct:
self: $self->{name} ($self->{rows} x $self->{cols})
matrix: $matrix->{name} ($matrix->{rows} x $matrix->{cols})
vector: $vector->{name} ($vector->{rows} x $vector->{cols})
ERR

}

sub wsPrepare {

    my ( $self, $wb, $ws ) = @_;

    my ( $matsheet, $matr, $matc ) = $self->{matrix}->wsWrite( $wb, $ws );
    $matsheet =
      $matsheet == $ws
      ? ''
      : "'" . ( $matsheet ? $matsheet->get_name : 'BROKEN LINK' ) . "'!";

    my ( $vecsheet, $vecr, $vecc ) = $self->{vector}->wsWrite( $wb, $ws );
    $vecsheet =
      $vecsheet == $ws
      ? ''
      : "'" . ( $vecsheet ? $vecsheet->get_name : 'BROKEN LINK' ) . "'!";

    my $formula =
      $ws->store_formula("=SUMPRODUCT(${matsheet}IV1:IV2,${vecsheet}IV3:IV4)");
    my $format = $wb->getFormat( $self->{defaultFormat} || '0.000soft' );

    if ( $self->{summingByRowGroup} ) {
        my ( @start, @end );
        my $groupid = $self->{matrix}{rows}{groupid};
        my $gr1;
        for ( my $i = 0 ; $i <= @$groupid ; ++$i ) {
            my $gr = $groupid->[$i];
            next if defined $gr && defined $gr1 && $gr == $gr1;
            $start[$gr] = $i     if defined $gr;
            $end[$gr1]  = $i - 1 if defined $gr1;
            $gr1        = $gr;
        }
        return sub {
            my ( $x, $y ) = @_;
            '', $format, $formula,
              IV1 => xl_rowcol_to_cell( $matr + $start[$y], $matc + $x, 1, 0 ),
              IV2 => xl_rowcol_to_cell( $matr + $end[$y],   $matc + $x, 1, 0 ),
              IV3 => xl_rowcol_to_cell( $vecr + $start[$y], $vecc + $x, 1, 0 ),
              IV4 => xl_rowcol_to_cell( $vecr + $end[$y],   $vecc + $x, 1, 0 );
        };
    }

    if ( $self->{summingByColumnGroup} ) {
        my @mody = map {
            my $c = $_->{rows};
            $c == $self->{rows}     ? 0
              : !$c                 ? 1
              : $c == $self->{cols} ? 2
              : $self->{rows}{groups} && $c->{list} == $self->{rows}{groups} ? 3
              :   $self->{rows}->supersetIndex($c);
        } @{$self}{qw(matrix vector)};
        my $n     = $self->{vector}->lastCol;
        my $veccl = $vecc + $n;
        $formula = $ws->store_formula("=${matsheet}IV1*${vecsheet}IV3")
          unless $n;
        my $mat0 = $matc + ( $n ? 1 : 0 );
        my $n1 = $n ? $n + 2 : 1;
        return sub {
            my ( $x, $y ) = @_;
            my $matcoff = $mat0 + $x * $n1;
            my ( $my, $myl, $vy, $vyl ) = map {
                    ref $_ ? ( $_->[$y] < 0 ? -1 - $_->[$y] : $_->[$y] )
                  : $_ == 0 ? $y
                  : $_ == 1 ? 0
                  : $_ == 2 ? $x
                  : $_ == 3 ? $self->{rows}{groupid}[$y]
                  : $_ < 0  ? $y % -$_
                  : die, ref $_ ? $_->[$y] >= 0
                  :               $_ > 0;
            } @mody;
            '', $format, $formula,
              IV1 => xl_rowcol_to_cell( $matr + $my, $matcoff,      $myl, 1 ),
              IV2 => xl_rowcol_to_cell( $matr + $my, $matcoff + $n, $myl, 1 ),
              IV3 => xl_rowcol_to_cell( $vecr + $vy, $vecc,         $vyl, 1 ),
              IV4 => xl_rowcol_to_cell( $vecr + $vy, $veccl,        $vyl, 1 );
        };
    }

    if (   $self->{rowIndex}
        && $self->{vector}->{rows}
        && $self->{vector}->{rows} == $self->{matrix}->{rows}
        && $self->{vector}{cols}
        && $self->{vector}{cols} == $self->{matrix}{cols}
        && ( !$self->{cols} || !$#{ $self->{cols}{list} } ) )
    {
        my $n = $self->{vector}->lastCol;
        $formula = $ws->store_formula("=${matsheet}IV1*${vecsheet}IV3")
          unless $n;
        my $matcl = $matc + $n;
        my $veccl = $vecc + $n;
        return sub {
            my ( $x, $y ) = @_;
            $y = $self->{rowIndex}[$y];
            '', $format, $formula,
              IV1 => xl_rowcol_to_cell( $matr + $y, $matc,  0, 1 ),
              IV2 => xl_rowcol_to_cell( $matr + $y, $matcl, 0, 1 ),
              IV3 => xl_rowcol_to_cell( $vecr + $y, $vecc,  0, 1 ),
              IV4 => xl_rowcol_to_cell( $vecr + $y, $veccl, 0, 1 );
        };
    }

    if (   $self->{vector}->{rows}
        && $self->{vector}->{rows} == $self->{matrix}->{rows} )
    {
        my $n = $self->{vector}->lastRow;
        $formula = $ws->store_formula("=${matsheet}IV1*${vecsheet}IV3")
          unless $n;
        my $matrl = $matr + $n;
        my $vecrl = $vecr + $n;
        return sub {
            my ( $x, $y ) = @_;
            '', $format, $formula,
              IV1 => xl_rowcol_to_cell( $matr,  $matc + $x, 1, 0 ),
              IV2 => xl_rowcol_to_cell( $matrl, $matc + $x, 1, 0 ),
              IV3 => xl_rowcol_to_cell( $vecr,  $vecc + $y, 1, 1 ),
              IV4 => xl_rowcol_to_cell( $vecrl, $vecc + $y, 1, 1 );
        };
    }

    if (   $self->{vector}{cols}
        && $self->{vector}{cols} == $self->{matrix}{cols} )
    {
        my $n = $self->{vector}->lastCol;
        $formula = $ws->store_formula("=${matsheet}IV1*${vecsheet}IV3")
          unless $n;
        my $matcl = $matc + $n;
        my $veccl = $vecc + $n;
        return sub {
            my ( $x, $y ) = @_;
            '', $format, $formula,
              IV1 => xl_rowcol_to_cell( $matr + $y, $matc,  0, 1 ),
              IV2 => xl_rowcol_to_cell( $matr + $y, $matcl, 0, 1 ),
              IV3 => xl_rowcol_to_cell( $vecr + $x, $vecc,  1, 1 ),
              IV4 => xl_rowcol_to_cell( $vecr + $x, $veccl, 1, 1 );
        };
    }

}

1;
