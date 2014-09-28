package Compilation;

=head Copyright licence and disclaimer

Copyright 2009-2014 Franck Latrémolière, Reckon LLP and others.

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

sub edcmTariffImpact {
    my ( $self, $wbmodule, %options ) = @_;
    $options{components} ||= [ split /\n/, <<EOL];
Import super-red unit rate (p/kWh)
Import fixed charge (p/day)
Import capacity rate (p/kVA/day)
Import exceeded capacity rate (p/kVA/day)
Export super-red unit rate (p/kWh)
Export fixed charge (p/day)
Export capacity rate (p/kVA/day)
Export exceeded capacity rate (p/kVA/day)
EOL
    $options{format1} ||= [
        qw(0.000copy 0.00copy 0.00copy 0.00copy 0.000copy 0.00copy 0.00copy),
        [ base => '0.00copy', right => 5, right_color => 8 ]
    ];
    $options{format2} ||= [
        qw(0.000softpm 0.00softpm 0.00softpm 0.00softpm 0.000softpm 0.00softpm 0.00softpm),
        [ base => '0.00softpm', right => 5, right_color => 8 ]
    ];
    $options{tableNumber}     ||= 4501;
    $options{firstColumn}     ||= 2;
    $options{nameExtraColumn} ||= 1;
    $self->genericTariffImpact( $wbmodule, %options );
}

sub cdcmTariffImpact {

    my ( $self, $wbmodule, %options ) = @_;

    $options{linesAfter} ||= [ split /\n/, <<EOL ];
Domestic Unrestricted
Domestic Two Rate
Domestic Off Peak (related MPAN)
Small Non Domestic Unrestricted
Small Non Domestic Two Rate
Small Non Domestic Off Peak (related MPAN)
LV Medium Non-Domestic
LV Sub Medium Non-Domestic
HV Medium Non-Domestic
LV HH Metered
LV Sub HH Metered
HV HH Metered
NHH UMS category A
NHH UMS category B
NHH UMS category C
NHH UMS category D
LV UMS (Pseudo HH Metered)
LV Generation NHH
LV Sub Generation NHH
LV Generation Intermittent
LV Generation Non-Intermittent
LV Sub Generation Intermittent
LV Sub Generation Non-Intermittent
HV Generation Intermittent
HV Generation Non-Intermittent
LDNO LV: Domestic Unrestricted
LDNO LV: Domestic Two Rate
LDNO LV: Domestic Off Peak (related MPAN)
LDNO LV: Small Non Domestic Unrestricted
LDNO LV: Small Non Domestic Two Rate
LDNO LV: Small Non Domestic Off Peak (related MPAN)
LDNO LV: LV Medium Non-Domestic
LDNO LV: LV HH Metered
LDNO LV: NHH UMS category A
LDNO LV: NHH UMS category B
LDNO LV: NHH UMS category C
LDNO LV: NHH UMS category D
LDNO LV: LV UMS (Pseudo HH Metered)
LDNO LV: LV Generation NHH
LDNO LV: LV Generation Intermittent
LDNO LV: LV Generation Non-Intermittent
LDNO HV: Domestic Unrestricted
LDNO HV: Domestic Two Rate
LDNO HV: Domestic Off Peak (related MPAN)
LDNO HV: Small Non Domestic Unrestricted
LDNO HV: Small Non Domestic Two Rate
LDNO HV: Small Non Domestic Off Peak (related MPAN)
LDNO HV: LV Medium Non-Domestic
LDNO HV: LV HH Metered
LDNO HV: LV Sub HH Metered
LDNO HV: HV HH Metered
LDNO HV: NHH UMS category A
LDNO HV: NHH UMS category B
LDNO HV: NHH UMS category C
LDNO HV: NHH UMS category D
LDNO HV: LV UMS (Pseudo HH Metered)
LDNO HV: LV Generation NHH
LDNO HV: LV Sub Generation NHH
LDNO HV: LV Generation Intermittent
LDNO HV: LV Generation Non-Intermittent
LDNO HV: LV Sub Generation Intermittent
LDNO HV: LV Sub Generation Non-Intermittent
LDNO HV: HV Generation Intermittent
LDNO HV: HV Generation Non-Intermittent
EOL

    $options{components} ||= [ split /\n/, <<EOL];
Unit rate 1 p/kWh
Unit rate 2 p/kWh
Unit rate 3 p/kWh
Fixed charge p/MPAN/day
Capacity charge p/kVA/day
Reactive power charge p/kVArh
EOL

    unless ( $options{format1} ) {
        $options{format1} = [ map { /k(W|VAr)h/ ? '0.000copy' : '0.00copy'; }
              @{ $options{components} } ];
        $options{format1}[ $#{ $options{format1} } ] = [
            base        => $options{format1}[ $#{ $options{format1} } ],
            right       => 5,
            right_color => 8
        ];
    }

    unless ( $options{format2} ) {
        $options{format2} =
          [ map { /k(W|VAr)h/ ? '0.000softpm' : '0.00softpm'; }
              @{ $options{components} } ];
        $options{format2}[ $#{ $options{format2} } ] = [
            base        => $options{format2}[ $#{ $options{format2} } ],
            right       => 5,
            right_color => 8
        ];
    }

    $options{tableNumber} ||= 3701;
    $options{firstColumn} ||= 3;

    $self->genericTariffImpact( $wbmodule, %options );

}

sub genericTariffImpact {

    my ( $self, $wbmodule, %options ) = @_;

    _defaultOptions( \%options );

    my $wb = $wbmodule->new(
        "Impact tariffs $options{dcpName}" . $wbmodule->fileExtension );
    $wb->setFormats(
        {
            $options{colour} ? ( colour => $options{colour} ) : (),
            alignment => 1
        }
    );

    my $linesAfter = $options{linesAfter};
    my $linesBefore = $options{linesBefore} || $linesAfter;

    my $titleFormat = $wb->getFormat('notes');
    my $thFormat    = $wb->getFormat('th');
    my $thcFormat   = $wb->getFormat('thc');
    my $thcFormatB =
      $wb->getFormat( [ base => 'thc', right => 5, right_color => 8 ] );
    my $thcaFormat = $wb->getFormat(
        [
            base        => 'caption',
            align       => 'center_across',
            right       => 5,
            right_color => 8
        ]
    );

    my $ncol = @{ $options{components} };
    my @format1 =
      map { $wb->getFormat($_); } @{ $options{format1} };
    my @format2 =
      map { $wb->getFormat($_); } @{ $options{format2} };
    my @format3 =
      map { $wb->getFormat($_); } ( map { '%softpm' } 2 .. $ncol ),
      [ base => '%softpm', right => 5, right_color => 8 ];

    my @books = $self->listModels;
    my $findRow =
      $self->prepare(
        'select row from data where bid=? and tab=? and col=0 and v=?');
    my $q =
      $self->prepare(
        'select v from data where bid=? and tab=? and row=? and col=?');

    foreach my $i ( 0 .. $#{ $options{sheetNames} } ) {
        my $qr = $options{sheetNames}[$i];
        $qr =~ tr/ /-/;
        my ($bidb) =
          grep { $_->[1] =~ /$qr/ && $options{basematch}->( $_->[1] ) } @books;
        next unless $bidb;
        $bidb = $bidb->[0];
        my ($bida) =
          grep { $_->[1] =~ /$qr/ && $options{dcpmatch}->( $_->[1] ) } @books;
        next unless $bida;
        $bida = $bida->[0];

        unless ( $options{linesAfter} ) {
            $linesBefore = $linesAfter = [
                map { @$_ } @{
                    $self->selectall_arrayref(
                        'select v from data where bid=? and tab=?'
                          . ' and col=0 and row>0',
                        undef, $bida, $options{tableNumber}
                    )
                }
            ];
        }

        my $ws = $wb->add_worksheet( $options{sheetNames}[$i] );
        $ws->set_column( 0, 0,   48 );
        $ws->set_column( 1, 254, 12 );
        $ws->hide_gridlines(2);
        $ws->freeze_panes( 4, 1 );
        $ws->write_string( 0, 0, $options{sheetTitles}[$i], $titleFormat );

        $ws->write_string( 2, 1,         'Baseline prices',     $thcaFormat );
        $ws->write_string( 2, 1 + $ncol, 'Prices on new basis', $thcaFormat );
        $ws->write_string( 2, 1 + $ncol * 2, 'Price change',      $thcaFormat );
        $ws->write_string( 2, 1 + $ncol * 3, 'Percentage change', $thcaFormat );

        my $diff = $ws->store_formula('=IV2-IV1');
        my $perc = $ws->store_formula('=IF(IV1,IV3/IV2-1,"")');

        $ws->write( 2, $_, undef, $thcaFormat )
          foreach 2 .. $ncol,
          2 + $ncol .. $ncol * 2,
          2 + $ncol * 2 .. $ncol * 3,
          2 + $ncol * 3 .. $ncol * 4;

        my @list = @{ $options{components} };

        for ( my $j = 1 ; $j < 2 + 3 * $ncol ; $j += $ncol ) {
            for ( my $k = 0 ; $k < @list ; ++$k ) {
                $ws->write_string( 3, $j + $k, $list[$k],
                    $k == $#list ? $thcFormatB : $thcFormat );
            }
        }

        for ( my $j = 0 ; $j < @$linesAfter ; ++$j ) {
            $findRow->execute( $bidb, $options{tableNumber},
                $linesBefore->[$j] );
            my ($rowb) = $findRow->fetchrow_array;
            $findRow->execute( $bida, $options{tableNumber},
                $linesAfter->[$j] );
            my ($rowa) = $findRow->fetchrow_array;
            $ws->write_string(
                4 + $j,
                0,
                $options{nameExtraColumn}
                ? eval {
                    $q->execute(
                        $bida, $options{tableNumber},
                        $rowa, $options{nameExtraColumn}
                    );
                    my ($x) = $q->fetchrow_array;
                    $x =~ s/\[.*\]//;
                    "Tariff $linesAfter->[$j]: $x";
                }
                : $linesAfter->[$j],
                $thFormat
            );
            for ( my $k = 0 ; $k < $ncol ; ++$k ) {
                $q->execute(
                    $bidb, $options{tableNumber},
                    $rowb, $k + $options{firstColumn}
                );
                my ($vb) = $q->fetchrow_array;
                $q->execute(
                    $bida, $options{tableNumber},
                    $rowa, $k + $options{firstColumn}
                );
                my ($va) = $q->fetchrow_array;

                $ws->write( 4 + $j, $k + 1,         $vb, $format1[$k] );
                $ws->write( 4 + $j, $k + 1 + $ncol, $va, $format1[$k] );

                use Spreadsheet::WriteExcel::Utility;
                my $old = xl_rowcol_to_cell( 4 + $j, $k + 1 );
                my $new = xl_rowcol_to_cell( 4 + $j, $k + 1 + $ncol );
                $ws->repeat_formula(
                    4 + $j, $k + 1 + 2 * $ncol, $diff, $format2[$k],
                    IV1 => $old,
                    IV2 => $new,
                );
                $ws->repeat_formula(
                    4 + $j, $k + 1 + 3 * $ncol, $perc, $format3[$k],
                    IV1 => $old,
                    IV2 => $old,
                    IV3 => $new,
                );
            }
        }
    }

}

sub cdcmPpuImpact {

    my ( $self, $wbmodule, %options ) = @_;

    _defaultOptions( \%options );

    my $wb = $wbmodule->new(
        "Impact pence per unit $options{dcpName}" . $wbmodule->fileExtension );
    $wb->setFormats(
        {
            $options{colour} ? ( colour => $options{colour} ) : (),
            alignment => 1
        }
    );

    my $linesAfter = $options{linesAfter} || [ split /\n/, <<EOL ];
Domestic Unrestricted
Domestic Two Rate
Domestic Off Peak (related MPAN)
Small Non Domestic Unrestricted
Small Non Domestic Two Rate
Small Non Domestic Off Peak (related MPAN)
LV Medium Non-Domestic
LV Sub Medium Non-Domestic
HV Medium Non-Domestic
LV HH Metered
LV Sub HH Metered
HV HH Metered
NHH UMS category A
NHH UMS category B
NHH UMS category C
NHH UMS category D
LV UMS (Pseudo HH Metered)
LV Generation NHH
LV Sub Generation NHH
LV Generation Intermittent
LV Generation Non-Intermittent
LV Sub Generation Intermittent
LV Sub Generation Non-Intermittent
HV Generation Intermittent
HV Generation Non-Intermittent
EOL

    my $linesBefore = $options{linesBefore} || $linesAfter;

    my $titleFormat = $wb->getFormat('notes');
    my $thFormat    = $wb->getFormat('th');
    my $thcFormat   = $wb->getFormat('thc');
    my $thcaFormat  = $wb->getFormat('caption');
    my @format1 =
      map { $wb->getFormat($_); } ( map { '0.000copy' } 1 .. 1 );
    my @format2 =
      map { $wb->getFormat($_); } ( map { '0.000softpm' } 1 .. 1 );
    my @format3 =
      map { $wb->getFormat($_); } ( map { '%softpm' } 1 .. 1 );

    my @books   = $self->listModels;
    my $findRow = $self->prepare(
        'select row from data where bid=? and tab=3802 and col=0 and v=?');
    my $q = $self->prepare(
        'select v from data where bid=? and tab=3802 and row=? and col=?');

    foreach my $i ( 0 .. $#{ $options{sheetNames} } ) {
        my $qr = $options{sheetNames}[$i];
        $qr =~ tr/ /-/;
        my ($bidb) =
          grep { $_->[1] =~ /$qr/ && $options{basematch}->( $_->[1] ) } @books;
        next unless $bidb;
        $bidb = $bidb->[0];
        my ($bida) =
          grep { $_->[1] =~ /$qr/ && $options{dcpmatch}->( $_->[1] ) } @books;
        next unless $bida;
        $bida = $bida->[0];
        my $ws = $wb->add_worksheet( $options{sheetNames}[$i] );
        $ws->set_column( 0, 0,   48 );
        $ws->set_column( 1, 254, 16 );
        $ws->hide_gridlines(2);
        $ws->freeze_panes( 1, 1 );
        $ws->write_string( 0, 0, $options{sheetTitles}[$i], $titleFormat );

        $ws->write_string( 2, 1, 'Baseline average p/kWh',     $thcFormat );
        $ws->write_string( 2, 2, 'Average p/kWh on new basis', $thcFormat );
        $ws->write_string( 2, 3, 'Change (p/kWh)',             $thcFormat );
        $ws->write_string( 2, 4, 'Percentage change',          $thcFormat );

        use Spreadsheet::WriteExcel::Utility;
        my $diff = $ws->store_formula('=IV2-IV1');
        my $perc = $ws->store_formula('=IF(IV1,IV3/IV2-1,0)');

        for ( my $j = 0 ; $j < @$linesAfter ; ++$j ) {
            $ws->write_string( 3 + $j, 0, $linesAfter->[$j], $thFormat );
            $findRow->execute( $bidb, $linesBefore->[$j] );
            my ($rowb) = $findRow->fetchrow_array;
            $findRow->execute( $bida, $linesAfter->[$j] );
            my ($rowa) = $findRow->fetchrow_array;
            $q->execute( $bidb, $rowb, 8 );
            my ($vb) = $q->fetchrow_array;
            $q->execute( $bida, $rowa, 8 );
            my ($va) = $q->fetchrow_array;
            $ws->write( 3 + $j, 1, $vb, $format1[0] );
            $ws->write( 3 + $j, 2, $va, $format1[0] );
            my $old = xl_rowcol_to_cell( 3 + $j, 1 );
            my $new = xl_rowcol_to_cell( 3 + $j, 2 );
            $ws->repeat_formula(
                3 + $j, 3, $diff, $format2[0],
                IV1 => $old,
                IV2 => $new,
            );
            $ws->repeat_formula(
                3 + $j, 4, $perc, $format3[0],
                IV1 => $old,
                IV2 => $old,
                IV3 => $new,
            );
        }
    }
}

sub cdcmRevenueMatrixImpact {

    my ( $self, $wbmodule, %options ) = @_;

    $options{linesAfter} ||= [ split /\n/, <<EOL ];
Domestic Unrestricted
Domestic Two Rate
Domestic Off Peak (related MPAN)
Small Non Domestic Unrestricted
Small Non Domestic Two Rate
Small Non Domestic Off Peak (related MPAN)
LV Medium Non-Domestic
LV Sub Medium Non-Domestic
HV Medium Non-Domestic
LV HH Metered
LV Sub HH Metered
HV HH Metered
NHH UMS category A
NHH UMS category B
NHH UMS category C
NHH UMS category D
LV UMS (Pseudo HH Metered)
LV Generation NHH
LV Sub Generation NHH
LV Generation Intermittent
LV Generation Non-Intermittent
LV Sub Generation Intermittent
LV Sub Generation Non-Intermittent
HV Generation Intermittent
HV Generation Non-Intermittent
EOL

    $options{tableNumber} = 3901;
    $options{col1}        = 0;
    $options{col2}        = 25;
    $options{columns}     = [24];

    $self->revenueMatrixImpact( $wbmodule, %options );

}

sub edcmRevenueMatrixImpact {
    my ( $self, $wbmodule, %options ) = @_;
    $options{tableNumber}     = 4601;
    $options{col1}            = 9;
    $options{col2}            = 33;
    $options{columns}         = [ 16, 20 ];
    $options{nameExtraColumn} = 1;
    $self->revenueMatrixImpact( $wbmodule, %options );
}

sub revenueMatrixImpact {

    my ( $self, $wbmodule, %options ) = @_;

    _defaultOptions( \%options );

    my $wb =
      $wbmodule->new(
        "Impact revenue $options{dcpName}" . $wbmodule->fileExtension );
    $wb->setFormats(
        { $options{colour} ? ( colour => $options{colour} ) : () } );

    my $linesAfter = $options{linesAfter};
    my $linesBefore = $options{linesBefore} || $linesAfter;

    my $titleFormat = $wb->getFormat('notes');
    my $thFormat    = $wb->getFormat('th');
    my $thcFormat   = $wb->getFormat('thc');
    my $thcaFormat  = $wb->getFormat('caption');
    my $format1     = $wb->getFormat('0copy');
    my $format2     = $wb->getFormat('0softpm');
    my $format3     = $wb->getFormat('%softpm');

    my @books = $self->listModels;
    my $findRow =
      $self->prepare(
        'select row from data where bid=? and tab=? and col=0 and v=?');
    my $q =
      $self->prepare(
        'select v from data where bid=? and tab=? and row=? and col=?');

    foreach my $i ( 0 .. $#{ $options{sheetNames} } ) {
        my $qr = $options{sheetNames}[$i];
        $qr =~ tr/ /-/;
        my ($bidb) =
          grep { $_->[1] =~ /$qr/ && $options{basematch}->( $_->[1] ) } @books;
        next unless $bidb;
        $bidb = $bidb->[0];
        my ($bida) =
          grep { $_->[1] =~ /$qr/ && $options{dcpmatch}->( $_->[1] ) } @books;
        next unless $bida;
        $bida = $bida->[0];

        unless ( $options{linesAfter} ) {
            $linesBefore = $linesAfter = [
                map { @$_ } @{
                    $self->selectall_arrayref(
                        'select v from data where bid=? and tab=?'
                          . ' and col=0 and row>0',
                        undef, $bida, $options{tableNumber}
                    )
                }
            ];
        }

        my $ws = $wb->add_worksheet( $options{sheetNames}[$i] );
        $ws->set_column( 0, 0,   48 );
        $ws->set_column( 1, 254, 16 );
        $ws->hide_gridlines(2);
        $ws->freeze_panes( 1, 1 );
        $ws->write_string( 0, 0, $options{sheetTitles}[$i], $titleFormat );

        $ws->write_string( 2, 1, 'Baseline revenue (£/year)',     $thcFormat );
        $ws->write_string( 2, 2, 'Revenue on new basis (£/year)', $thcFormat );
        $ws->write_string( 2, 3, 'Change (£/year)',               $thcFormat );
        $ws->write_string( 2, 4, 'Percentage change',             $thcFormat );

        $ws->write_string( 4 + @$linesAfter,
            0, 'Baseline revenue matrix (£/year)', $thcaFormat );
        $ws->write_string( 7 + 2 * @$linesAfter,
            0, 'Revenue matrix on new basis (£/year)', $thcaFormat );
        $ws->write_string( 10 + 3 * @$linesAfter,
            0, 'Change (£/year)', $thcaFormat );
        $ws->write_string( 13 + 4 * @$linesAfter,
            0, 'Percentage change', $thcaFormat );

        my @list = map { local $_ = $_->[0]; s/[\r\n]+/\n/g; $_; } @{
            $self->selectall_arrayref(
                'select v from data where bid=? and tab='
                  . $options{tableNumber}
                  . ' and row=0 and col>'
                  . $options{col1}
                  . ' and col<'
                  . $options{col2}
                  . ' order by col',
                undef, $bidb
            )
        };

        foreach my $r (
            4 + @$linesAfter,
            7 + 2 * @$linesAfter,
            10 + 3 * @$linesAfter,
            13 + 4 * @$linesAfter
          )
        {
            for ( my $k = 0 ; $k < @list ; ++$k ) {
                $ws->write_string( $r + 1, 1 + $k, $list[$k], $thcFormat );
            }
        }

        use Spreadsheet::WriteExcel::Utility;
        my $diff = $ws->store_formula('=IV2-IV1');
        my $perc = $ws->store_formula('=IF(IV1,IV3/IV2-1,0)');

        for ( my $j = 0 ; $j < @$linesAfter ; ++$j ) {
            $findRow->execute( $bidb, $options{tableNumber},
                $linesBefore->[$j] );
            my ($rowb) = $findRow->fetchrow_array;
            $findRow->execute( $bida, $options{tableNumber},
                $linesAfter->[$j] );
            my ($rowa) = $findRow->fetchrow_array;
            my $tariffName = $options{nameExtraColumn}
              ? eval {
                $q->execute(
                    $bida, $options{tableNumber},
                    $rowa, $options{nameExtraColumn}
                );
                my ($x) = $q->fetchrow_array;
                $x =~ s/\[.*\]//;
                "Tariff $linesAfter->[$j]: $x";
              }
              : $linesAfter->[$j];
            foreach my $r (
                3,
                6 + @$linesAfter,
                9 + 2 * @$linesAfter,
                12 + 3 * @$linesAfter,
                15 + 4 * @$linesAfter
              )
            {
                $ws->write( $r + $j, 0, $tariffName, $thFormat );
            }
            my $tota = 0;
            my $totb = 0;
            for ( my $k = 1 ; $k < $options{col2} - $options{col1} ; ++$k ) {
                $q->execute(
                    $bidb, $options{tableNumber},
                    $rowb, $k + $options{col1}
                );
                my ($vb) = $q->fetchrow_array;
                $q->execute(
                    $bida, $options{tableNumber},
                    $rowa, $k + $options{col1}
                );
                my ($va) = $q->fetchrow_array;
                $ws->write( 6 + @$linesAfter + $j,     $k, $vb, $format1 );
                $ws->write( 9 + 2 * @$linesAfter + $j, $k, $va, $format1 );
                my $old = xl_rowcol_to_cell( 6 + @$linesAfter + $j,     $k );
                my $new = xl_rowcol_to_cell( 9 + 2 * @$linesAfter + $j, $k );
                $ws->repeat_formula(
                    12 + 3 * @$linesAfter + $j, $k, $diff, $format2,
                    IV1 => $old,
                    IV2 => $new,
                );
                $ws->repeat_formula(
                    15 + 4 * @$linesAfter + $j, $k, $perc, $format3,
                    IV1 => $old,
                    IV2 => $old,
                    IV3 => $new,
                );

                if ( grep { $k + $options{col1} == $_ } @{ $options{columns} } )
                {
                    $tota += $va if defined $va;
                    $totb += $vb if defined $vb;
                }
            }

            $ws->write( 3 + $j, 1, $totb, $format1 );
            $ws->write( 3 + $j, 2, $tota, $format1 );
            my $old = xl_rowcol_to_cell( 3 + $j, 1 );
            my $new = xl_rowcol_to_cell( 3 + $j, 2 );
            $ws->repeat_formula(
                3 + $j, 3, $diff, $format2,
                IV1 => $old,
                IV2 => $new,
            );
            $ws->repeat_formula(
                3 + $j, 4, $perc, $format3,
                IV1 => $old,
                IV2 => $old,
                IV3 => $new,
            );
        }

    }

}

sub _defaultOptions {
    my ($or) = @_;
    $or->{dcpName}    ||= 'DCP';
    $or->{basematch}  ||= sub { $_[0] !~ /DCP/i };
    $or->{dcpmatch}   ||= sub { $_[0] =~ /DCP/i };
    $or->{sheetNames} ||= [ split /\n/, <<EOL ];
ENWL
NPG Northeast
NPG Yorkshire
SPEN SPD
SPEN SPM
SSEPD SEPD
SSEPD SHEPD
UKPN EPN
UKPN LPN
UKPN SPN
WPD EastM
WPD SWales
WPD SWest
WPD WestM
SimplePower
DuckPower
GoosePower
TestPowerF
TestPowerL
EOL
    $or->{sheetTitles} ||=
      [ map { "$_: illustrative impact of $or->{dcpName}" } split /\n/, <<EOL ];
Electricity North West
Northern Powergrid Northeast
Northern Powergrid Yorkshire
SP Distribution
SP Manweb
SEPD
SHEPD
Eastern Power Networks
London Power Networks
South Eastern Power Networks
WPD East Midlands
WPD South Wales
WPD South West
WPD West Midlands
Simple Power Networks
Duck Power Networks
Goose Power Networks
Test Power Networks FCP
Test Power Networks LRIC
EOL
}

sub cdcmUserImpact {

    my ( $self, $wbmodule, %options ) = @_;

    _defaultOptions( \%options );

    my $wb = $wbmodule->new(
        "Impact users $options{dcpName}" . $wbmodule->fileExtension );
    $wb->setFormats(
        {
            $options{colour} ? ( colour => $options{colour} ) : (),
        }
    );

    my $linesAfter = $options{linesAfter} || [ split /\n/, <<EOL ];
Average home (Domestic Unrestricted)
Average home (LDNO LV: Domestic Unrestricted)
Average home (LDNO HV: Domestic Unrestricted)
Average home (Margin LV: Domestic Unrestricted)
Average home (Margin HV: Domestic Unrestricted)
Average home x250 (LV HH Metered)
Average home x250 (LV Sub HH Metered)
Average home x250 (HV HH Metered)
Average home x250 (LDNO LV: LV HH Metered)
Average home x250 (LDNO HV: LV HH Metered)
Average home x250 (LDNO HV: LV Sub HH Metered)
Average home x250 (LDNO HV: HV HH Metered)
Average home x2500 (HV HH Metered)
Electric heating home (Domestic Unrestricted)
Electric heating home (Domestic Two Rate)
Electric heating home (LDNO LV: Domestic Unrestricted)
Electric heating home (LDNO LV: Domestic Two Rate)
Electric heating home (LDNO HV: Domestic Unrestricted)
Electric heating home (LDNO HV: Domestic Two Rate)
Electric heating home (Margin LV: Domestic Unrestricted)
Electric heating home (Margin LV: Domestic Two Rate)
Electric heating home (Margin HV: Domestic Unrestricted)
Electric heating home (Margin HV: Domestic Two Rate)
Electric heating home x100 (LV HH Metered)
Electric heating home x100 (LV Sub HH Metered)
Electric heating home x100 (HV HH Metered)
Electric heating home x100 (LDNO LV: LV HH Metered)
Electric heating home x100 (LDNO HV: LV HH Metered)
Electric heating home x100 (LDNO HV: LV Sub HH Metered)
Electric heating home x100 (LDNO HV: HV HH Metered)
Electric heating home x1000 (HV HH Metered)
Low use home (Domestic Unrestricted)
Low use home (LDNO LV: Domestic Unrestricted)
Low use home (LDNO HV: Domestic Unrestricted)
Low use home (Margin LV: Domestic Unrestricted)
Low use home (Margin HV: Domestic Unrestricted)
68kVA business (Small Non Domestic Unrestricted)
68kVA business (Small Non Domestic Two Rate)
68kVA business (LV Medium Non-Domestic)
68kVA business (LV Sub Medium Non-Domestic)
68kVA business (LV HH Metered)
68kVA business (LV Sub HH Metered)
68kVA business (LDNO LV: Small Non Domestic Unrestricted)
68kVA business (LDNO LV: Small Non Domestic Two Rate)
68kVA business (LDNO LV: LV Medium Non-Domestic)
68kVA business (LDNO LV: LV HH Metered)
68kVA business (LDNO HV: Small Non Domestic Unrestricted)
68kVA business (LDNO HV: Small Non Domestic Two Rate)
68kVA business (LDNO HV: LV Medium Non-Domestic)
68kVA business (LDNO HV: LV HH Metered)
68kVA business (LDNO HV: LV Sub HH Metered)
68kVA business (Margin LV: Small Non Domestic Unrestricted)
68kVA business (Margin LV: Small Non Domestic Two Rate)
68kVA business (Margin LV: LV Medium Non-Domestic)
68kVA business (Margin LV: LV HH Metered)
68kVA business (Margin HV: Small Non Domestic Unrestricted)
68kVA business (Margin HV: Small Non Domestic Two Rate)
68kVA business (Margin HV: LV Medium Non-Domestic)
68kVA business (Margin HV: LV HH Metered)
68kVA business (Margin HV: LV Sub HH Metered)
68kVA continuous (Small Non Domestic Unrestricted)
68kVA continuous (Small Non Domestic Two Rate)
68kVA continuous (LV Medium Non-Domestic)
68kVA continuous (LV Sub Medium Non-Domestic)
68kVA continuous (LV HH Metered)
68kVA continuous (LV Sub HH Metered)
68kVA off-peak (Small Non Domestic Unrestricted)
68kVA off-peak (Small Non Domestic Two Rate)
68kVA off-peak (LV Medium Non-Domestic)
68kVA off-peak (LV Sub Medium Non-Domestic)
68kVA off-peak (LV HH Metered)
68kVA off-peak (LV Sub HH Metered)
68kVA random (Small Non Domestic Unrestricted)
68kVA random (Small Non Domestic Two Rate)
68kVA random (LV Medium Non-Domestic)
68kVA random (LV Sub Medium Non-Domestic)
68kVA random (LV HH Metered)
68kVA random (LV Sub HH Metered)
500kVA business (LV HH Metered)
500kVA business (LV Sub HH Metered)
500kVA business (HV HH Metered)
500kVA business (LDNO LV: LV HH Metered)
500kVA business (LDNO HV: LV HH Metered)
500kVA business (LDNO HV: LV Sub HH Metered)
500kVA business (LDNO HV: HV HH Metered)
500kVA continuous (LV HH Metered)
500kVA continuous (LV Sub HH Metered)
500kVA continuous (HV HH Metered)
500kVA continuous (LDNO LV: LV HH Metered)
500kVA continuous (LDNO HV: LV HH Metered)
500kVA continuous (LDNO HV: LV Sub HH Metered)
500kVA continuous (LDNO HV: HV HH Metered)
500kVA off-peak (LV HH Metered)
500kVA off-peak (LV Sub HH Metered)
500kVA off-peak (HV HH Metered)
500kVA off-peak (LDNO LV: LV HH Metered)
500kVA off-peak (LDNO HV: LV HH Metered)
500kVA off-peak (LDNO HV: LV Sub HH Metered)
500kVA off-peak (LDNO HV: HV HH Metered)
500kVA random (LV HH Metered)
500kVA random (LV Sub HH Metered)
500kVA random (HV HH Metered)
500kVA random (LDNO LV: LV HH Metered)
500kVA random (LDNO HV: LV HH Metered)
500kVA random (LDNO HV: LV Sub HH Metered)
500kVA random (LDNO HV: HV HH Metered)
5MVA business (HV HH Metered)
5MVA continuous (HV HH Metered)
5MVA off-peak (HV HH Metered)
5MVA random (HV HH Metered)
EOL

    my $linesBefore = $options{linesBefore} || $linesAfter;

    my $titleFormat   = $wb->getFormat('notes');
    my $thFormat      = $wb->getFormat('th');
    my $thcFormat     = $wb->getFormat('thc');
    my $thcaFormat    = $wb->getFormat('caption');
    my $scalingFactor = $options{MWh} ? 1 : 0.1;
    my $ppuFormatCore = $options{MWh} ? '0.00' : '0.000';
    my @format1 =
      map { $wb->getFormat($_); } '0copy', $ppuFormatCore . 'copy';
    my @format2 =
      map { $wb->getFormat($_); } '0softpm', $ppuFormatCore . 'softpm';
    my @format3 =
      map { $wb->getFormat($_); } '%softpm';

    my @books   = $self->listModels;
    my $findRow = $self->prepare(
        'select row from data where bid=? and tab=4003 and col=0 and v=?');
    my $q = $self->prepare(
        'select v from data where bid=? and tab=4003 and row=? and col=?');

    foreach my $i ( 0 .. $#{ $options{sheetNames} } ) {
        my $qr = $options{sheetNames}[$i];
        $qr =~ tr/ /-/;
        my ($bidb) =
          grep { $_->[1] =~ /$qr/ && $options{basematch}->( $_->[1] ) } @books;
        next unless $bidb;
        $bidb = $bidb->[0];
        my ($bida) =
          grep { $_->[1] =~ /$qr/ && $options{dcpmatch}->( $_->[1] ) } @books;
        next unless $bida;
        $bida = $bida->[0];
        my $ws = $wb->add_worksheet( $options{sheetNames}[$i] );
        $ws->set_column( 0, 0,   48 );
        $ws->set_column( 1, 254, 16 );
        $ws->hide_gridlines(2);
        $ws->freeze_panes( 1, 1 );
        $ws->write_string( 0, 0, $options{sheetTitles}[$i], $titleFormat );

        $ws->write_string( 2, 1, 'Baseline £/year',     $thcFormat );
        $ws->write_string( 2, 2, 'Baseline p/kWh',      $thcFormat );
        $ws->write_string( 2, 3, '£/year on new basis', $thcFormat );
        $ws->write_string( 2, 4, 'p/kWh on new basis',  $thcFormat );
        $ws->write_string( 2, 5, 'Change (£/year)',     $thcFormat );
        $ws->write_string( 2, 6, 'Change (p/kWh)',      $thcFormat );
        $ws->write_string( 2, 7, 'Percentage change',   $thcFormat );

        use Spreadsheet::WriteExcel::Utility;
        my $diff = $ws->store_formula('=IV2-IV1');
        my $perc = $ws->store_formula('=IF(IV1,IV3/IV2-1,0)');

        for ( my $j = 0 ; $j < @$linesAfter ; ++$j ) {
            $ws->write_string( 3 + $j, 0, $linesAfter->[$j], $thFormat );
            $findRow->execute( $bidb, $linesBefore->[$j] );
            my ($rowb) = $findRow->fetchrow_array;
            $findRow->execute( $bida, $linesAfter->[$j] );
            my ($rowa) = $findRow->fetchrow_array;
            {
                $q->execute( $bidb, $rowb, 1 );
                my ($vb) = $q->fetchrow_array;
                $q->execute( $bida, $rowa, 1 );
                my ($va) = $q->fetchrow_array;
                $ws->write( 3 + $j, 1, $vb, $format1[0] );
                $ws->write( 3 + $j, 3, $va, $format1[0] );
                my $old = xl_rowcol_to_cell( 3 + $j, 1 );
                my $new = xl_rowcol_to_cell( 3 + $j, 3 );
                $ws->repeat_formula(
                    3 + $j, 5, $diff, $format2[0],
                    IV1 => $old,
                    IV2 => $new,
                );
                $ws->repeat_formula(
                    3 + $j, 7, $perc, $format3[0],
                    IV1 => $old,
                    IV2 => $old,
                    IV3 => $new,
                );
            }
            {
                $q->execute( $bidb, $rowb, 2 );
                my ($vb) = $q->fetchrow_array;
                $q->execute( $bida, $rowa, 2 );
                my ($va) = $q->fetchrow_array;
                $ws->write( 3 + $j, 2, $vb * $scalingFactor, $format1[1] );
                $ws->write( 3 + $j, 4, $va * $scalingFactor, $format1[1] );
                my $old = xl_rowcol_to_cell( 3 + $j, 2 );
                my $new = xl_rowcol_to_cell( 3 + $j, 4 );
                $ws->repeat_formula(
                    3 + $j, 6, $diff, $format2[1],
                    IV1 => $old,
                    IV2 => $new,
                );
            }
        }
    }
}

1;
