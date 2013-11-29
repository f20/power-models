package Compilation;

=head Copyright licence and disclaimer

Copyright 2009-2013 Franck Latrémolière, Reckon LLP and others.

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

sub cdcmTariffImpact {

    my ( $self, $wbmodule, $fileExtension, %options ) = @_;
    my $db = $$self;

    $options{dcpName} ||= 'DCP';

    my $wb = $wbmodule->new("Tariff impact $options{dcpName}$fileExtension");
    $wb->setFormats( { colour => 'orange', alignment => 1 } );

    $options{basematch} ||= sub { $_[0] !~ /DCP/i };
    $options{dcpmatch}  ||= sub { $_[0] =~ /DCP/i };

    my $sheetNames = $options{sheetNames} || [ split /\n/, <<EOL ];
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
EOL

    my $sheetTitles = $options{sheetTitles}
      || [
        map { "$_: illustrative impact of $options{dcpName}" } split /\n/,
        <<EOL ];
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
EOL

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
    my @format1 =
      map { $wb->getFormat($_); }
      qw(0.000copy 0.000copy 0.000copy 0.00copy 0.00copy),
      [ base => '0.000copy', right => 5, right_color => 8 ];
    my @format2 =
      map { $wb->getFormat($_); }
      qw(0.000softpm 0.000softpm 0.000softpm 0.00softpm 0.00softpm),
      [ base => '0.000softpm', right => 5, right_color => 8 ];
    my @format3 =
      map { $wb->getFormat($_); } qw(%softpm %softpm %softpm %softpm %softpm),
      [ base => '%softpm', right => 5, right_color => 8 ];

    my @books = $self->listModels;

    foreach my $i ( 0 .. $#$sheetNames ) {
        my $qr = $sheetNames->[$i];
        $qr =~ tr/ /-/;
        my ($bidb) =
          grep { $_->[1] =~ /$qr/ && $options{basematch}->( $_->[1] ) } @books;
        next unless $bidb;
        $bidb = $bidb->[0];
        my ($bida) =
          grep { $_->[1] =~ /$qr/ && $options{dcpmatch}->( $_->[1] ) } @books;
        next unless $bida;
        $bida = $bida->[0];
        my $findRow = $db->prepare(
            'select row from data where bid=? and tab=3701 and col=0 and v=?');
        my $q = $db->prepare(
            'select v from data where bid=? and tab=3701 and row=? and col=?');
        my $ws = $wb->add_worksheet( $sheetNames->[$i] );
        $ws->set_column( 0, 0,   48 );
        $ws->set_column( 1, 254, 10 );
        $ws->hide_gridlines(2);
        $ws->freeze_panes( 4, 1 );
        $ws->write_string( 0, 0, $sheetTitles->[$i], $titleFormat );

        $ws->write_string( 2, 1,  'Baseline prices',     $thcaFormat );
        $ws->write_string( 2, 7,  'Prices on new basis', $thcaFormat );
        $ws->write_string( 2, 13, 'Price change',        $thcaFormat );
        $ws->write_string( 2, 19, 'Percentage change',   $thcaFormat );

        my $diff = $ws->store_formula('=IV2-IV1');
        my $perc = $ws->store_formula('=IF(IV1,IV3/IV2-1,"")');

        $ws->write( 2, $_, undef, $thcaFormat )
          foreach 2 .. 6, 8 .. 12, 14 .. 18, 20 .. 24;

        my @list = split /\n/, <<EOL;
Unit rate 1 p/kWh
Unit rate 2 p/kWh
Unit rate 3 p/kWh
Fixed charge p/MPAN/day
Capacity charge p/kVA/day
Reactive power charge p/kVArh
EOL

        for ( my $j = 1 ; $j < 20 ; $j += 6 ) {
            for ( my $k = 0 ; $k < @list ; ++$k ) {
                $ws->write_string( 3, $j + $k, $list[$k],
                    $k == 5 ? $thcFormatB : $thcFormat );
            }
        }

        for ( my $j = 0 ; $j < @$linesAfter ; ++$j ) {
            $ws->write_string( 4 + $j, 0, $linesAfter->[$j], $thFormat );
            $findRow->execute( $bidb, $linesBefore->[$j] );
            my ($rowb) = $findRow->fetchrow_array;
            $findRow->execute( $bida, $linesAfter->[$j] );
            my ($rowa) = $findRow->fetchrow_array;
            for ( my $k = 3 ; $k < 9 ; ++$k ) {
                $q->execute( $bidb, $rowb, $k );
                my ($vb) = $q->fetchrow_array;
                $q->execute( $bida, $rowa, $k );
                my ($va) = $q->fetchrow_array;

                $ws->write( 4 + $j, $k - 2, $vb, $format1[ $k - 3 ] );
                $ws->write( 4 + $j, $k + 4, $va, $format1[ $k - 3 ] );

                if (undef) {
                    $ws->write( 4 + $j, $k + 10, $va - $vb,
                        $format2[ $k - 3 ] );
                    $ws->write(
                        4 + $j, $k + 16,
                        $vb ? $va / $vb - 1 : '',
                        $format3[ $k - 3 ]
                    );
                }
                else {
                    use Spreadsheet::WriteExcel::Utility;
                    my $old = xl_rowcol_to_cell( 4 + $j, $k - 2 );
                    my $new = xl_rowcol_to_cell( 4 + $j, $k + 4 );
                    $ws->repeat_formula(
                        4 + $j, $k + 10, $diff, $format2[ $k - 3 ],
                        IV1 => $old,
                        IV2 => $new,
                    );
                    $ws->repeat_formula(
                        4 + $j, $k + 16, $perc, $format3[ $k - 3 ],
                        IV1 => $old,
                        IV2 => $old,
                        IV3 => $new,
                    );
                }
            }
        }
    }

}

sub cdcmPpuImpact {

    my ( $self, $wbmodule, $fileExtension, %options ) = @_;
    my $db = $$self;

    $options{dcpName} ||= 'DCP';

    my $wb = $wbmodule->new("PPU impact $options{dcpName}$fileExtension");
    $wb->setFormats( { colour => 'orange', alignment => 1 } );

    $options{basematch} ||= sub { $_[0] !~ /DCP/i };
    $options{dcpmatch}  ||= sub { $_[0] =~ /DCP/i };

    my $sheetNames = $options{sheetNames} || [ split /\n/, <<EOL ];
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
EOL

    my $sheetTitles = $options{sheetTitles}
      || [
        map { "$_: illustrative impact of $options{dcpName}" } split /\n/,
        <<EOL ];
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
EOL

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

    my @books = $self->listModels;

    foreach my $i ( 0 .. $#$sheetNames ) {
        my $qr = $sheetNames->[$i];
        $qr =~ tr/ /-/;
        my ($bidb) =
          grep { $_->[1] =~ /$qr/ && $options{basematch}->( $_->[1] ) } @books;
        next unless $bidb;
        $bidb = $bidb->[0];
        my ($bida) =
          grep { $_->[1] =~ /$qr/ && $options{dcpmatch}->( $_->[1] ) } @books;
        next unless $bida;
        $bida = $bida->[0];
        my $findRow = $db->prepare(
            'select row from data where bid=? and tab=3802 and col=0 and v=?');
        my $q = $db->prepare(
            'select v from data where bid=? and tab=3802 and row=? and col=?');
        my $ws = $wb->add_worksheet( $sheetNames->[$i] );
        $ws->set_column( 0, 0,   44 );
        $ws->set_column( 1, 254, 14 );
        $ws->hide_gridlines(2);
        $ws->freeze_panes( 1, 1 );
        $ws->write_string( 0, 0, $sheetTitles->[$i], $titleFormat );

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

    my ( $self, $wbmodule, $fileExtension, %options ) = @_;
    my $db = $$self;

    $options{dcpName} ||= 'DCP';

    my $wb =
      $wbmodule->new("Revenue matrix impact $options{dcpName}$fileExtension");
    $wb->setFormats( { colour => 'orange' } );

    $options{basematch} ||= sub { $_[0] !~ /DCP/i };
    $options{dcpmatch}  ||= sub { $_[0] =~ /DCP/i };

    my $sheetNames = $options{sheetNames} || [ split /\n/, <<EOL ];
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
EOL

    my $sheetTitles = $options{sheetTitles}
      || [
        map { "$_: illustrative impact of $options{dcpName}" } split /\n/,
        <<EOL ];
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
EOL

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
      map { $wb->getFormat($_); } ( map { '0copy' } 1 .. 24 );
    my @format2 =
      map { $wb->getFormat($_); } ( map { '0softpm' } 1 .. 24 );
    my @format3 =
      map { $wb->getFormat($_); } ( map { '%softpm' } 1 .. 24 );

    my @books = $self->listModels;

    foreach my $i ( 0 .. $#$sheetNames ) {
        my $qr = $sheetNames->[$i];
        $qr =~ tr/ /-/;
        my ($bidb) =
          grep { $_->[1] =~ /$qr/ && $options{basematch}->( $_->[1] ) } @books;
        next unless $bidb;
        $bidb = $bidb->[0];
        my ($bida) =
          grep { $_->[1] =~ /$qr/ && $options{dcpmatch}->( $_->[1] ) } @books;
        next unless $bida;
        $bida = $bida->[0];
        my $findRow = $db->prepare(
            'select row from data where bid=? and tab=3901 and col=0 and v=?');
        my $q = $db->prepare(
            'select v from data where bid=? and tab=3901 and row=? and col=?');
        my $ws = $wb->add_worksheet( $sheetNames->[$i] );
        $ws->set_column( 0, 0,   44 );
        $ws->set_column( 1, 254, 14 );
        $ws->hide_gridlines(2);
        $ws->freeze_panes( 1, 1 );
        $ws->write_string( 0, 0, $sheetTitles->[$i], $titleFormat );

        $ws->write_string( 2, 1, 'Baseline revenue (£/period)', $thcFormat );
        $ws->write_string( 2, 2, 'Revenue on new basis (£/period)',
            $thcFormat );
        $ws->write_string( 2, 3, 'Change (£/period)', $thcFormat );
        $ws->write_string( 2, 4, 'Percentage change', $thcFormat );

        $ws->write_string( 4 + @$linesAfter,
            0, 'Baseline revenue matrix (£/period)', $thcaFormat );
        $ws->write_string( 7 + 2 * @$linesAfter,
            0, 'Revenue matrix on new basis (£/period)', $thcaFormat );
        $ws->write_string( 10 + 3 * @$linesAfter,
            0, 'Change (£/period)', $thcaFormat );
        $ws->write_string( 13 + 4 * @$linesAfter,
            0, 'Percentage change', $thcaFormat );

        my @list = map { $_->[0] } @{
            $db->selectall_arrayref(
'select v from data where bid=? and tab=3901 and row=0 and col>0 and col<25 order by col',
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
            foreach my $r (
                3,
                6 + @$linesAfter,
                9 + 2 * @$linesAfter,
                12 + 3 * @$linesAfter,
                15 + 4 * @$linesAfter
              )
            {
                $ws->write_string( $r + $j, 0, $linesAfter->[$j], $thFormat );
            }
            $findRow->execute( $bidb, $linesBefore->[$j] );
            my ($rowb) = $findRow->fetchrow_array;
            $findRow->execute( $bida, $linesAfter->[$j] );
            my ($rowa) = $findRow->fetchrow_array;
            for ( my $k = 1 ; $k < 25 ; ++$k ) {
                $q->execute( $bidb, $rowb, $k );
                my ($vb) = $q->fetchrow_array;
                $q->execute( $bida, $rowa, $k );
                my ($va) = $q->fetchrow_array;
                $ws->write( 6 + @$linesAfter + $j, $k, $vb,
                    $format1[ $k - 1 ] );
                $ws->write( 9 + 2 * @$linesAfter + $j,
                    $k, $va, $format1[ $k - 1 ] );
                my $old = xl_rowcol_to_cell( 6 + @$linesAfter + $j,     $k );
                my $new = xl_rowcol_to_cell( 9 + 2 * @$linesAfter + $j, $k );
                $ws->repeat_formula(
                    12 + 3 * @$linesAfter + $j, $k, $diff, $format2[ $k - 1 ],
                    IV1 => $old,
                    IV2 => $new,
                );
                $ws->repeat_formula(
                    15 + 4 * @$linesAfter + $j, $k, $perc, $format3[ $k - 1 ],
                    IV1 => $old,
                    IV2 => $old,
                    IV3 => $new,
                );

                if ( $k == 24 ) {
                    $ws->write( 3 + $j, 1, $vb, $format1[ $k - 1 ] );
                    $ws->write( 3 + $j, 2, $va, $format1[ $k - 1 ] );
                    my $old = xl_rowcol_to_cell( 3 + $j, 1 );
                    my $new = xl_rowcol_to_cell( 3 + $j, 2 );
                    $ws->repeat_formula(
                        3 + $j, 3, $diff, $format2[ $k - 1 ],
                        IV1 => $old,
                        IV2 => $new,
                    );
                    $ws->repeat_formula(
                        3 + $j, 4, $perc, $format3[ $k - 1 ],
                        IV1 => $old,
                        IV2 => $old,
                        IV3 => $new,
                    );
                }

            }
        }
    }

}

1;
