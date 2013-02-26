#!/usr/bin/perl

=head Copyright licence and disclaimer

Copyright 2012 Reckon LLP and others.

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
use File::Spec::Functions qw(rel2abs catdir);
use File::Basename 'dirname';
my $homedir;

BEGIN {
    $homedir = dirname( rel2abs( -l $0 ? ( readlink $0, dirname $0) : $0 ) );
    while (1) {
        last if -d catdir( $homedir, 'lib', 'SpreadsheetModel' );
        my $parent = dirname $homedir;
        last if $parent eq $homedir;
        $homedir = $parent;
    }
}
use lib map { catdir( $homedir, $_ ); } qw(cpan lib);

my $workbookModule = 'SpreadsheetModel::Workbook';
my $fileExtension  = '.xls';
require SpreadsheetModel::Workbook;

if ( require SpreadsheetModel::WorkbookXLSX ) {
    $workbookModule = 'SpreadsheetModel::WorkbookXLSX';
    $fileExtension .= 'x';
}

my $db = DBI->connect('dbi:SQLite:dbname=~$database.sqlite')
  or die "Cannot open sqlite database: $!";

summaryTariffImpact(
    \$db, $workbookModule->new("DCP130 illustrative impact$fileExtension"),
    [ split /\n/, <<EOL ],
ENW-2012-02-clean100.xlsx
NP-Northeast-2012-02-clean100.xlsx
NP-Yorkshire-2012-02-clean100.xlsx
SPEN-SPD-2012-02-clean100.xlsx
SPEN-SPM-2012-02-clean100.xlsx
SSEPD-SEPD-2012-02-clean100.xlsx
SSEPD-SHEPD-2012-02-clean100.xlsx
UKPN-EPN-2012-02-clean100.xlsx
UKPN-LPN-2012-02-clean100.xlsx
UKPN-SPN-2012-02-clean100.xlsx
WPD-EastM-2012-02-clean100.xlsx
WPD-SWales-2012-02-clean100.xlsx
WPD-SWest-2012-02-clean100.xlsx
WPD-WestM-2012-02-clean100.xlsx
EOL
    [ split /\n/, <<EOL ],
ENW4-DCP130.xlsx
NP-Northeast4-DCP130.xlsx
NP-Yorkshire4-DCP130.xlsx
SPEN-SPD4-DCP130.xlsx
SPEN-SPM4-DCP130.xlsx
SSEPD-SEPD4-DCP130.xlsx
SSEPD-SHEPD5-DCP130.xlsx
UKPN-EPN4-DCP130.xlsx
UKPN-LPN5-DCP130.xlsx
UKPN-SPN4-DCP130.xlsx
WPD-EastM4-DCP130.xlsx
WPD-SWales5-DCP130.xlsx
WPD-SWest5-DCP130.xlsx
WPD-WestM4-DCP130.xlsx
EOL
    [ split /\n/, <<EOL ],
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
HV Sub HH Metered
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
HV Sub Generation Intermittent
HV Sub Generation Non-Intermittent
LDNO LV: Domestic Unrestricted
LDNO LV: Domestic Two Rate
LDNO LV: Domestic Off Peak
LDNO LV: Small Non Domestic Unrestricted
LDNO LV: Small Non Domestic Two Rate
LDNO LV: Small Non Domestic Off Peak
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
LDNO HV: Domestic Off Peak
LDNO HV: Small Non Domestic Unrestricted
LDNO HV: Small Non Domestic Two Rate
LDNO HV: Small Non Domestic Off Peak 
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
    [ split /\n/, <<EOL ],
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
HV Sub HH Metered
NHH UMS
NHH UMS
NHH UMS
NHH UMS
LV UMS (Pseudo HH Metered)
LV Generation NHH
LV Sub Generation NHH
LV Generation Intermittent
LV Generation Non-Intermittent
LV Sub Generation Intermittent
LV Sub Generation Non-Intermittent
HV Generation Intermittent
HV Generation Non-Intermittent
HV Sub Generation Intermittent
HV Sub Generation Non-Intermittent
LDNO LV: Domestic Unrestricted
LDNO LV: Domestic Two Rate
LDNO LV: Domestic Off Peak
LDNO LV: Small Non Domestic Unrestricted
LDNO LV: Small Non Domestic Two Rate
LDNO LV: Small Non Domestic Off Peak
LDNO LV: LV Medium Non-Domestic
LDNO LV: LV HH Metered
LDNO LV: NHH UMS
LDNO LV: NHH UMS
LDNO LV: NHH UMS
LDNO LV: NHH UMS
LDNO LV: LV UMS (Pseudo HH Metered)
LDNO LV: LV Generation NHH
LDNO LV: LV Generation Intermittent
LDNO LV: LV Generation Non-Intermittent
LDNO HV: Domestic Unrestricted
LDNO HV: Domestic Two Rate
LDNO HV: Domestic Off Peak
LDNO HV: Small Non Domestic Unrestricted
LDNO HV: Small Non Domestic Two Rate
LDNO HV: Small Non Domestic Off Peak 
LDNO HV: LV Medium Non-Domestic
LDNO HV: LV HH Metered
LDNO HV: LV Sub HH Metered
LDNO HV: HV HH Metered
LDNO HV: NHH UMS
LDNO HV: NHH UMS
LDNO HV: NHH UMS
LDNO HV: NHH UMS
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
    [ split /\n/, <<EOL ],
ENW
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
    [ map { "$_: illustrative impact of DCP 130" } split /\n/, <<EOL ],
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
);

sub summaryTariffImpact {
    my ( $self, $wb, $booksBefore, $booksAfter,
        $linesAfter, $linesBefore, $sheetNames, $sheetTitles, )
      = @_;

    $wb->setFormats( { alignment => 1 } );
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

    for ( my $i = 0 ; $i < @$booksBefore ; ++$i ) {
        my $q = $db->prepare('select bid from books where filename=?');
        $q->execute( $booksBefore->[$i] );
        die unless my ($bidb) = $q->fetchrow_array;
        $q->execute( $booksAfter->[$i] );
        die unless my ($bida) = $q->fetchrow_array;
        my $findRow = $db->prepare(
            'select row from data where bid=? and tab=3701 and col=0 and v=?');
        $q = $db->prepare(
            'select v from data where bid=? and tab=3701 and row=? and col=?');
        my $ws = $wb->add_worksheet( $sheetNames->[$i] );
        $ws->set_column( 0, 0,   35 );
        $ws->set_column( 1, 254, 10 );
        $ws->hide_gridlines(2);
        $ws->freeze_panes( 4, 1 );
        $ws->write_string( 0, 0, $sheetTitles->[$i], $titleFormat );

        $ws->write_string( 2, 1,  'Current prices',      $thcaFormat );
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

                #    $va ||= 0;
                #    $vb ||= 0;

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
