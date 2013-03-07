#!/usr/bin/env perl

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
if ( grep { /xlsx/i } @ARGV ) {
    $workbookModule = 'SpreadsheetModel::WorkbookXLSX';
    $fileExtension .= 'x';
    require SpreadsheetModel::WorkbookXLSX;
}
my $options = ( grep { /right/i } @ARGV ) ? { alignment => 'right' } : {};

if ( grep { /extract/i } @ARGV ) {
    require Ancillary::DatabaseExtract;
    require Extractor::DataLocator;
    my ( $dataReader, $bookTableIndexHash ) =
      Ancillary::DatabaseExtract->makeDatabaseReader;
    foreach ( sort keys %$bookTableIndexHash ) {
        my $d =
          $dataReader->( $bookTableIndexHash->{$_}{bid},
            Extractor->locate1076 );
        require YAML;
        YAML::DumpFile(
            ( (/^(.+?)-20/)[0] || $_ ) . "-1076.yml",
            {
                1076 => [
                    [],
                    [
                        '"Allowed revenue" (£/year)',
                        $d->{cdcm} -
                          $d->{passthrough} -
                          $d->{k} -
                          $d->{noncdcmded}
                    ],
                    [ '"Pass-through charges" (£/year)', $d->{passthrough} ],
                    [
                        "Adjustment for previous year's"
                          . ' under (over) recovery (£/year)',
                        $d->{k}
                    ],
                    [
                        'Revenue raised outside this model (£/year)',
                        -$d->{noncdcmded}
                    ]
                ]
            }
        );
    }
    exit 0;
}

if ( grep { /chedam/i } @ARGV ) {
    require Ancillary::DatabaseExtract;
    require Ancillary::DotDiagrams;
    require Chedam::DataLocator;
    require Chedam::Calculate;
    require Chedam::ToDot;
    my ( $dataReader, $bookTableIndexHash ) =
      Ancillary::DatabaseExtract->makeDatabaseReader;
    Ancillary::DotDiagrams::writeDotDiagrams(
        map { $_->calculate->toDot } map {
            my $filename = $_;
            map { $dataReader->( $bookTableIndexHash->{$filename}{bid}, $_ ); }
              exists $bookTableIndexHash->{$filename}{1703}
              ? (
                Chedam->locateHidamModelled($filename),
                Chedam->locateHidamAdjMMD($filename),
                Chedam->locateHidamActualCap($filename),
                Chedam->locateHidamActualMD($filename),
              )
              : (),
              exists $bookTableIndexHash->{$filename}{1017}
              ? ( Chedam->locateDrm($filename), )
              : (),
        } sort keys %$bookTableIndexHash
    );
    exit 0;
}

require Ancillary::DatabaseExport;
my $db = Ancillary::DatabaseExport->new;

if ( grep { /\btscs/i } @ARGV ) {
    $db->tscsCreateIntermediateTables unless grep { /norebuild/i } @ARGV;
    $db->tscsCreateOutputFiles( $workbookModule, $fileExtension,
        { %$options, ( ( grep { /csv/i } @ARGV ) ? 'csv' : 'wb' ) => 1 } );
    exit 0;
}

if ( grep { /\ball\b/i } @ARGV ) {
    $db->tableCompilations( $workbookModule, $fileExtension, $options,
        all => qw(. .) );
}

if ( grep { /100/i } @ARGV ) {
    $db->tableCompilations( $workbookModule, $fileExtension, $options,
        only100 => qw(100 .) );
}
elsif ( my ($num) = grep { /^[0-9]+$/ } @ARGV ) {
    $db->tableCompilations( $workbookModule, $fileExtension, $options,, 'all',
        '.', "^$num" );
}

if ( grep { /\bcsv\b/i } @ARGV ) {
    $db->csvCreate( grep { /small/i } @ARGV );
}

