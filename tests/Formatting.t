#!/usr/bin/env perl
use warnings;
use strict;
use utf8;

use File::Spec::Functions qw(rel2abs catdir);
use File::Basename 'dirname';
my $homedir;

BEGIN {
    $homedir =
      dirname dirname( rel2abs( -l $0 ? ( readlink $0, dirname $0) : $0 ) );
}
use lib map { catdir( $homedir, $_ ); } qw(cpan perl5);

require SpreadsheetModel::Workbook;
my $workbookModule = 'SpreadsheetModel::Workbook';
my $fileExtension  = '.xls';
if ( grep { /^-+xlsx/i } @ARGV ) {
    $workbookModule = 'SpreadsheetModel::WorkbookXLSX';
    $fileExtension .= 'x';
    require SpreadsheetModel::WorkbookXLSX;
}

sub newTestArea {
    my ( $wbook, $wsheet, @formats ) = @_;
    $wbook ||= 'Test workbook' . $fileExtension;
    unless ( UNIVERSAL::can( $wbook, 'add_worksheet' ) ) {
        $wbook = $workbookModule->new($wbook);
        $wbook->setFormats(@formats);
    }
    unless ( UNIVERSAL::can( $wsheet, 'repeat_formula' ) ) {
        $wsheet = $wbook->add_worksheet($wsheet);
    }
    $wbook, $wsheet;
}

use SpreadsheetModel::Shortcuts ':all';

use Test::More tests => 2;
ok( linkFormatExperiment( newTestArea('linkFormatExperiment.xls') ) );
ok( !eval { brokenFormatTest( newTestArea('brokenFormatTest.xls') ); } && $@ );

sub linkFormatExperiment {
    my ( $wbook, $wsheet ) = @_;
    $wsheet->write( 1, 0, 'Used in', $wbook->getFormat('text') );
    my $linkFormat = $wbook->getFormat( [ base => 'link', text_wrap => 1 ] );
    $wsheet->write_url( 1, 1, "http://whatever/", "101. Whatever this table is",
        $linkFormat );
    $wsheet->write_url( 1, 2, "http://whatever/", "102. Whatever this table is",
        $linkFormat );
    undef $wbook;
    1;
}

sub brokenFormatTest {
    my ( $wbook, $wsheet ) = @_;
    $wsheet->write( 1, 0, 'Used in', $wbook->getFormat('text') );
    my $linkFormat = $wbook->getFormat('broken');
    $wsheet->write_url( 1, 1, "http://whatever/", "101. Whatever this table is",
        $linkFormat );
    $wsheet->write_url( 1, 2, "http://whatever/", "102. Whatever this table is",
        $linkFormat );
    undef $wbook;
    1;
}

1;
