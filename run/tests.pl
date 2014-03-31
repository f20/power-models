#!/usr/bin/env perl
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

use Test::More tests => 9;
ok( newTestArea('test.xls') );
ok( newTestArea('test.xlsx') );
ok( !eval { mustCrash20121201_1( newTestArea('test-mustcrash.xls') ); } && $@ );
ok( !eval { mustCrash20130223_1( newTestArea('test-mustcrash.xls') ); } && $@ );
ok( !eval { mustCrash20130223_2( newTestArea('test-mustcrash.xls') ); } && $@ );
ok( test_sumif( newTestArea('test-sumif_1.xls'),  42 ) );
ok( test_sumif( newTestArea('test-sumif_1.xlsx'), 42 ) );
ok( test_sumif( newTestArea('test-sumif_2.xls'),  '"forty two"' ) );
ok( test_sumif( newTestArea('test-sumif_2.xlsx'), '"forty two"' ) );

if ( -f 'README.md' ) {
    system 'grep ^\  README.md | while read x; do $x; done';
    system
'perl run/make.pl -onefile ModelM/Current/%-postDCP118.yml ModelM/Data-2014-02/*';
}

sub newTestArea {
    my ( $wbook, $wsheet, @formats ) = @_;
    unless ( UNIVERSAL::can( $wbook, 'add_worksheet' ) ) {
        require SpreadsheetModel::Workbook;
        my $workbookModule = 'SpreadsheetModel::Workbook';
        if ( $wbook =~ /\.xlsx$/is ) {
            $workbookModule = 'SpreadsheetModel::WorkbookXLSX';
            require SpreadsheetModel::WorkbookXLSX;
        }
        $wbook = $workbookModule->new($wbook);
        $wbook->setFormats(@formats);
    }
    unless ( UNIVERSAL::can( $wsheet, 'repeat_formula' ) ) {
        $wsheet = $wbook->add_worksheet($wsheet);
    }
    $wbook, $wsheet;
}

use SpreadsheetModel::Shortcuts ':all';

sub mustCrash20121201_1 {
    my ( $wbook, $wsheet ) = @_;
    my $c1 = Dataset( name => 'c1', data => [ [1] ] );
    my $c2 = Stack( name => 'c2', sources => [$c1] );
    my $c3 = Stack( name => 'c3', sources => [$c2] );
    Columnset( columns => [ $c1, $c3 ] )->wsWrite( $wbook, $wsheet );
}

sub mustCrash20130223_1 {
    my ( $wbook, $wsheet ) = @_;
    my $c1 = Dataset( name => 'c1', data => [ [1] ] );
    my $c2 = Dataset(
        name => 'c2',
        rows => Labelset( list => ['The row'] ),
        data => [ [1] ]
    );
    Columnset( columns => [ $c1, $c2 ] )->wsWrite( $wbook, $wsheet );
}

sub mustCrash20130223_2 {
    my ( $wbook, $wsheet ) = @_;
    my $c1 = Dataset( name => 'c1', data => [ [1] ] );
    my $c2 = Dataset(
        name => 'c2',
        rows => Labelset( list => [ 'Row A', 'Row B' ] ),
        data => [ [ 2, 3 ] ]
    );
    Columnset( columns => [ $c1, $c2 ] )->wsWrite( $wbook, $wsheet );
}

sub test_sumif {
    my ( $wbook, $wsheet, $arg ) = @_;
    $wsheet->set_column( 0, 5, 20 );
    my $rows = Labelset( list => [qw(A B C D)] );
    my $c1 = Dataset(
        name => 'c1',
        rows => $rows,
        data => [ [ 41, 42, 'forty one', 'forty two', ] ],
    );
    my $c2 = Dataset(
        name => 'c2',
        rows => $rows,
        data => [ [ 43, 44, 45, 46, ] ],
    );
    Arithmetic(
        name       => 'sumif',
        arithmetic => '=SUMIF(IV1_IV2,' . $arg . ',IV3_IV4)',
        arguments  => { IV1_IV2 => $c1, IV3_IV4 => $c2, },
    )->wsWrite( $wbook, $wsheet );
    1;
}

__END__
