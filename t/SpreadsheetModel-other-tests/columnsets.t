
=head Copyright licence and disclaimer

Copyright 2012-2015 Franck Latrémolière, Reckon LLP and others.

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

use strict;
use warnings;
use lib qw(cpan lib t/lib);
use SpreadsheetModel::Tests::PowerModelTesting qw(newTestArea);

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

use Test::More tests => 3;
ok( !eval { mustCrash20121201_1( newTestArea('test-mustcrash.xls') ); } && $@ );
ok( !eval { mustCrash20130223_1( newTestArea('test-mustcrash.xls') ); } && $@ );
ok( !eval { mustCrash20130223_2( newTestArea('test-mustcrash.xls') ); } && $@ );
