
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
use Ancillary::PowerModelTesting qw(newTestArea);

use SpreadsheetModel::Shortcuts ':all';

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

use Test::More tests => 4;
ok( test_sumif( newTestArea('test-sumif_1.xls'),  42 ) );
ok( test_sumif( newTestArea('test-sumif_1.xlsx'), 42 ) );
ok( test_sumif( newTestArea('test-sumif_2.xls'),  '"forty two"' ) );
ok( test_sumif( newTestArea('test-sumif_2.xlsx'), '"forty two"' ) );
