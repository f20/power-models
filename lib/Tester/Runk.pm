package Tester::Runk;

# Copyright 2020-2021 Franck Latrémolière and others.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice,
# this list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
# this list of conditions and the following disclaimer in the documentation
# and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
# EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
# THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Shortcuts ':all';
use SpreadsheetModel::CalcBlock;

sub new {
    my ( $class, $model, @options ) = @_;
    bless { model => $model, @options }, $class;
}

sub resultTables {
    my ($component) = @_;
    my $data = Constant(
        name => 'Vector',
        cols => Labelset( list => [qw(One Two Three)] ),
        data => [ [1], [2], [3] ],
    );
    my $scalar = Constant( name => 'Scalar', data => [ [5] ], );
    my $calc = Arithmetic(
        name       => 'Calc',
        arithmetic => '=A1^A2',
        arguments  => { A1 => $data, A2 => $scalar, },
    );
    CalcBlock(
        name  => 'Block',
        items => [ $data, $scalar, $calc, ]
    );
}

sub appendCode {
    my ( $component, $wbook, $wsheet ) = @_;
    my ( $wb, $ro, $co ) = $component->totalFruits->wsWrite( $wbook, $wsheet );
    my $ref = q^'^
      . $wb->get_name . q^'!^
      . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co );
    qq^&IF(ISERROR($ref),"",^
      . qq^" ("&TEXT($ref,"#,##0")&IF($ref<1.5," fruit)"," fruits)"))^;
}

1;
