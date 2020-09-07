package StarterModel::Calculator;

# Copyright 2020 Franck Latrémolière and others.
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

sub new {
    my ( $class, $options ) = @_;
    bless { options => $options }, $class;
}

sub input1 {
    my ($calc) = @_;
    $calc->{input1} ||= Dataset(
        name          => 'Number of each type',
        rows          => Labelset( list => [qw(Apples Oranges Pears)] ),
        data          => [ [ 0, 0, 0 ] ],
        dataset       => $calc->{options}{dataset},
        number        => 101,
        defaultFormat => '0hard',
        validation    => {
            validate      => 'decimal',
            criteria      => '>=',
            value         => 0,
            input_title   => 'Number of fruits:',
            input_message => 'Enter the number of fruits of this type',
            error_message => 'This must be positive or zero.',
        }
    );
}

sub result1 {
    my ($calc) = @_;
    $calc->{result1} ||= GroupBy(
        name          => 'Number of fruits',
        singleRowName => 'Total',
        defaultFormat => '0soft',
        source        => $calc->input1,
    );
}

sub inputsTables {
    my ($calc) = @_;
    $calc->input1;
}

sub resultTables {
    my ($calc) = @_;
    $calc->result1;
}

sub feedbackItems {
    my ($calc) = @_;
    $calc->{feedbackItems} ||= Stack(
        name          => 'Feedback copy of result',
        sources       => [ $calc->result1 ],
        singleRowName => 'Total',
    );
}

sub appendCode {
    my ( $calc, $wbook, $wsheet ) = @_;
    my ( $wb, $ro, $co ) = $calc->result1->wsWrite( $wbook, $wsheet );
    my $ref = q^'^
      . $wb->get_name . q^'!^
      . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co );
    qq^&IF(ISERROR($ref),"",^
      . qq^" ("&TEXT($ref,"#,##0")&IF($ref<1.5," fruit)"," fruits)"))^;
}

1;
