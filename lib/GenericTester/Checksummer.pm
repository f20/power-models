package GenericTester::Checksummer;

# Copyright 2021 Franck Latrémolière and others.
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
use SpreadsheetModel::Checksum;
use SpreadsheetModel::Custom;

sub new {
    my ( $class, $model, @options ) = @_;
    my $component = bless { inputTables => [], }, $class;
    my @diff;
    for ( my $counter = 0 ; $counter < @options ; ++$counter ) {
        my $name   = $options[$counter]{name};
        my $rowset = Labelset( list => $options[$counter]{rows} );
        my ( @a, @b, @factors );
        foreach ( @{ $options[$counter]{columns} } ) {
            while ( my ( $k, $v ) = each %$_ ) {
                my $format = $v ? ( '0.' . ( '0' x $v ) . 'hard' ) : '0hard';
                push @factors, 10 ** $v;
                push @a,
                  Dataset(
                    name          => $k,
                    rows          => $rowset,
                    data          => [ map { ''; } $rowset->indices ],
                    defaultFormat => $format,
                  );
                push @b,
                  Dataset(
                    name          => $k,
                    rows          => $rowset,
                    data          => [ map { ''; } $rowset->indices ],
                    defaultFormat => $format,
                  );
            }
        }
        Columnset(
            name     => "$name (A)",
            columns  => \@a,
            dataset  => $model->{dataset},
            number   => 1000 + 10 * $counter + 1,
            appendTo => $component->{inputTables},
        );
        Columnset(
            name     => "$name (B)",
            columns  => \@b,
            dataset  => $model->{dataset},
            number   => 1000 + 10 * $counter + 2,
            appendTo => $component->{inputTables},
        );
        my $cksa = SpreadsheetModel::Checksum->new(
            name      => "A checksums",
            recursive => 1,
            digits    => 7,
            columns   => \@a,
            factors   => \@factors,
        );
        my $cksb = SpreadsheetModel::Checksum->new(
            name      => "B checksums",
            recursive => 1,
            digits    => 7,
            columns   => \@b,
            factors   => \@factors,
        );
        my $diff = Arithmetic(
            name          => "Agreement",
            defaultFormat => 'boolsoft',
            arithmetic    => '=A1=A2',
            arguments     => { A1 => $cksa, A2 => $cksb, },
        );
        Columnset(
            name    => "$name checksums",
            columns => [ $cksa, $cksb, $diff, ],
        );
        push @diff, $diff;
    }
    my $andFormula = '=AND(' . ( join ',', map { "A2$_"; } 0 .. $#diff ) . ')';
    $component->{flagTable} = SpreadsheetModel::Custom->new(
        name          => 'Overall agreement',
        defaultFormat => 'boolsoft',
        custom        => [$andFormula],
        arithmetic    => $andFormula,
        arguments     => { map { ( "A2$_" => $diff[$_] ); } 0 .. $#diff },
        wsPrepare     => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0], map {
                    (
                        qr/\b$_\b/ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{$_}, $colh->{$_}, 1
                          )
                    )
                } @$pha;
            };
        },
    );
    $component->{calcTables} = [ @diff, $component->{flagTable}, ];
    $component;
}

sub inputTables {
    my ($component) = @_;
    @{ $component->{inputTables} };
}

sub calcTables {
    my ($component) = @_;
    @{ $component->{calcTables} };
}

sub appendCode {
    my ( $component, $wbook, $wsheet ) = @_;
    my ( $wb, $ro, $co ) = $component->{flagTable}->wsWrite( $wbook, $wsheet );
    my $ref = q^'^
      . $wb->get_name . q^'!^
      . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co );
    qq^&IF(ISERROR($ref),"",^ . qq^IF($ref," (matching)"," (not matching)"))^;
}

1;
