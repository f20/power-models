package SpreadsheetModel::Checksum;

# Copyright 2014-2015 Franck Latrémolière, Reckon LLP and others.
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

our @ISA = qw(SpreadsheetModel::Dataset);

use Spreadsheet::WriteExcel::Utility;

sub check {
    my ($self) = @_;
    return 'No columns to checksum' unless ref $self->{columns} eq 'ARRAY';
    return 'No columns to checksum' unless @{ $self->{columns} };
    $self->{rows} = $self->{columns}[0]{rows};
    return 'Mismatched rows'
      if grep { $self->{rows} != $_->{rows} } @{ $self->{columns} };
    return 'No factors to apply' unless ref $self->{factors} eq 'ARRAY';
    return 'Inconsistent array lengths for columns and factors'
      unless @{ $self->{factors} } == @{ $self->{columns} };
    return "No modulus/generator pair for $self->{digits} digits"
      unless $self->{parameters} = [
        undef,
        [ 7,       3,       '0' ],
        [ 97,      41,      '00' ],
        [ 997,     492,     '000' ],
        [ 9973,    4922,    '0000' ],
        [ 99991,   49222,   '00000' ],
        [ 999983,  492223,  '000 000' ],
        [ 9999991, 4922236, '000 0000' ]
      ]->[ $self->{digits} ];
    $self->{arithmetic} = 'Checksum' unless defined $self->{arithmetic};
    $self->{objectType} ||= 'Checksum';
    $self->SUPER::check;
}

=head Perl code to test whether $t is a generator for prime $p

my $i = 0;
my $x = $t;
my @h = ( undef, undef );
while (1) {
    ++$i;
    $x = ( $x * $t ) % $p;
    last if exists $h[$x];
    undef $h[$x];
}
print "$t^$i mod $p = $x\n" if $i > $p - 3;

=cut

sub objectType {
    $_[0]{objectType};
}

sub populateCore {
    my ($self) = @_;
    $self->{core}{$_} = $self->{$_}
      foreach grep { exists $self->{$_}; } qw(digits factors);
    $self->{core}{columns} = [ map { $_->getCore } @{ $self->{columns} } ];
}

sub wsPrepare {
    my ( $self, $wb, $ws ) = @_;
    my $wsWorkings = $ws->{workingsSheet} || $ws;
    my ( @placeholder, @row, @col );
    my $someArgumentsAreMissingAtThisStage;
    my ( $modulus, $generator, $numFormat ) = @{ $self->{parameters} };
    my $arithmetic = '';
    if ( $self->{recursive} ) {
        $arithmetic = '+A1';
        push @placeholder, 'A1';
        push @row,         undef;
        push @col,         undef;
    }
    my @factors = @{ $self->{factors} };
    foreach ( @{ $self->{columns} } ) {
        my $factor = shift @factors;
        die 'Not implemented' if $_->lastCol;
        push @placeholder, my $ph = 'A' . ( 1 + @placeholder );
        ( my $ws2, $row[$#placeholder], $col[$#placeholder] ) =
          $_->wsWrite( $wb, $wsWorkings );
        $someArgumentsAreMissingAtThisStage =
          "Unfeasible link $ph for $self->{name} $self->{debug}"
          unless $ws2;
        if ( $ws2 && $ws2 != $ws ) {
            my $sheet = $ws2->get_name;
            $ph = "'$sheet'!$ph";
        }
        $arithmetic =
          "+MOD($generator*(ROUND($ph*$factor,0)$arithmetic),$modulus)";
    }
    return sub { die $someArgumentsAreMissingAtThisStage; }
      if $someArgumentsAreMissingAtThisStage;
    $arithmetic =~ s/^\+/=/s;
    my $formula = $ws->store_formula($arithmetic);
    my $format  = $wb->getFormat( $self->{defaultFormat}
          || [ base => '0soft', num_format => $numFormat ] );
    sub {
        my ( $x, $y ) = @_;
        unless ( defined $row[0] ) {
            $row[0] = $self->{$wb}{row} + 1;
            $col[0] = $self->{$wb}{col};
        }
        '', $format, $formula, map {
            qr/\b$placeholder[$_]\b/ =>
              xl_rowcol_to_cell( $row[$_] + $y, $col[$_] + $x, );
        } 0 .. $#placeholder;
    };
}

1;
