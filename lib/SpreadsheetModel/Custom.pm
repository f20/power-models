package SpreadsheetModel::Custom;

=head Copyright licence and disclaimer

Copyright 2008-2013 Franck Latrémolière, Reckon LLP and others.

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

use SpreadsheetModel::Dataset;
our @ISA = qw(SpreadsheetModel::Dataset);

use Spreadsheet::WriteExcel::Utility;

sub objectType {
    $_[0]{objectType};
}

sub populateCore {
    my ($self) = @_;
    $self->{core}{$_} = $self->{$_}
      foreach grep { exists $self->{$_}; } qw(arithmetic);
    while ( my ( $k, $v ) = each %{ $self->{arguments} } ) {
        $self->{core}{arguments}{$k} = $v->getCore;
    }
}

sub check {
    my ($self) = @_;
    $self->{objectType} ||= 'Special calculation';
    $self->{arithmetic} = join ' or ', map {
        local $_ = $_;
        s/([A-Z]+[0-9]+):([A-Z]+[0-9]+)/${1}_${2}/g;
        $_;
      } @{ $self->{custom} }
      unless defined $self->{arithmetic};
    push @{ $self->{sourceLines} }, values %{ $self->{arguments} };
    $self->{wsPrepare} ||= sub {
        my ( $self, $wb, $ws, $formula, $format, $pha, $rowh, $colh ) = @_;
        sub {
            my ( $x, $y ) = @_;
            '', $format, $formula, map {
                $_ => Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{$_} + $y,
                    $colh->{$_} + $x,
                    0, 0
                  )
            } @$pha;
        };
    };
    $self->SUPER::check;
}

sub wsPrepare {
    my ( $self, $wb, $ws ) = @_;
    my @custom       = @{ $self->{custom} };
    my @placeholders = keys %{ $self->{arguments} };
    my ( %row, %col );
    my $provisionallyBroken;
    for my $ph (@placeholders) {
        0 and warn "$self->{name} $ph";
        ( my $ws2, $row{$ph}, $col{$ph} ) =
          $self->{arguments}{$ph}->wsWrite( $wb, $ws );
        $provisionallyBroken =
          "UNFEASIBLE LINK $ph for $self->{name} $self->{debug}"
          unless $ws2;
        unless ( !$ws2 || $ws2 == $ws ) {
            my $sheet = $ws2->get_name;
            use bytes;
            s/\b$ph\b/'$sheet'!$ph/ foreach @custom;
        }
    }
    return sub { die $provisionallyBroken; }
      if $provisionallyBroken;
    $self->{wsPrepare}->(
        $self, $wb, $ws,
        $wb->getFormat( $self->{defaultFormat} || '0.000soft' ),
        [
            map {
                my $formula = $ws->store_formula($_);
                if (/\b(?:MIN|MAX|AVERAGE|INDEX|MATCH)\b/) {
                    s/_ref2d/_ref2dV/ foreach @$formula;
                    s/_ref3d/_ref3dV/ foreach @$formula;
                }
                $formula;
            } @custom
        ],
        \@placeholders,
        \%row,
        \%col
    );
}

1;
