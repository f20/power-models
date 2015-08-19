package SpreadsheetModel::Reshape;

=head Copyright licence and disclaimer

Copyright 2008-2015 Franck Latrémolière, Reckon LLP and others.

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
    'Reshape table';
}

sub populateCore {
    my ($self) = @_;
    $self->{core}{$_} = $self->{$_}->getCore foreach qw(source);
}

sub check {
    my ($self) = @_;
    return 'No source to reshape' unless $self->{source};
    return 'Reshape does not like the job'
      unless $self->{cols}{groups} == $self->{source}{rows}{list}
      and !grep { $self->{source}{cols}{list} != $_->{list} }
      @{ $self->{source}{rows}{list} };
    return 'Reshape cannot do rows'
      if $self->{rows};
    push @{ $self->{sourceLines} }, $self->{source};
    $self->{arithmetic} = '= A1';
    $self->{arguments} = { A1 => $self->{source} };
    $self->SUPER::check;
}

sub wsPrepare {
    my ( $self, $wb, $ws ) = @_;
    my $wsWorkings = $ws->{workingsSheet} || $ws;
    my $provisionallyBroken;

    my ( $srcsheet, $srcr, $srcc ) =
      $self->{source}->wsWrite( $wb, $wsWorkings );
    return
      sub { die "Unfeasible link to source for $self->{name} $self->{debug}"; }
      unless $srcsheet;
    $srcsheet =
      $srcsheet == $ws
      ? ''
      : "'" . $srcsheet->get_name . "'!";

    my $formula = $ws->store_formula("=${srcsheet}A1");
    my $format = $wb->getFormat( $self->{defaultFormat} || '0.000copy' );

    my ( @x, @y );

    my $row = 0;
    my $col = 0;
    foreach ( @{ $self->{cols}{groups} } ) {
        my @i = $self->{source}{cols}->indices;
        push @x, undef, @i;
        push @y, undef, map { $row } @i;
        ++$row;
    }

    sub {
        '', $format, $formula, qr/\bA1\b/ =>
          xl_rowcol_to_cell( $srcr + $y[ $_[0] ], $srcc + $x[ $_[0] ], 1, 0 );
    };
}

1;
