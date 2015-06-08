package SpreadsheetModel::Arithmetic;

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

require SpreadsheetModel::Dataset;
our @ISA = qw(SpreadsheetModel::Dataset);

use Spreadsheet::WriteExcel::Utility;

sub objectType {
    $_[0]{arithmetic} =~ /^=\s*[A-Z]+[0-9]+$/
      ? 'Copy cells'
      : 'Calculation';
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

    my $lead = $self->{arguments}{IV1};
    $self->{rows} = $lead->{rows} unless defined $self->{rows};
    $self->{cols} = $lead->{cols} unless defined $self->{cols};

    push @{ $self->{sourceLines} }, values %{ $self->{arguments} };

    while ( my ( $ph, $arg ) = each %{ $self->{arguments} } ) {

        0
          and warn
"$self->{name} $self->{debug} $ph $arg->{name} $self->{rows} $self->{cols} $arg->{rows} $arg->{cols}";

        next if $ph =~ /_/;

        local $_ = $arg;

        return <<EOM
Mismatch (Arithmetic):
$_->{name}: $_->{rows} x $_->{cols}
 - v -
$self->{name}: $self->{rows} x $self->{cols}
EOM
          if (  $_->{rows}
            and $_->{rows} != $self->{rows}
            and $#{ $_->{rows}{list} }
            and !$self->{rows}
            || !$self->{rows}{groups}
            || $_->{rows}{list} != $self->{rows}{groups}
            and !$self->{rows}
            || !defined $self->{rows}->supersetIndex( $_->{rows} )
            or $_->{cols}
            and $#{ $_->{cols}{list} }
            and $_->{cols} != $self->{cols}
            and !$self->{cols}
            || !$self->{cols}{groups}
            || $_->{cols}{list} != $self->{cols}{groups}
            and !$self->{cols}
            || !defined $self->{cols}->supersetIndex( $_->{cols} ) )
          and (
               $_->{rows}
            && ( !$self->{cols} || $_->{rows} != $self->{cols} )
            && !(
                   $self->{cols}
                && $self->{cols}{accepts}
                && $self->{cols}{accepts}[0] == $_->{rows}
                && $#{ $self->{cols}{list} } == $#{ $_->{rows}{list} }
            )
            || $_->{cols} && $_->{cols} != $self->{rows}
          );
    }

=head Development note

There is a horrible hack when accepts and transposition interact.  That broke something once.

=cut

    $self->SUPER::check;
}

sub wsPrepare {

    my ( $self, $wb, $ws ) = @_;

    my $arithmetic = $self->{arithmetic};
    my $volatile;
    my $provisionallyBroken;

    my @placeholders = sort keys %{ $self->{arguments} };

    my ( %row, %col );
    for my $ph (@placeholders) {
        0 and warn "$self->{name} $self->{debug} $ph $self->{arguments}{$ph}";
        die "$self->{name} $self->{debug} $ph is undefined"
          unless defined $self->{arguments}{$ph};
        ( my $ws2, $row{$ph}, $col{$ph} ) =
          $self->{arguments}{$ph}->wsWrite( $wb, $ws );
        if ( !$ws2 ) {
            $provisionallyBroken =
              "UNFEASIBLE LINK: $ph in $self->{name} $self->{debug}";
        }
        if ( $ws2 && $ws2 != $ws ) {
            my $sheet = $ws2->get_name;
            use bytes;
            $arithmetic =~ s/\b$ph\b/'$sheet'!$ph/;
        }
        if ( my ( $a, $b ) = ( $ph =~ /^([A-Z0-9]+)_([A-Z0-9]+)$/ ) ) {
            $arithmetic =~ s/\b$ph\b/$a:$b/;
            $volatile = 1;
        }
    }

    return sub { die $provisionallyBroken; }
      if $provisionallyBroken;

    $volatile = 1 if $arithmetic =~ /\bM(IN|AX)\b/;

    my $formula = $ws->store_formula($arithmetic);
    if ($volatile) {
        s/_ref2d/_ref2dV/ foreach @$formula;
        s/_ref3d/_ref3dV/ foreach @$formula;
    }

    my $format = $wb->getFormat( $self->{defaultFormat} || '0.000soft' );

    my @stdph = grep { !/_/ } @placeholders;

    my %modx;
    @modx{@stdph} = map {
        my $c = $self->{arguments}{$_}{cols};
        $c == $self->{cols} ? 0
          : !$c || !$#{ $c->{list} } ? 1
          : $c == $self->{rows} ? 2
          : $self->{cols}{groups} && $c->{list} == $self->{cols}{groups} ? 3
          :   $self->{cols}->supersetIndex($c);
    } @stdph;

    my %mody;
    @mody{@stdph} = map {
        my $c = $self->{arguments}{$_}{rows};
        $c == $self->{rows} ? 0
          : !$c || !$#{ $c->{list} } ? 1
          : $c == $self->{cols} ? 2
          : $self->{cols}
          && $self->{cols}{accepts} && $self->{cols}{accepts}[0] == $c     ? 2
          : $self->{rows}{groups}   && $c->{list} == $self->{rows}{groups} ? 3
          :   $self->{rows}->supersetIndex($c);
    } @stdph;

    sub {
        my ( $x, $y ) = @_;
        return '', $wb->getFormat('unavailable')
          if $self->{rowFormats}
          && $self->{rowFormats}[$y]
          && $self->{rowFormats}[$y] eq 'unavailable';
        '',
          $self->{rowFormats} && $self->{rowFormats}[$y]
          ? $wb->getFormat( $self->{rowFormats}[$y] )
          : $format, $formula, map {
            if ( my ( $a, $b ) = (/^([A-Z0-9]+)_([A-Z0-9]+)$/) ) {
                my $arg = $self->{arguments}{$_};
                qr/\b$a\b/   => xl_rowcol_to_cell( $row{$_}, $col{$_}, 1, 1 ),
                  qr/\b$b\b/ => xl_rowcol_to_cell(
                    $row{$_} + $arg->lastRow,
                    $col{$_} + $arg->lastCol,
                    1, 1
                  );
            }
            else {
                qr/\b$_\b/ => xl_rowcol_to_cell(
                    $row{$_} + (
                        ref $mody{$_}
                        ? (
                            $mody{$_}[$y] < 0
                            ? -1 - $mody{$_}[$y]
                            : $mody{$_}[$y]
                          )
                        : $mody{$_} == 0 ? $y
                        : $mody{$_} == 1 ? 0
                        : $mody{$_} == 2 ? $x
                        : $mody{$_} == 3 ? $self->{rows}{groupid}[$y]
                        : $mody{$_} < 0  ? $y % -$mody{$_}
                        :                  die
                    ),
                    $col{$_} + (
                        ref $modx{$_}
                        ? (
                            $modx{$_}[$x] < 0
                            ? -1 - $modx{$_}[$x]
                            : $modx{$_}[$x]
                          )
                        : $modx{$_} == 0 ? $x
                        : $modx{$_} == 1 ? 0
                        : $modx{$_} == 2 ? $y
                        : $modx{$_} == 3 ? $self->{cols}{groupid}[$x]
                        : $modx{$_} < 0  ? $x % -$modx{$_}
                        :                  die
                    ),
                    ref $mody{$_} ? $mody{$_}[$y] >= 0 : $mody{$_} > 0,
                    ref $modx{$_} ? $modx{$_}[$x] >= 0 : $modx{$_} > 0
                );
            }
          } @placeholders;
    };
}

1;
