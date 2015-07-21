package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others.
Copyright 2012-2014 Franck Latrémolière, Reckon LLP and others.

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

sub preprocessDataset {

    my ($model) = @_;

    if ( my $d = $model->{dataset} ) {
        foreach (
            qw(numTariffs numLocations numExtraTariffs numSampleTariffs numExtraLocations)
          )
        {
            $model->{$_} = $d->{$_} if exists $d->{$_};
        }
    }

    $model->{dataset}{1100}[3]{'Company charging year data version'} =
      $model->{version}
      if $model->{version};

    if ( $model->{dataset}{1113} && $model->{revenueAdj} ) {
        my ($key) =
          grep { !/^_/ } keys %{ $model->{dataset}{1113}[4] };
        $model->{dataset}{1113}[4]{$key} += $model->{revenueAdj};
    }

    my ( $daysInYearKey, $hoursInPurpleKey );
    if ( $model->{dataset}{1113} && ref $model->{dataset}{1113}[1] eq 'HASH' ) {
        ($daysInYearKey) =
          grep { !/^_/ } keys %{ $model->{dataset}{1113}[1] };
        ($hoursInPurpleKey) =
          grep { !/^_/ } keys %{ $model->{dataset}{1113}[3] };
    }

    if ( $model->{dataset}
        && ( $model->{ldnoRev} && $model->{ldnoRev} =~ /nopop/i ) )
    {
        delete $model->{dataset}{$_} foreach qw(1181 1182 1183);
    }

    if (   $model->{method}
        && $model->{method} !~ /none/i )
    {
        if (
            my $h = $model->{dataset}{
                $model->{method} =~ /FCP/i
                ? 911
                : 913
            }[1]
          )
        {
            if ( ref $h eq 'HASH' ) {
                my $max = $model->{numLocations} || 0;
                foreach ( keys %$h ) {
                    if (   $h->{$_}
                        && $h->{$_} ne 'Not used'
                        && $h->{$_} ne '#VALUE!'
                        && $h->{$_} !~ /^\s*$/s )
                    {
                        $max = $_ if /^[0-9]+$/ && $_ > $max;
                    }
                    else {
                        $h->{$_} = 'Not used';
                    }
                }
                $model->{numLocations} =
                  (      $model->{randomise}
                      || $model->{small}
                      || defined $model->{numLocations} ? 0 : 16 ) +
                  $max;
                my $ds = $model->{dataset}{
                    $model->{method} =~ /FCP/i
                    ? 911
                    : 913
                };
                foreach my $k ( 1 .. $model->{numLocations} ) {
                    exists $ds->[$_]{$k} || ( $ds->[$_]{$k} = '' )
                      foreach 0 .. $#$ds;
                }
            }
        }
    }

    if ( my $ds = $model->{dataset}{935} ) {
        if ( ref $ds->[1] eq 'HASH' ) {
            my %tariffs;

            splice @$ds, 8, 0,
              $model->{dataset}{t935dcp189}
              ? $model->{dataset}{t935dcp189}[0]
              : {
                map { ( $_ => $model->{dcp189default} || '' ); }
                  keys %{ $ds->[1] }
              }
              if $model->{dcp189}
              and !$ds->[8]
              || !$ds->[8]{'_column'}
              || $ds->[8]{'_column'} !~ /reduction/i;

            my $max = 0;
            $ds->[1]{$_} ||= "Tariff $_" foreach grep {
                my $t = $_;
                grep {
                         defined $ds->[$_]{$t}
                      && $ds->[$_]{$t} ne 'VOID'
                      && $ds->[$_]{$t} ne '#VALUE!';
                } 2 .. 6;
            } keys %{ $ds->[2] };
            while ( my ( $k, $v ) = each %{ $ds->[1] } ) {
                next
                  unless $k =~ /^[0-9]+$/
                  && $v
                  && $v ne 'Not used'
                  && $v ne '#VALUE!'
                  && $v !~ /^\s*$/s;
                undef $tariffs{$k};
                $max = $k
                  if $k > $max;
            }

            if ($max) {
                my @tariffs;
                if (   $model->{numTariffs}
                    && $max <= $model->{numTariffs} )
                {
                    $model->{numTariffs} = 2
                      unless $model->{transparency}
                      && $model->{transparency} =~ /impact/i
                      || $model->{numTariffs} > 1;
                    @tariffs = ( 1 .. $model->{numTariffs} );
                }
                else {
                    @tariffs =
                      sort { $a <=> $b } keys %tariffs,
                      defined $model->{numTariffs}
                      ? ()
                      : ( $max + 1 .. $max + 16 );
                    push @tariffs, $max + 1
                      unless $model->{transparency}
                      && $model->{transparency} =~ /impact/i
                      || @tariffs > 1;
                    $model->{numTariffs} = @tariffs;
                }
                $model->{tariffSet} = Labelset(
                    name          => 'Tariffs',
                    list          => \@tariffs,
                    defaultFormat => 'thtar',
                );
                foreach my $k (@tariffs) {
                    my $v = $ds->[1]{$k};
                    if (    $v
                        and $v ne 'Not used'
                        and $v ne '#VALUE!'
                        and $v ne '#N/A'
                        and $v ne 'VOID'
                        and $v !~ /^\s*$/s )
                    {
                        $ds->[$_]{$k} || ( $ds->[$_]{$k} = 'VOID' )
                          foreach 2 .. 6;
                        exists $ds->[$_]{$k} || ( $ds->[$_]{$k} = '' )
                          foreach 7 .. $#$ds;
                        $_
                          && /^[0-9.]+$/s
                          && $daysInYearKey
                          && $_ > $model->{dataset}{1113}[1]{$daysInYearKey}
                          && ( $model->{dataset}{1113}[1]{$daysInYearKey} = $_ )
                          foreach $ds->[22]{$k};
                        $_
                          && /^[0-9.]+$/s
                          && $hoursInPurpleKey
                          && $_ > $model->{dataset}{1113}[3]{$hoursInPurpleKey}
                          && ( $model->{dataset}{1113}[3]{$hoursInPurpleKey} =
                            $_ )
                          foreach $ds->[23]{$k};
                    }
                    else {
                        $ds->[1]{$k} = ' ';
                        $ds->[$_]{$k} = 'VOID' foreach 2 .. 6;
                        $ds->[$_]{$k} = ''     foreach 7 .. $#$ds;
                    }
                }
            }

            if ( $model->{nonames} ) {
                $ds->[1]{$_} = "Tariff $_" foreach keys %{ $ds->[1] };
            }
        }

        else {

            my %tariffs;

            splice @$ds, 8, 0,
              {
                map { ( $_ => $model->{dcp189default} || '' ); }
                  keys %{ $ds->[1] }
              }
              if $model->{dcp189}
              and !$ds->[8] || !$ds->[8][0] || $ds->[8][0] !~ /reduction/i;

            my $max = 0;
            $ds->[1][$_] ||= "Tariff $_" foreach grep {
                my $t = $_;
                grep {
                         defined $ds->[$_][$t]
                      && $ds->[$_][$t] ne 'VOID'
                      && $ds->[$_][$t] ne '#VALUE!';
                } 2 .. 6;
            } 1 .. $#{ $ds->[2] };
            for ( my $k = 1 ; $k < @{ $ds->[1] } ; ++$k ) {
                my $v = $ds->[1][$k];
                next
                  unless $k =~ /^[0-9]+$/
                  && $v
                  && $v ne 'Not used'
                  && $v ne '#VALUE!'
                  && $v !~ /^\s*$/s;
                undef $tariffs{$k};
                $max = $k
                  if $k > $max;
            }

            if ($max) {
                my @tariffs;
                if (   $model->{numTariffs}
                    && $max <= $model->{numTariffs} )
                {
                    $model->{numTariffs} = 2
                      unless $model->{transparency}
                      && $model->{transparency} =~ /impact/i
                      || $model->{numTariffs} > 1;
                    @tariffs = ( 1 .. $model->{numTariffs} );
                }
                else {
                    @tariffs =
                      sort { $a <=> $b } keys %tariffs,
                      defined $model->{numTariffs}
                      ? ()
                      : ( $max + 1 .. $max + 16 );
                    push @tariffs, $max + 1
                      unless $model->{transparency}
                      && $model->{transparency} =~ /impact/i
                      || @tariffs > 1;
                    $model->{numTariffs} = @tariffs;
                }
                $model->{tariffSet} = Labelset(
                    name          => 'Tariffs',
                    list          => \@tariffs,
                    defaultFormat => 'thtar',
                );
                foreach my $k (@tariffs) {
                    my $v = $ds->[1][$k];
                    if (    $v
                        and $v ne 'Not used'
                        and $v ne '#VALUE!'
                        and $v ne '#N/A'
                        and $v ne 'VOID'
                        and $v !~ /^\s*$/s )
                    {
                        $ds->[$_][$k] || ( $ds->[$_][$k] = 'VOID' )
                          foreach 2 .. 6;
                        exists $ds->[$_][$k] || ( $ds->[$_][$k] = '' )
                          foreach 7 .. $#$ds;
                        $_
                          && /^[0-9.]+$/s
                          && $daysInYearKey
                          && $_ > $model->{dataset}{1113}[1]{$daysInYearKey}
                          && ( $model->{dataset}{1113}[1]{$daysInYearKey} = $_ )
                          foreach $ds->[22][$k];
                        $_
                          && /^[0-9.]+$/s
                          && $hoursInPurpleKey
                          && $_ > $model->{dataset}{1113}[3]{$hoursInPurpleKey}
                          && ( $model->{dataset}{1113}[3]{$hoursInPurpleKey} =
                            $_ )
                          foreach $ds->[23][$k];
                    }
                    else {
                        $ds->[1][$k] = ' ';
                        $ds->[$_][$k] = 'VOID' foreach 2 .. 6;
                        $ds->[$_][$k] = ''     foreach 7 .. $#$ds;
                    }
                }
            }

            if ( $model->{nonames} ) {
                $ds->[1][$_] = "Tariff $_" foreach 1 .. $#{ $ds->[1] };
            }
        }
    }

    if ( my @tables = grep { $_ } @{ $model->{dataset} }{qw(1133 1134 1136)} ) {
        if ( $model->{tableGrouping} || $model->{transparency} ) {
            foreach ( grep { !$_->[1]{_column} || $_->[1]{_column} =~ /GSP/ }
                @tables )
            {
                splice @$_, 1, 1;
            }
        }
        else {
            foreach ( grep { $_->[1]{_column} && $_->[1]{_column} !~ /GSP/ }
                @tables )
            {
                splice @$_, 1, 0, $_->[1];
                foreach my $k ( keys %{ $_->[1] } ) {
                    $_->[1]{$k} = $k eq '_column' ? 'GSP' : '';
                }
            }
        }
        if ( $model->{dataset}{1136} ) {
            $model->{dataset}{1133} ||= [
                map {
                    my %a;
                    while ( my ( $k, $v ) = each %$_ ) {
                        $a{$k} = $v if $k eq '_column' || $k =~ /maximum/i;
                    }
                    \%a;
                } @{ $model->{dataset}{1136} }
            ];
            $model->{dataset}{1134} ||= [
                map {
                    my %a;
                    while ( my ( $k, $v ) = each %$_ ) {
                        $a{$k} = $v if $k eq '_column' || $k =~ /minimum/i;
                    }
                    \%a;
                } @{ $model->{dataset}{1136} }
            ];
        }
        else {
            my $a = $model->{dataset}{1136} =
              $model->{dataset}{1133} || $model->{dataset}{1134};
            foreach (qw(1133 1134)) {
                my $t = $model->{dataset}{$_} or next;
                my ($k1) = grep { !/^_/ } keys %{ $t->[1] } or next;
                my $k2 = ( $_ == 1133 ? 'Maximum' : 'Minimum' )
                  . ' network use factor';
                for ( my $c = 1 ; $c < @$t ; ++$c ) {
                    $a->[$c]{$k2} = $t->[$c]{$k1};
                }
            }
        }
    }

    if ( $model->{dataset}{1113} ) {
        my ($k1113) = grep { !/^_/ } keys %{ $model->{dataset}{1113}[1] };
        $model->{dataset}{1101} ||= [
            undef,
            @{ $model->{dataset}{1113} }[ 2, 6, 7, 8 ],
            {
                $k1113 => $model->{dataset}{1113}[4]{$k1113} +
                  $model->{dataset}{1113}[5]{$k1113}
            },
            $model->{dataset}{1113}[5]
        ];
        $model->{dataset}{1110} ||=
          [ undef, @{ $model->{dataset}{1113} }[ 1, 3 ] ];
        $model->{dataset}{1118} ||=
          [ undef, @{ $model->{dataset}{1113} }[ 9 .. 12 ] ];
    }
    else {
        my ($k1101) = grep { !/^_/ } keys %{ $model->{dataset}{1001}[1] };
        $model->{dataset}{1113} ||= [
            undef,
            $model->{dataset}{1110}[1],
            $model->{dataset}{1101}[1],
            $model->{dataset}{1110}[2],
            {
                $k1101 => $model->{dataset}{1101}[5]{$k1101} -
                  $model->{dataset}{1101}[6]{$k1101}
            },
            @{ $model->{dataset}{1101} }[ 6, 2, 3, 4 ],
            @{ $model->{dataset}{1118} }[ 1 .. 4 ],
        ];
    }

    if ( $model->{dataset}{1140} ) {
        $model->{dataset}{1105} ||= [
            map {
                my %a;
                while ( my ( $k, $v ) = each %$_ ) {
                    $a{$k} = $v if $k eq '_column' || $k =~ /diversity/i;
                }
                \%a;
            } @{ $model->{dataset}{1140} }
        ];
        $model->{dataset}{1122} ||= [
            map {
                my %a;
                while ( my ( $k, $v ) = each %$_ ) {
                    $a{$k} = $v if $k eq '_column' || $k =~ /simultaneous/i;
                }
                \%a;
            } @{ $model->{dataset}{1140} }
        ];
        $model->{dataset}{1131} ||= [
            map {
                my %a;
                while ( my ( $k, $v ) = each %$_ ) {
                    $a{$k} = $v if $k eq '_column' || $k =~ /asset/i;
                }
                \%a;
            } @{ $model->{dataset}{1140} }
        ];
        $model->{dataset}{1135} ||= [
            map {
                my %a;
                while ( my ( $k, $v ) = each %$_ ) {
                    $a{$k} = $v if $k eq '_column' || $k =~ /loss/i;
                }
                \%a;
            } @{ $model->{dataset}{1140} }
        ];
    }
    else {
        my $a = $model->{dataset}{1140} ||= [ {} ];
        foreach (qw(1105 1122 1131 1135)) {
            my $t = $model->{dataset}{$_} or next;
            my ( $kk, $match );
            if ( $_ == 1105 ) {
                $match = qr/diversity/i;
                $kk    = 'Diversity allowance between level exit and GSP Group';
            }
            elsif ( $_ == 1122 ) {
                $match = qr/simultaneous/i;
                $kk    = 'System simultaneous maximum load kW';
            }
            elsif ( $_ == 1131 ) {
                $match = qr/Assets/i;
                $kk    = 'Assets in CDCM model';
            }
            elsif ( $_ == 1135 ) {
                $match = qr/Loss/i;
                $kk    = 'Loss adjustment factor to transmission';
            }
            else {
                next;
            }
            for ( my $c = 1 ; $c < @$t ; ++$c ) {
                while ( my ( $k, $v ) = each %{ $t->[$c] } ) {
                    if ( $k eq '_column' ) { $a->[$c]{$k} ||= $v; }
                    else {
                        $a->[$c]{$kk} ||= $v;
                    }
                }
            }
        }
    }

    if ( $model->{dataset}{1194} ) {
        my ( $key1191, $key1192, $key1194 ) = map {
            ( grep { !/^_/ } keys %{ $model->{dataset}{$_}[1] } )[0]
        } qw(1191 1192 1194);
        $model->{dataset}{1191}[4]{$key1191} =
          $model->{dataset}{1194}[3]{$key1194};
        $model->{dataset}{1192}[4]{$key1192} =
          $model->{dataset}{1194}[1]{$key1194};
    }

}

1;
