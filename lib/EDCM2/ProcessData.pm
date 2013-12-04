package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others.
Copyright 2012-2013 Franck Latrémolière, Reckon LLP and others.

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

sub processData {
    my ($model) = @_;
    if ( $model->{dataset}
        && ( $model->{ldnoRev} && $model->{ldnoRev} =~ /nopop/i ) )
    {
        delete $model->{dataset}{$_} foreach qw(1181 1182 1183);
    }
    if (
           $model->{method}
        && $model->{method} !~ /none/i
        && (
            my $h = $model->{dataset}{
                $model->{method} =~ /FCP/i
                ? 911
                : 913
            }[1]
        )
      )
    {
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
            exists $ds->[$_]{$k} || ( $ds->[$_]{$k} = '' ) foreach 0 .. $#$ds;
        }
    }

    my ($daysInYearKey) = grep { !/^_/ } keys %{ $model->{dataset}{1113}[1] };
    my $daysInYear = $model->{dataset}{1113}[1]{$daysInYearKey};
    my ($hoursInRedKey) = grep { !/^_/ } keys %{ $model->{dataset}{1113}[3] };
    my $hoursInRed = $model->{dataset}{1113}[3]{$hoursInRedKey};

    if ( my $ds = $model->{dataset}{935} ) {
        my %tariffs;
        my $max = 0;
        $ds->[1]{$_} ||= "Tariff $_" foreach grep {
            my $t = $_;
            grep {
                     defined $ds->[$_]{$t}
                  && $ds->[$_]{$t} ne 'VOID'
                  && $ds->[$_]{$t} ne '#VALUE!';
            } 2 .. 66;
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
                @tariffs = sort { $a <=> $b } keys %tariffs,
                  defined $model->{numTariffs} ? () : ( $max + 1 .. $max + 16 );
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
        }
        if ( $model->{nonames} ) {
            $ds->[1]{$_} = "Tariff $_" foreach keys %{ $ds->[1] };
        }
        foreach my $k ( 1 .. $model->{numTariffs} ) {
            my $v = $ds->[1]{$k};
            if (    $v
                and $v ne 'Not used'
                and $v ne '#VALUE!'
                and $v ne '#N/A'
                and $v ne 'VOID'
                and $v !~ /^\s*$/s )
            {
                $ds->[$_]{$k} || ( $ds->[$_]{$k} = 'VOID' ) foreach 2 .. 6;
                exists $ds->[$_]{$k} || ( $ds->[$_]{$k} = '' )
                  foreach 7 .. $#$ds;
                $_ && /^[0-9.]+$/s && $_ > $daysInYear && ( $daysInYear = $_ )
                  foreach $ds->[22]{$k};
                $_ && /^[0-9.]+$/s && $_ > $hoursInRed && ( $hoursInRed = $_ )
                  foreach $ds->[23]{$k};
            }
            else {
                $ds->[1]{$k} = ' ';
                $ds->[$_]{$k} = 'VOID' foreach 2 .. 6;
                $ds->[$_]{$k} = ''     foreach 7 .. $#$ds;
            }
        }
    }

    $model->{dataset}{1113}[1]{$daysInYearKey} = $daysInYear;
    $model->{dataset}{1113}[3]{$hoursInRedKey} = $hoursInRed;

}

1;
