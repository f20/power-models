package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and contributors. All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted
provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of
conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of
conditions and the following disclaimer in the documentation and/or other materials provided
with the distribution.

THIS SOFTWARE IS PROVIDED BY ENERGY NETWORKS ASSOCIATION LIMITED AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL ENERGY
NETWORKS ASSOCIATION LIMITED OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF
THE POSSIBILITY OF SUCH DAMAGE.

=cut

use warnings;
use strict;
use utf8;

sub processData {
    my ($model) = @_;
    if ( $model->{dataset}
        && !( $model->{ldnoRev} && $model->{ldnoRev} =~ /pop/i ) )
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
              || $model->{numLocations} ? 0 : $max > 24 ? 12 : 3 ) + $max;
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
        my $max = $model->{numTariffs} || 0;
        while ( my ( $k, $v ) = each %{ $ds->[1] } ) {
            $max = $k
              if $k =~ /^[0-9]+$/
              and $k > $max
              and $v
              and $v ne 'Not used'
              and $v ne '#VALUE!'
              and $v !~ /^\s*$/s
              and $ds->[1]{$k};
        }
        $model->{numTariffs} =
          (      $model->{randomise}
              || $model->{small}
              || $model->{numTariffs} ? 0 : $max > 24 ? 12 : 3 ) + $max;
        if ( $model->{nonames} ) {
            $ds->[1]{$_} =
                $ds->[1]{$_} =~ /^NR_/ ? 'Customer group 1'
              : $ds->[1]{$_} =~ /^LU_/ ? 'Customer group 2'
              : 'Other customer'
              foreach keys %{ $ds->[1] };
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
                exists $ds->[$_]{$k} || ( $ds->[$_]{$k} = '' )
                  foreach 2 .. $#$ds;
                $ds->[$_]{$k} = 'VOID'
                  foreach
                  grep { defined $ds->[$_]{$k} && $ds->[$_]{$k} eq '#N/A' }
                  2 .. $#$ds;
                $ds->[2]{$k} = 1.42
                  if !$ds->[2]{$k}
                  && !grep { $ds->[$_]{$k} && $ds->[$_]{$k} ne 'VOID' } 3 .. 6;
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
