package EDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others. All rights reserved.

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

sub randomiseCut {
    my ($model) = @_;
    srand();
    my @oldtariffs = grep {
        my ( $name, $type );
        ( $name = $model->{dataset}{953}[2]{$_} )
          && ( $type = $model->{dataset}{953}[3]{$_} )
          && $type eq 'Demand'
          && $name ne 'Not used'
          && $name ne '#VALUE!';
    } keys %{ $model->{dataset}{953}[1] };
    my $numTariffs = 4;
    my @tariffs    =
      map { $oldtariffs[ int( rand() * $#oldtariffs ) ] } 1 .. $numTariffs;
    foreach my $col ( @{ $model->{dataset}{953} } ) {
        $col =
          { map { ( $_ => $col->{ $tariffs[ $_ - 1 ] } ); } 1 .. $numTariffs };
    }
    $model->{numTariffs} = $numTariffs;
}

sub randomiseAggressive {

    my ($model) = @_;

    srand();

    foreach my $tc (
        qw(
        t911c4 t911c5 t911c6 t911c7 t911c8 t911c9 t911c10 t911c11 t911c12 t911c13 t913c4 t913c5 t913c6 t913c7 t913c8 t913c9 t913c10 t913c11 t953c7 t953c10 t953c11 t953c12 t953c13 t953c14 t953c15 t953c18 t953c20 t953c21 t953c22 t953c23 t953c24 t953c27 t953c28 t953c5 t953c6 t953c8
        )
      )
    {
        next
          unless my ( $table, $column ) = ( $tc =~ /([0-9]+)c([0-9]+)/ );
        my $hashref = $model->{dataset}{$table}[$column];
        foreach ( keys %$hashref ) {
            $hashref->{$_} = rand() > 0.5 ? 0 : -200 + 600 * rand();
        }
    }

    foreach my $tc (
        qw(t953c9)
      )
    {
        next
          unless my ( $table, $column ) = ( $tc =~ /([0-9]+)c([0-9]+)/ );
        my $hashref = $model->{dataset}{$table}[$column];
        foreach ( keys %$hashref ) {
            $hashref->{$_} = rand() > 0.5 ? 0 : 600 * rand();
        }
    }

    foreach my $tc (
        qw(
        t1105c1 t1105c2 t1105c3 t1105c4 t1105c5 t1105c6
        t1108c1 t1109c1 t1111c1 t1111c2 t1111c3 t1111c4
        t1112c1 t1112c2 t1112c3 t1112c4 t11112c5
        t1122c2 t1122c3 t1122c4 t1122c5 t1122c6
        t1131c2 t1131c3 t1131c4 t1131c5 t1131c6 t1131c7 t1131c8 t1131c9 t1131c10 t1131c11
        t1132c1 t1168c1
        t1134c1 t1134c2 t1134c3 t1134c4 t1134c5 t1134c6
        t1135c1 t1135c2 t1135c3 t1135c4 t1135c5 t1135c6
        )
      )
    {
        next
          unless my ( $table, $column ) = ( $tc =~ /([0-9]+)c([0-9]+)/ );
        my $hashref = $model->{dataset}{$table}[$column];
        foreach ( keys %$hashref ) {
            $hashref->{$_} = -200 + 600 * rand();
        }
    }

    foreach my $column ( 1 .. 6 ) {
        my $hashref = $model->{dataset}{1133}[$column];
        foreach ( keys %$hashref ) {
            $hashref->{$_} =
              $model->{dataset}{1134}[$column] {'Minimum network use factor'} +
              600 * rand();
        }
    }

    foreach my $tc (qw(t953c25 t953c26)) {
        next
          unless my ( $table, $column ) = ( $tc =~ /([0-9]+)c([0-9]+)/ );
        my $hashref = $model->{dataset}{$table}[$column];
        foreach ( keys %$hashref ) {
            $hashref->{$_} = rand() > 0.25 ? 'FALSE' : 'TRUE';
        }
    }

}

sub randomiseNormal {

    my ($model) = @_;

    srand();

    foreach my $tc (
        qw(
        t911c4 t911c5 t911c6 t911c7 t911c8 t911c9 t911c10 t911c11 t911c12 t911c13 t913c4 t913c5 t913c6 t913c7 t913c8 t913c9 t913c10 t913c11 t953c10 t953c11 t953c12 t953c13 t953c15 t953c18 t953c20 t953c21 t953c22 t953c23 t953c24 t953c27 t953c28
        )
      )
    {
        next
          unless my ( $table, $column ) = ( $tc =~ /([0-9]+)c([0-9]+)/ );
        my $hashref = $model->{dataset}{$table}[$column];
        foreach ( keys %$hashref ) {
            $hashref->{$_} = rand() > 0.5 ? 0 : 0 + 600 * rand();
        }
    }

    foreach my $tc (
        qw(
        t953c5 t953c6
        )
      )
    {
        next
          unless my ( $table, $column ) = ( $tc =~ /([0-9]+)c([0-9]+)/ );
        my $hashref = $model->{dataset}{$table}[$column];
        foreach ( keys %$hashref ) {
            $hashref->{$_} = rand() > 0.3 ? 0 : 0 + 100 * rand();
        }
    }

    foreach my $tc (
        qw(t953c14)
      )
    {
        next
          unless my ( $table, $column ) = ( $tc =~ /([0-9]+)c([0-9]+)/ );
        my $hashref = $model->{dataset}{$table}[$column];
        foreach ( keys %$hashref ) {
            $hashref->{$_} = rand() > 0.5 ? 0 : rand();
        }
    }

    foreach my $tc (
        qw(t953c9)
      )
    {
        next
          unless my ( $table, $column ) = ( $tc =~ /([0-9]+)c([0-9]+)/ );
        my $hashref = $model->{dataset}{$table}[$column];
        foreach ( keys %$hashref ) {
            $hashref->{$_} = rand() > 0.5 ? 0 : 600 * rand();
        }
    }

    foreach my $tc (
        qw(t953c8)
      )
    {
        next
          unless my ( $table, $column ) = ( $tc =~ /([0-9]+)c([0-9]+)/ );
        my $hashref = $model->{dataset}{$table}[$column];
        foreach ( keys %$hashref ) {
            $hashref->{$_} = rand() > 0.5 ? 0 : 50000 * rand();
        }
    }

    foreach my $tc (
        qw(t953c7)
      )
    {
        next
          unless my ( $table, $column ) = ( $tc =~ /([0-9]+)c([0-9]+)/ );
        my $hashref = $model->{dataset}{$table}[$column];
        foreach ( keys %$hashref ) {
            $hashref->{$_} =
              $model->{dataset}{953}[8] + ( rand() > 0.5 ? 0 : 50000 * rand() );
        }
    }

    foreach my $tc (
        qw(
        t1105c1 t1105c2 t1105c3 t1105c4 t1105c5 t1105c6
        t1108c1 t1109c1
        t1122c2 t1122c3 t1122c4 t1122c5 t1122c6
        t1131c2 t1131c3 t1131c4 t1131c5 t1131c6 t1131c7 t1131c8 t1131c9 t1131c10 t1131c11
        t1132c1
        t1135c1 t1135c2 t1135c3 t1135c4 t1135c5 t1135c6
        )
      )
    {
        next
          unless my ( $table, $column ) = ( $tc =~ /([0-9]+)c([0-9]+)/ );
        my $hashref = $model->{dataset}{$table}[$column];
        foreach ( keys %$hashref ) {
            $hashref->{$_} = 0 + 600 * rand();
        }
    }

    foreach my $tc (
        qw(
        t1112c2 t1112c3 t1112c4 t1112c5 t1114c1
        )
      )
    {
        next
          unless my ( $table, $column ) = ( $tc =~ /([0-9]+)c([0-9]+)/ );
        my $hashref = $model->{dataset}{$table}[$column];
        foreach ( keys %$hashref ) {
            $hashref->{$_} = 0 + 200_000_000 * rand();
        }
    }

    foreach my $tc (
        qw(
        t1112c1
        )
      )
    {
        next
          unless my ( $table, $column ) = ( $tc =~ /([0-9]+)c([0-9]+)/ );
        my $hashref = $model->{dataset}{$table}[$column];
        foreach ( keys %$hashref ) {
            $hashref->{$_} = 200_000_000 + 400_000_000 * rand();
        }
    }

    foreach my $tc (qw(t953c25 t953c26)) {
        next
          unless my ( $table, $column ) = ( $tc =~ /([0-9]+)c([0-9]+)/ );
        my $hashref = $model->{dataset}{$table}[$column];
        foreach ( keys %$hashref ) {
            $hashref->{$_} = rand() > 0.25 ? 'FALSE' : 'TRUE';
        }
    }

}

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
                  $model->{method} =~ /FCP/i         ? 911
                : $model->{method} =~ /LRIC.*split/i ? 913
                : 912
            }[1]
        )
      )
    {
        my $max = $model->{numLocations} || 0;
        foreach ( keys %$h ) {
            if (   $h->{$_}
                && $h->{$_} ne 'Not used'
                && $h->{$_} ne '#VALUE!' )
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
              || $model->{numLocations} ? 0 : $max > 6 ? 12 : 1 ) + $max;
    }

    if ( my $ds = $model->{dataset}{953} ) {
        my $max = $model->{numTariffs} || 0;
        while ( my ( $k, $v ) = each %{ $ds->[2] } ) {
            $max = $k
              if $k =~ /^[0-9]+$/
              and $k > $max
              and $v
              and $v ne 'Not used'
              and $v ne '#VALUE!'
              and $ds->[1]{$k} || $ds->[2]{$k};
        }
        $model->{numTariffs} =
          (      $model->{randomise}
              || $model->{small}
              || $model->{numTariffs} ? 0 : $max > 6 ? 12 : 1 ) + $max;
        if ( $model->{nonames} ) {
            $ds->[1] = {};
            $ds->[2]{$_} =
                $ds->[2]{$_} =~ /^NR_/ ? 'Customer group 1'
              : $ds->[2]{$_} =~ /^LU_/ ? 'Customer group 2'
              : 'Other customer'
              foreach keys %{ $ds->[2] };
        }
        0 and warn join "\n", map { $_->{_column} || 'undef' } @$ds;
        if ( @$ds < -26 ) {
            $ds->[25] = { map { ( $_ => '' ); } keys %{ $ds->[1] } };
            @{$ds}[ map { 1 + $_ }
              qw(4 5 6 7 8 9 10 11 12 13 14 15 16 17 23 24) ] =
              @{$ds}[ map { 1 + $_ }
              qw(12 24 13 9 10 14 8 24 16 5 6 11 7 4 15 23) ];
        }
        if ( @$ds < -27 ) {
            splice @{$ds}, 11, 0, { map { ( $_ => '' ); } keys %{ $ds->[1] } };
        }
    }

    if ( $model->{vedcm} > 52 && $model->{vedcm} < 61 ) {

=head Shankar spreadsheet

GSP	132kV circuits	132kV/EHV	EHV circuits	EHV/HV	132kV/HV
#VALUE!	2.246	1.558	3.290	2.380	2.768

GSP	132kV circuits	132kV/EHV	EHV circuits	EHV/HV	132kV/HV
#VALUE!	0.273	0.677	0.332	0.631	0.697

=cut

        $model->{dataset}{1133} = [
            {},
            { 'Maximum network use factor' => '#VALUE' },
            { 'Maximum network use factor' => 2.246 },
            { 'Maximum network use factor' => 1.558 },
            { 'Maximum network use factor' => 3.290 },
            { 'Maximum network use factor' => 2.380 },
            { 'Maximum network use factor' => 2.768 },
            { 'Maximum network use factor' => 'Thursday 31 March 2011' },
        ];

        $model->{dataset}{1134} = [
            {},
            { 'Minimum network use factor' => '#VALUE' },
            { 'Minimum network use factor' => .273 },
            { 'Minimum network use factor' => .677 },
            { 'Minimum network use factor' => .332 },
            { 'Minimum network use factor' => .631 },
            { 'Minimum network use factor' => .697 },
            { 'Minimum network use factor' => 'Thursday 31 March 2011' },
        ];

    }
}

1;
