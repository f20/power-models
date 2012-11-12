package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2012 DCUSA Limited and contributors. All rights reserved.

Redistribution and use in source and binary forms, with or without modification,
are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation and/or
other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY DCUSA LIMITED AND CONTRIBUTORS "AS IS" AND ANY
EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL DCUSA LIMITED OR CONTRIBUTORS BE LIABLE FOR ANY
DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

=cut

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Shortcuts ':all';
use CDCM::TariffList;

sub tariffs {

    my ($model) = @_;

    my @allComponents;
    my @nonExcludedComponents;
    my %componentMap;
    my %componentUsed;

    my ( $allTariffs, $allTariffsByEndUser, $allEndUsers );

    my @endUserTypeList;

    my @allTariffs;

    my $hardMaxUnitRates = 10;

    my $maxUnitRates = 1;

    my $noLdno = !$model->{boundary} && !$model->{portfolio};

    {
        my @hashes = $model->tariffList;
        my @reserve;

        while ( @hashes || @reserve ) {

            my ( $endUser, $components ) =
              %{ @hashes ? shift @hashes : shift @reserve };

            if ( $endUser =~ /^(GSP|132|33)/i ) {
                next if !$model->{ehv};
                next if $model->{ehv} =~ /gen/i && $endUser !~ /gener/i;
                next
                  if $model->{ehv} =~ /33/ && $endUser =~ /132|33kV sub/i;
            }

            my %hash = map { %$_ } grep { ref $_ eq 'HASH' } @$components;

            next
              unless $hash{Included} && ( my $included = qr/$hash{Included}/ );

            $endUser .= "\n$hash{Name}" if !$model->{rawNames} && $hash{Name};

            my $boundary;

            if ( $model->{boundary} && 'boundary' =~ $included ) {
                $boundary = $model->{boundary};
            }
            elsif ( $model->{tariffs} ) {
                next if $model->{tariffs} !~ $included;
                next
                  if $hash{Excluded}
                  && $model->{tariffs} =~ qr/$hash{Excluded}/;
            }
            else {
                next if $included !~ /common/i;
            }

            my @tariffs = $endUser;    # "$endUser (directly connected)"

            if ( $model->{portfolio}
                and 'portfolio' =~ $included
                || @allTariffs < $model->{portfolio} )
            {
                foreach my $l (
                    'LV',              # 'LV Sub'
                    'HV',
                    $model->{pcd} && $model->{portfolio} == 15
                    ? (
                        split /\n/, <<EOL
0000
0001
0002
0010
0011
0100
0101
0110
0111
1000
1001
1100
1101
1110
1111
EOL
                    )
                    : $model->{pcd} && $model->{portfolio} == 5
                    ? ( 'GSP', 'EHV Sub', 'EHV', 'HV Sub' )
                    : (
                        $model->{portfolio} > 1 ? 'HV Sub' : (),
                        $model->{portfolio} > 2 ? '33kV'   : (),
                        $model->{portfolio} > 3 ? ( '33kV Sub', '132kV' )
                        : ()
                    )
                  )
                {
                         $endUser =~ /^(HV|33|132)/i  && $l =~ /^LV/i
                      || $endUser =~ /^(33|132)/i     && $l =~ /^HV/i
                      || $endUser =~ /^132/i          && $l =~ /^33/i
                      || $endUser =~ /^LV sub/i       && $l =~ /^LV/i
                      || $endUser =~ /^HV (pri|sub)/i && $l =~ /^HV/i
                      || push @tariffs, join "\n",
                      map { "LDNO $l: $_" } split /\n/, $endUser;
                }

            }

            my %map = map { $_ => 'nothing' } grep { !ref $_ } @$components;
            $map{'Unit rate 0 p/kWh'} = 'nothing'
              if $model->{alwaysUseRAG} && $endUser !~ /gener/i;

            $map{
                $model->{unauth} && $model->{unauth} =~ /day/i
                ? 'Exceeded capacity charge p/kVA/day'
                : 'Unauthorised demand charge p/kVAh'
              }
              = 'Capacity'
              if $model->{unauth}
              && $map{'Capacity charge p/kVA/day'};

            if ($boundary) {
                push @tariffs, "LDNO LV $_: $endUser"
                  foreach $boundary > 1
                  ? map { "B$_" } 1 .. $boundary
                  : 'boundary';
            }

            if ( $map{'Unit rates p/kWh'} ) {
                $map{"Unit rate $_ p/kWh"} = 'nothing'
                  foreach 1 .. $model->{timebands};
            }

            undef $componentUsed{$_} foreach keys %map;

            if ( $map{'Generation capacity rate p/kW/day'} ) {
                $map{'Generation capacity rate p/kW/day'} = 'PAYG yardstick kW';
                $map{'Reactive power charge p/kVArh'}     = 'PAYG kVArh'
                  if $map{'Reactive power charge p/kVArh'};
            }

            elsif ( $map{'Capacity charge p/kVA/day'} ) {
                $map{'Capacity charge p/kVA/day'} = 'Capacity';
                if ( $map{'Unit rate 2 p/kWh'} || $map{'Unit rate 0 p/kWh'} ) {
                    foreach ( grep { $map{"Unit rate $_ p/kWh"} }
                        1 .. $hardMaxUnitRates )
                    {
                        $map{"Unit rate $_ p/kWh"} = "Standard $_ kWh";
                        $maxUnitRates = $_ if $maxUnitRates < $_;
                    }
                }
                else {
                    $map{'Unit rate 1 p/kWh'} = 'Standard yardstick kWh';
                }
                $map{'Reactive power charge p/kVArh'} = 'Standard kVArh'
                  if $map{'Reactive power charge p/kVArh'};
            }

            elsif ($endUser !~ /(generat|unmeter)/i
                && $map{'Fixed charge p/MPAN/day'} )
            {
                $map{'Fixed charge p/MPAN/day'} = 'Fixed from network';
                if ( $map{'Unit rate 2 p/kWh'} || $map{'Unit rate 0 p/kWh'} ) {
                    foreach ( grep { $map{"Unit rate $_ p/kWh"} }
                        1 .. $hardMaxUnitRates )
                    {
                        $map{"Unit rate $_ p/kWh"} = "Standard $_ kWh";
                        $maxUnitRates = $_ if $maxUnitRates < $_;
                    }
                }
                else {
                    $map{'Unit rate 1 p/kWh'} = 'Standard yardstick kWh';
                }
                $map{'Reactive power charge p/kVArh'} = 'PAYG kVArh'
                  if $map{'Reactive power charge p/kVArh'};
            }

            else {
                if ( $map{'Unit rate 2 p/kWh'} || $map{'Unit rate 0 p/kWh'} ) {
                    foreach ( grep { $map{"Unit rate $_ p/kWh"} }
                        1 .. $hardMaxUnitRates )
                    {
                        $map{"Unit rate $_ p/kWh"} = (
                            $endUser =~ /(additional|related) MPAN/i
                            ? 'Standard'
                            : 'PAYG'
                        ) . " $_ kWh";
                        $maxUnitRates = $_ if $maxUnitRates < $_;
                    }
                }
                else {
                    $map{'Unit rate 1 p/kWh'} = (
                        $endUser =~ /(additional|related) MPAN/i
                        ? 'Standard'
                        : 'PAYG'
                    ) . ' yardstick kWh';
                }
                $map{'Reactive power charge p/kVArh'} = (
                    $endUser =~ /(additional|related) MPAN/i
                    ? 'Standard'
                    : 'PAYG'
                  )
                  . ' kVArh'
                  if $map{'Reactive power charge p/kVArh'};
            }

            if ( $map{'Fixed charge p/MPAN/day'} ) {
                $map{'Fixed charge p/MPAN/day'} =
                  $map{'Fixed charge p/MPAN/day'} eq 'nothing'
                  ? 'Customer'
                  : "$map{'Fixed charge p/MPAN/day'} & customer";
            }
            elsif ( $endUser !~ /(additional|related) MPAN/i ) {
                $map{"Unit rate $_ p/kWh"} .= " & customer"
                  foreach grep { $map{"Unit rate $_ p/kWh"} }
                  1 .. $hardMaxUnitRates;
            }

            $endUser = join "\n", map { $noLdno ? $_ : "> $_" } split /\n/,
              $endUser;

            for my $tariff ( $endUser, @tariffs ) {
                $componentMap{$tariff} = \%map;
            }

            push @endUserTypeList,
              Labelset(
                name => $endUser,
                list => \@tariffs
              );

            push @allTariffs, @tariffs;

        }
    }

    $model->{maxUnitRates} = $maxUnitRates;

    @nonExcludedComponents = grep { exists $componentUsed{$_} } (
        ( map { "Unit rate $_ p/kWh" } 1 .. $maxUnitRates ),
        split( /\n/, <<'EOL') ),
Fixed charge p/MPAN/day
Capacity charge p/kVA/day
Exceeded capacity charge p/kVA/day
Unauthorised demand charge p/kVAh
Generation capacity rate p/kW/day
EOL
      $model->{reactiveExcluded} ? () : 'Reactive power charge p/kVArh';

    $allEndUsers =
      Labelset( name => 'All end users', list => \@endUserTypeList );

    if ($noLdno) {

        $allTariffsByEndUser = $allTariffs = $allEndUsers;

    }

    else {

        $allTariffsByEndUser = Labelset(
            name   => 'All tariffs, grouped by end user',
            groups => \@endUserTypeList
        );

        $allTariffs = $noLdno
          || $model->{tariffOrder} ? $allTariffsByEndUser : Labelset(
            name => 'All tariffs',
            list => $model->{tariffOrder} ? \@allTariffs : [
                ( grep { !/LDNO/i } @allTariffs ),
                ( grep { /LDNO lv/i } @allTariffs ),
                ( grep { /LDNO hv/i && !/LDNO hv sub/i } @allTariffs ),
                ( grep { /LDNO hv sub/i } @allTariffs ),
                ( grep { /LDNO 33/i && !/LDNO 33kV sub/i } @allTariffs ),
                ( grep { /LDNO 33kV sub/i } @allTariffs ),
                ( grep { /LDNO 132/i } @allTariffs ),
                (
                    grep {
                             /LDNO/i
                          && !/LDNO lv/i
                          && !/LDNO hv/i
                          && !/LDNO 33/i
                          && !/LDNO 132/i
                    } @allTariffs
                )
            ]
          );

    }

    @allComponents = (
        @nonExcludedComponents,
        $model->{reactiveExcluded} ? 'Reactive power charge p/kVArh' : ()
    );

    $model->{tariffComponentMap} = Columnset(
        name    => 'Tariff components',
        columns => [
            map {
                my $component = $_;
                my @rules;
                $rules[$_] =
                  $componentMap{ $allTariffsByEndUser->{list}[$_] }{$component}
                  foreach $allTariffsByEndUser->indices;
                Constant(
                    rows => $allTariffsByEndUser,
                    name => $component,
                    data => [ map { $_ ? ucfirst($_) : undef } @rules ],
                    defaultFormat => 'textcell'
                  )
            } @allComponents
        ]
    );

    $allTariffs, $allTariffsByEndUser, $allEndUsers, \@allComponents,
      \@nonExcludedComponents, \%componentMap;

}

1;
