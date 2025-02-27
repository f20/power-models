﻿package CDCM;

# Copyright 2009-2011 Energy Networks Association Limited and others.
# Copyright 2011-2025 Franck Latrémolière and others.
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
use YAML;

local undef $/;
binmode DATA, ':utf8';
my @tariffSpec = Load <DATA>;
$/ = "\n";

sub tariffSpec {
    my ($model) = @_;
    my @list = @tariffSpec;
    @list =
      grep { !$_->[1]{Name} || $_->[1]{Name} !~ /HV Sub/i; } @tariffSpec
      unless $model->{tariffs} && $model->{tariffs} =~ /hvsub/i;
    @list =
      grep { !$_->[1]{Name} || $_->[1]{Name} !~ /no rp charge/i; } @tariffSpec
      unless $model->{tariffs} && $model->{tariffs} =~ /gennoreact/i;
    @list = grep {
        $_->[0] !~ /related /i || !grep { $_ eq 'Unit rates p/kWh' } @$_;
      } @list
      unless $model->{tariffs} && $model->{tariffs} =~ /offpeakhh/i;
    @list = grep {
        grep { $_ eq 'Unit rates p/kWh' } @$_;
    } @list if $model->{tariffs} && $model->{tariffs} =~ /omitnhh/i;
    map { $_->[1]{Name} = $_->[1]{ $model->{tariffNameField} }; }
      grep { $_->[1]{ $model->{tariffNameField} }; } @list
      if $model->{tariffNameField};
    @list = grep { $_->[1]{Name}; } @list unless $model->{tariffsWithNoNames};

    if ( $model->{tariffs} && $model->{tariffs} =~ /tcrbands/i ) {
        @list = map {
            my @a = @$_;
            $a[1]{Name} !~ /gener/i
              && $a[1]{Name} =~ /Site Specific|Non-Domestic Aggregated/i
              ? (
                map {
                    [
                        "$a[0] $_",
                        {
                            %{ $a[1] }, Name => "$a[1]{Name} $_",
                        },
                        @a[ 2 .. $#a ]
                    ];
                } 'No Residual',
                'Band 1', 'Band 2', 'Band 3', 'Band 4',
              )
              : $_;
        } @list;
    }

    $_->[1]{Portfolio} = $model->{portfolio}
      foreach grep { $_->[1]{Portfolio}; } @list;

    @list;

}

1;

__DATA__
---
- LV domestic unrestricted
- Name: Domestic Unrestricted
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC1
- Unit rate 1 p/kWh
---
- LV domestic two rates
- Name: Domestic Two Rate
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC2
- Unit rate 1 p/kWh
- Unit rate 2 p/kWh
---
- LV related MPAN domestic off peak
- Name: Domestic Off Peak (related MPAN)
  Portfolio: 1
- PC2
- Unit rate 0 p/kWh
- Unit rate 1 p/kWh
---
- LV non-domestic small unrestricted
- Name: Small Non Domestic Unrestricted
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC3
- Unit rate 1 p/kWh
---
- LV non-domestic small two rates
- Name: Small Non Domestic Two Rate
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC4
- Unit rate 1 p/kWh
- Unit rate 2 p/kWh
---
- LV related MPAN non-domestic off peak
- Name: Small Non Domestic Off Peak (related MPAN)
  Portfolio: 1
- PC4
- Unit rate 0 p/kWh
- Unit rate 1 p/kWh
---
- LV non-domestic profiles 5-8 two rates
- Name: LV Medium Non-Domestic
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC5-8
- Unit rate 1 p/kWh
- Unit rate 2 p/kWh
---
- LV substation non-domestic profiles 5-8 two rates
- Name: LV Sub Medium Non-Domestic
- Fixed charge p/MPAN/day
- PC5-8
- Unit rate 1 p/kWh
- Unit rate 2 p/kWh
---
- HV non-domestic profiles 5-8 two rates
- Name: HV Medium Non-Domestic
- Fixed charge p/MPAN/day
- PC5-8
- Unit rate 1 p/kWh
- Unit rate 2 p/kWh
---
- LV half hourly domestic
- Name: LV Network Domestic
  Name268: Domestic Aggregated
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC0
- Unit rates p/kWh
---
- LV half hourly domestic (related MPAN)
- Name: LV Network Domestic (related MPAN)
  Name268: Domestic Aggregated (Related MPAN)
  Portfolio: 1
- PC0
- Unit rates p/kWh
---
- LV half hourly WC
- Name: LV Network Non-Domestic Non-CT
  Name268: Non-Domestic Aggregated
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC0
- Unit rates p/kWh
---
- LV half hourly WC (related MPAN)
- Name: LV Network Non-Domestic Non-CT (related MPAN)
  Name268: Non-Domestic Aggregated (Related MPAN)
  Portfolio: 1
- PC0
- Unit rates p/kWh
---
- LV half hourly
- Name: LV HH Metered
  Name268: LV Site Specific
  Portfolio: 1
- Capacity charge p/kVA/day
- Fixed charge p/MPAN/day
- PC0
- Reactive power charge p/kVArh
- Unit rates p/kWh
---
- LV substation half hourly
- Name: LV Sub HH Metered
  Name268: LV Sub Site Specific
  Portfolio: 1
- Capacity charge p/kVA/day
- Fixed charge p/MPAN/day
- PC0
- Reactive power charge p/kVArh
- Unit rates p/kWh
---
- HV half hourly
- Name: HV HH Metered
  Name268: HV Site Specific
  Portfolio: 1
- Capacity charge p/kVA/day
- Fixed charge p/MPAN/day
- PC0
- Reactive power charge p/kVArh
- Unit rates p/kWh
---
- HV substation half hourly
- Name: HV Sub HH Metered
  Name268: HV Sub Site Specific
- Capacity charge p/kVA/day
- Fixed charge p/MPAN/day
- PC0
- Reactive power charge p/kVArh
- Unit rates p/kWh
---
- LV unmetered category A
- Name: NHH UMS category A
  Portfolio: 1
- PC8
- Unit rate 0 p/kWh
- Unit rate 1 p/kWh
---
- LV unmetered category B
- Name: NHH UMS category B
  Portfolio: 1
- PC1
- Unit rate 0 p/kWh
- Unit rate 1 p/kWh
---
- LV unmetered category C
- Name: NHH UMS category C
  Portfolio: 1
- PC1
- Unit rate 0 p/kWh
- Unit rate 1 p/kWh
---
- LV unmetered category D
- Name: NHH UMS category D
  Portfolio: 1
- PC1
- Unit rate 0 p/kWh
- Unit rate 1 p/kWh
---
- LV unmetered pseudo half hourly
- Name: LV UMS (Pseudo HH Metered)
  Name268: Unmetered Supplies
  Portfolio: 1
- PC0
- Unit rates p/kWh
---
- LV generation non half hourly
- Name: LV Generation NHH or Aggregate HH
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC8&0
- Unit rate 1 p/kWh
---
- LV substation generation non half hourly
- Name: LV Sub Generation NHH
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC8
- Unit rate 1 p/kWh
---
- LV generation half hourly single rate
- Name: LV Generation Intermittent
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC0
- Reactive power charge p/kVArh
- Unit rate 1 p/kWh
---
- LV generation half hourly single rate (no reactive)
- Name: LV Generation Intermittent no RP charge
- Fixed charge p/MPAN/day
- PC0
- Unit rate 1 p/kWh
---
- LV generation half hourly (no reactive)
- Name268: LV Generation Aggregated
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC0
- Unit rates p/kWh
---
- LV substation generation half hourly (no reactive)
- Name268: LV Sub Generation Aggregated
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC0
- Unit rates p/kWh
---
- LV generation half hourly
- Name: LV Generation Non-Intermittent
  Name268: LV Generation Site Specific
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC0
- Reactive power charge p/kVArh
- Unit rates p/kWh
---
- LV generation half hourly (no reactive)
- Name: LV Generation Non-Intermittent no RP charge
  Name268: LV Generation Site Specific no RP charge
- Fixed charge p/MPAN/day
- PC0
- Unit rates p/kWh
---
- LV substation generation half hourly single rate
- Name: LV Sub Generation Intermittent
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC0
- Reactive power charge p/kVArh
- Unit rate 1 p/kWh
---
- LV substation generation half hourly single rate (no reactive)
- Name: LV Sub Generation Intermittent no RP charge
- Fixed charge p/MPAN/day
- PC0
- Unit rate 1 p/kWh
---
- LV substation generation half hourly
- Name: LV Sub Generation Non-Intermittent
  Name268: LV Sub Generation Site Specific
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC0
- Reactive power charge p/kVArh
- Unit rates p/kWh
---
- LV substation generation half hourly (no reactive)
- Name: LV Sub Generation Non-Intermittent no RP charge
  Name268: LV Sub Generation Site Specific no RP charge
- Fixed charge p/MPAN/day
- PC0
- Unit rates p/kWh
---
- HV generation half hourly single rate
- Name: HV Generation Intermittent
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC0
- Reactive power charge p/kVArh
- Unit rate 1 p/kWh
---
- HV generation half hourly single rate (no reactive)
- Name: HV Generation Intermittent no RP charge
- Fixed charge p/MPAN/day
- PC0
- Unit rate 1 p/kWh
---
- HV generation half hourly
- Name: HV Generation Non-Intermittent
  Name268: HV Generation Site Specific
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC0
- Reactive power charge p/kVArh
- Unit rates p/kWh
---
- HV generation half hourly (no reactive)
- Name: HV Generation Non-Intermittent no RP charge
  Name268: HV Generation Site Specific no RP charge
- Fixed charge p/MPAN/day
- PC0
- Unit rates p/kWh
