package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others.
Copyright 2011-2016 Franck Latrémolière, Reckon LLP and others.

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
use YAML;

local undef $/;
binmode DATA, ':utf8';
my @tariffSpec = Load <DATA>;

sub tariffSpec {
    my ($model) = @_;
    my @list = @tariffSpec;
    @list = grep { $_->[0] !~ /no reactive/i; } @tariffSpec
      unless $model->{tariffs} && $model->{tariffs} =~ /gennoreact/i;
    @list = grep {
        $_->[0] !~ /related/i || !grep { $_ eq 'Unit rates p/kWh' } @$_;
      } @list
      unless $model->{tariffs} && $model->{tariffs} =~ /offpeakhh/i;
    @list = grep {
        grep { $_ eq 'Unit rates p/kWh' } @$_;
    } @list if $model->{tariffs} && $model->{tariffs} =~ /omitnhh/i;
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
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC0
- Unit rates p/kWh
---
- LV half hourly WC
- Name: LV Network Non-Domestic Non-CT
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC0
- Unit rates p/kWh
---
- LV half hourly
- Name: LV HH Metered
  Portfolio: 1
- Capacity charge p/kVA/day
- Fixed charge p/MPAN/day
- PC0
- Reactive power charge p/kVArh
- Unit rates p/kWh
---
- LV substation half hourly
- Name: LV Sub HH Metered
  Portfolio: 1
- Capacity charge p/kVA/day
- Fixed charge p/MPAN/day
- PC0
- Reactive power charge p/kVArh
- Unit rates p/kWh
---
- HV half hourly
- Name: HV HH Metered
  Portfolio: 1
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
- LV generation half hourly single rate no reactive
- Name: LV Generation Intermittent no RP charge
- Fixed charge p/MPAN/day
- PC0
- Unit rate 1 p/kWh
---
- LV generation half hourly
- Name: LV Generation Non-Intermittent
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC0
- Reactive power charge p/kVArh
- Unit rates p/kWh
---
- LV generation half hourly no reactive
- Name: LV Generation Non-Intermittent no RP charge
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
- LV substation generation half hourly single rate no reactive
- Name: LV Sub Generation Intermittent no RP charge
- Fixed charge p/MPAN/day
- PC0
- Unit rate 1 p/kWh
---
- LV substation generation half hourly
- Name: LV Sub Generation Non-Intermittent
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC0
- Reactive power charge p/kVArh
- Unit rates p/kWh
---
- LV substation generation half hourly no reactive
- Name: LV Sub Generation Non-Intermittent no RP charge
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
- HV generation half hourly single rate no reactive
- Name: HV Generation Intermittent no RP charge
- Fixed charge p/MPAN/day
- PC0
- Unit rate 1 p/kWh
---
- HV generation half hourly
- Name: HV Generation Non-Intermittent
  Portfolio: 1
- Fixed charge p/MPAN/day
- PC0
- Reactive power charge p/kVArh
- Unit rates p/kWh
---
- HV generation half hourly no reactive
- Name: HV Generation Non-Intermittent no RP charge
- Fixed charge p/MPAN/day
- PC0
- Unit rates p/kWh
