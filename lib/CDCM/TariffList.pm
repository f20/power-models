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
my @tariffList = Load <DATA>;
$/ = "\n";

sub tariffList {
    my ($model) = @_;
    my @tspec;
    foreach (@tariffList) {
        my ( $endUser, $components ) = %$_;
        if ( $endUser =~ /^(GSP|132|33)/i ) {
            next if !$model->{ehv};
            next if $model->{ehv} =~ /gen/i && $endUser !~ /gener/i;
            next
              if $model->{ehv} =~ /33/ && $endUser =~ /132|33kV sub/i;
        }
        my %hash = map { %$_ } grep { ref $_ eq 'HASH' } @$components;
        next unless $hash{Included};
        my $included = qr/$hash{Included}/ or next;
        my $boundary;
        if ( $model->{boundary} && 'boundary' =~ $included ) {
            $boundary = $model->{boundary};
        }
        my $portfolio;
        if ( $model->{portfolio} && 'portfolio' =~ $included ) {
            $portfolio = $model->{portfolio};
        }
        if ( $model->{tariffs} ) {
            next if $model->{tariffs} !~ $included;
            next
              if $hash{Excluded}
              && $model->{tariffs} =~ qr/$hash{Excluded}/;
        }
        else {
            next if $included !~ /common/i;
        }
        my @components = sort grep { !ref $_ } @$components;
        push @tspec,
          [
            $endUser,
            {
                $portfolio  ? ( Portfolio => $portfolio )  : (),
                $boundary   ? ( Boundary  => $boundary )   : (),
                $hash{Name} ? ( Name      => $hash{Name} ) : (),
            },
            @components,
          ];
    }
    @tspec;
}

1;

__DATA__
---
LV domestic unrestricted:
  - PC1
  - Name: Domestic Unrestricted
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Included: special|common|CE|CN|EDF|ENW|SPEN|SSE|WPD|WS2|test|t4|Opt1|Opt3|Opt4|Opt5|Opt6|portfolio|T9|toy|dcp161toy|simple
---
LV domestic two rates:
  - PC2
  - Name: Domestic Two Rate
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Unit rate 2 p/kWh
  - Included: common|CE|CN|EDF|SPEN|SSE|WPD|WS2|test|t4|Opt1|Opt4|Opt5|Opt6|portfolio|T9|simple
---
LV related MPAN domestic off peak:
  - PC2
  - Name: Domestic Off Peak (related MPAN)
  - Unit rate 0 p/kWh
  - Unit rate 1 p/kWh
  - Included: common|CE|EDF|SSE|WS2|test|t4|Opt1|Opt4|portfolio
---
LV non-domestic small unrestricted:
  - PC3
  - Name: Small Non Domestic Unrestricted
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Included: common|CE|CN|EDF|ENW|SPEN|SSE|WPD|WS2|Opt1|Opt3|Opt4|Opt5|Opt6|portfolio|dcp161toy|simple
---
LV non-domestic small two rates:
  - PC4
  - Name: Small Non Domestic Two Rate
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Unit rate 2 p/kWh
  - Included: common|CE|EDF|SPEN|SSE|WPD|Opt1|Opt4|Opt5|Opt6|portfolio|simple
---
LV related MPAN non-domestic off peak:
  - PC4
  - Name: Small Non Domestic Off Peak (related MPAN)
  - Unit rate 0 p/kWh
  - Unit rate 1 p/kWh
  - Included: common|CE|CN|EDF|ENW|SPEN|SSE|portfolio
---
LV non-domestic profiles 5-8 two rates:
  - PC5-8
  - Name: LV Medium Non-Domestic
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Unit rate 2 p/kWh
  - Included: special|common|CE|CN|SPEN|WPD|WS2|test|t4|Opt1|Opt4|Opt5|Opt6|portfolio|simple
  - Excluded: dcp270
---
LV non-domestic profiles 5-8 two rates whole current meter:
  - PC5-8
  - Name: LV Medium Non-Domestic WC
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Unit rate 2 p/kWh
  - Included: sparecap|portfolio
---
LV substation non-domestic profiles 5-8 two rates:
  - PC5-8
  - Name: LV Sub Medium Non-Domestic
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Unit rate 2 p/kWh
  - Included: common|SPEN|SSE|WPD|Opt1
  - Excluded: dcp270
---
HV non-domestic profiles 5-8 two rates:
  - PC5-8
  - Name: HV Medium Non-Domestic
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Unit rate 2 p/kWh
  - Included: common|CE|CN|SPEN|SSE|WPD|Opt1
  - Excluded: dcp270
---
LV Agg WC Domestic:
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Included: sparecap|portfolio
---
LV Agg WC Non-Domestic:
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Included: sparecap|portfolio
---
LV Agg CT Non-Domestic:
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Included: sparecap|portfolio
---
LV HH WC:
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: sparecap|portfolio
---
LV half hourly domestic:
  - PC0
  - Name: LV Network Domestic
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Included: pc12hh|dcp179|portfolio
---
LV half hourly domestic HV netting:
  - PC0
  - Name: LV Network Domestic HV Local Source
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Included: nettinghvnonct|portfolio
---
LV half hourly non-domestic aggregated:
  - PC0
  - Name: LV Network Non-Domestic Non-CT
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Included: pc34hh|dcp179|portfolio
---
LV half hourly non-domestic aggregated HV netting:
  - PC0
  - Name: LV Network Non-Domestic Non-CT HV Local Source
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Included: nettinghvnonct|portfolio
---
LV half hourly CT:
  - PC0
  - Name: LV Network Non-Domestic CT
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: dcp179|portfolio
---
LV substation half hourly CT:
  - PC0
  - Name: LV Sub CT
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: dcp179|portfolio
---
HV half hourly CT:
  - PC0
  - Name: HV Network CT
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: dcp179|portfolio
---
LV half hourly:
  - PC0
  - Name: LV HH Metered
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: special|common|EDF|WS2|WS3|test|t4|Opt2|Opt3|Opt4|Opt5|Opt6|boundary|portfolio|dcp161toy|simple|toy2
  - Excluded: dcp179
---
LV half hourly HV netting:
  - PC0
  - Name: LV HH Metered HV Local Source
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: nettinghv|portfolio
---
LV half hourly EHV netting:
  - PC0
  - Name: LV HH Metered EHV Local Source
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: nettingehv|portfolio
---
LV substation half hourly:
  - PC0
  - Name: LV Sub HH Metered
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: common|EDF|WS2|WS3|Opt2|Opt3|Opt4|portfolio|simple
  - Excluded: dcp179
---
LV substation half hourly HV netting:
  - PC0
  - Name: LV Sub HH Metered HV Local Source
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: nettinghv|portfolio
---
LV substation half hourly EHV netting:
  - PC0
  - Name: LV Sub HH Metered EHV Local Source
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: nettingehv|portfolio
---
HV half hourly:
  - PC0
  - Name: HV HH Metered
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: common|EDF|WS2|WS3|test|t4|Opt2|Opt3|Opt4|Opt5|Opt6|portfolio|T9|dcp161toy|simple|toy2
  - Excluded: dcp179
---
HV half hourly HV netting:
  - PC0
  - Name: HV HH Metered HV Local Source
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: nettinghv|portfolio
---
HV half hourly EHV netting:
  - PC0
  - Name: HV HH Metered EHV Local Source
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: nettingehv|portfolio
---
HV substation half hourly:
  - PC0
  - Name: HV Sub HH Metered
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: common|WS2|WS3|Opt2|Opt3|Opt4|portfolio|simple
  - Excluded: dcp163
---
HV substation (132kV) half hourly:
  - PC0
  - Name: HV Sub (132kV) HH Metered
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: sub132|portfolio
---
HV substation (33kV EDCM) half hourly:
  - Name: Demand Category 1111
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: edcm
---
HV substation (132kV EDCM) half hourly:
  - Name: Demand Category 1001
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: edcm
---
33kV half hourly EDCM:
  - Name: Demand Category 1110
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: edcm
---
33kV half hourly:
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: WS2|Opt2|Opt3|Opt4|simple
---
33kV non-domestic half hourly single rate:
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: ehv1
---
33kV non-domestic half hourly two rates:
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Unit rate 2 p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: CE|CN|ENW|SPEN|SSE|WPD|WS2|Opt1|ehv2
---
33kV substation half hourly:
  - Name: Demand Category 1100
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: edcm
---
33kV substation non-domestic half hourly two rates:
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Unit rate 2 p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: ehv2
---
132kV half hourly EDCM:
  - Name: Demand Category 1000
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: edcm
---
132kV half hourly:
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: WS2|test|t4|Opt2|Opt3|Opt4
---
132kV non-domestic half hourly two rates:
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Unit rate 2 p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: CE|CN|ENW|SPEN|SSE|WPD|WS2|Opt1|ehv2
---
GSP half hourly:
  - Name: Demand Category 0000
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: edcm
---
GSP non-domestic half hourly single rate:
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: ehv1
---
GSP non-domestic half hourly two rates:
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Unit rate 2 p/kWh
  - Capacity charge p/kVA/day
  - Reactive power charge p/kVArh
  - Included: ehv2
---
LV unmetered category A:
  - Name: Unmetered Continuous
  - Unit rate 1 p/kWh
  - Included: portfolio|tums|simple
---
LV unmetered category A:
  - PC8
  - Name: NHH UMS category A
  - Unit rate 0 p/kWh
  - Unit rate 1 p/kWh
  - Included: portfolio|tums|dcp130
---
LV unmetered category B:
  - Name: Unmetered Dusk to Dawn
  - Unit rate 1 p/kWh
  - Included: portfolio|tums|simple
---
LV unmetered category B:
  - PC1
  - Name: NHH UMS category B
  - Unit rate 0 p/kWh
  - Unit rate 1 p/kWh
  - Included: portfolio|tums|dcp130
---
LV unmetered category C:
  - Unit rate 1 p/kWh
  - Included: portfolio|tums
---
LV unmetered category C:
  - PC1
  - Name: NHH UMS category C
  - Unit rate 0 p/kWh
  - Unit rate 1 p/kWh
  - Included: portfolio|tums|dcp130
---
LV unmetered category D:
  - Unit rate 1 p/kWh
  - Included: portfolio|tums
---
LV unmetered category D:
  - PC1
  - Name: NHH UMS category D
  - Unit rate 0 p/kWh
  - Unit rate 1 p/kWh
  - Included: portfolio|tums|dcp130
---
LV unmetered non half hourly:
  - PC1&8
  - Name: NHH UMS
  - Unit rate 1 p/kWh
  - Included: common|CE|CN|SPEN|SSE|WS2|test|t4|Opt1|Opt4|Opt5|Opt6|portfolio
  - Excluded: dcp130
---
LV unmetered pseudo half hourly:
  - PC0
  - Name: LV UMS (Pseudo HH Metered)
  - Unit rates p/kWh
  - Included: common|EDF|WS2|WS3|test|t4|Opt2|Opt3|Opt4|portfolio
---
LV generation non half hourly:
  - PC8
  - Name: LV Generation NHH
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Included: common|CE|CN|EDF|ENW|SPEN|SSE|WPD|WS2|test|t4|Opt1|Opt2|Opt3|Opt4|Opt5|Opt6|portfolio
  - Excluded: pc12hh|pc34hh|dcp179
---
LV generation non half hourly:
  - PC8&0
  - Name: LV Generation NHH or Aggregate HH
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Included: pc12hh|pc34hh|dcp179|portfolio
---
LV generation (GDP) non half hourly:
  - PC8
  - Name: LV Generation (GDP area) NHH
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Included: gendom|portfolio
---
LV generation (GDT) non half hourly:
  - PC8
  - Name: LV Generation (GDT area) NHH
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Included: gendom|portfolio
---
LV substation generation non half hourly:
  - PC8
  - Name: LV Sub Generation NHH
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Included: gensub|portfolio
---
LV substation generation (GDP) non half hourly:
  - PC8
  - Name: LV Sub Generation (GDP area) NHH
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Included: gendom|portfolio
---
LV substation generation (GDT) non half hourly:
  - PC8
  - Name: LV Sub Generation (GDT area) NHH
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Included: gendom|portfolio
---
LV generation half hourly single rate:
  - PC0
  - Name: LV Generation Intermittent
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: common|CE|CN|EDF|ENW|SPEN|SSE|WPD|WS2|test|t4|Opt1|Opt2|Opt3|Opt4|Opt5|Opt6|portfolio
---
LV generation half hourly single rate no reactive:
  - PC0
  - Name: LV Generation Intermittent no RP charge
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Included: gennoreact
---
LV generation (GDP) half hourly single rate:
  - PC0
  - Name: LV Generation (GDP area) Intermittent
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: gendom|portfolio
---
LV generation (GDT) half hourly single rate:
  - PC0
  - Name: LV Generation (GDT area) Intermittent
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: gendom|portfolio
---
LV generation half hourly:
  - PC0
  - Name: LV Generation Non-Intermittent
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: common|CE|CN|EDF|ENW|SPEN|SSE|WPD|WS2|test|t4|Opt1|Opt2|Opt3|Opt4|Opt5|Opt6|portfolio|boundary|simplegen
---
LV generation half hourly HV netting:
  - PC0
  - Name: LV Generation HV Local Supply
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: nettinghvgen|portfolio
---
LV generation half hourly EHV netting:
  - PC0
  - Name: LV Generation EHV Local Supply
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: nettingehvgen|portfolio
---
LV generation half hourly no reactive:
  - PC0
  - Name: LV Generation Non-Intermittent no RP charge
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Included: gennoreact
---
LV generation (GDP) half hourly:
  - PC0
  - Name: LV Generation (GDP area) Non-Intermittent
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: gendom|portfolio
---
LV generation (GDT) half hourly:
  - PC0
  - Name: LV Generation (GDT area) Non-Intermittent
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: gendom|portfolio
---
LV substation generation half hourly single rate:
  - PC0
  - Name: LV Sub Generation Intermittent
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: gensub|portfolio
---
LV substation generation half hourly single rate no reactive:
  - PC0
  - Name: LV Sub Generation Intermittent no RP charge
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Included: gennoreact
---
LV substation generation (GDP) half hourly single rate:
  - PC0
  - Name: LV Sub Generation (GDP area) Intermittent
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: gendom|portfolio
---
LV substation generation (GDT) half hourly single rate:
  - PC0
  - Name: LV Sub Generation (GDT area) Intermittent
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: gendom|portfolio
---
LV substation generation half hourly:
  - PC0
  - Name: LV Sub Generation Non-Intermittent
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: gensub|portfolio
---
LV substation generation half hourly HV netting:
  - PC0
  - Name: LV Sub Generation HV Local Supply
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: nettinghvgen|portfolio
---
LV substation generation half hourly EHV netting:
  - PC0
  - Name: LV Sub Generation EHV Local Supply
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: nettingehvgen|portfolio
---
LV substation generation half hourly no reactive:
  - PC0
  - Name: LV Sub Generation Non-Intermittent no RP charge
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Included: gennoreact
---
LV substation generation (GDP) half hourly:
  - PC0
  - Name: LV Sub Generation (GDP area) Non-Intermittent
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: gendom|portfolio
---
LV substation generation (GDT) half hourly:
  - PC0
  - Name: LV Sub Generation (GDT area) Non-Intermittent
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: gendom|portfolio
---
HV generation half hourly single rate:
  - PC0
  - Name: HV Generation Intermittent
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: common|CE|CN|EDF|ENW|SPEN|SSE|WPD|WS2|Opt1|Opt2|Opt3|Opt4|Opt5|Opt6|portfolio
---
HV generation half hourly single rate no reactive:
  - PC0
  - Name: HV Generation Intermittent no RP charge
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Included: gennoreact
---
HV generation half hourly single rate (67% credit):
  - PC0
  - Name: HV Generation Intermittent Low GDA
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: dcp137|portfolio
---
HV generation half hourly single rate (33% credit):
  - PC0
  - Name: HV Generation Intermittent Medium GDA
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: dcp137|portfolio
---
HV generation half hourly single rate (0% credit):
  - PC0
  - Name: HV Generation Intermittent High GDA
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: dcp137|portfolio
---
HV generation (GDP) half hourly single rate:
  - PC0
  - Name: HV Generation (GDP area) Intermittent
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: gendom|portfolio
---
HV generation (GDT) half hourly single rate:
  - PC0
  - Name: HV Generation (GDT area) Intermittent
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: gendom|portfolio
---
HV generation half hourly:
  - PC0
  - Name: HV Generation Non-Intermittent
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: common|WS2|WS3|Opt6|portfolio|simplegen|toy2
---
HV generation half hourly HV netting:
  - PC0
  - Name: HV Generation HV Local Supply
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: nettinghvgen|portfolio
---
HV generation half hourly EHV netting:
  - PC0
  - Name: HV Generation EHV Local Supply
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: nettingehvgen|portfolio
---
HV generation half hourly no reactive:
  - PC0
  - Name: HV Generation Non-Intermittent no RP charge
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Included: gennoreact
---
HV generation half hourly (67% credit):
  - PC0
  - Name: HV Generation Non-Intermittent Low GDA
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: dcp137|portfolio
---
HV generation half hourly (33% credit):
  - PC0
  - Name: HV Generation Non-Intermittent Medium GDA
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: dcp137|portfolio
---
HV generation half hourly (0% credit):
  - PC0
  - Name: HV Generation Non-Intermittent High GDA
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: dcp137|portfolio
---
HV generation (GDP) half hourly:
  - PC0
  - Name: HV Generation (GDP area) Non-Intermittent
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: gendom|portfolio
---
HV generation (GDT) half hourly:
  - PC0
  - Name: HV Generation (GDT area) Non-Intermittent
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: gendom|portfolio
---
HV substation generation half hourly single rate:
  - PC0
  - Name: HV Sub Generation Intermittent
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: gensub|portfolio
  - Excluded: dcp163
---
HV substation generation half hourly single rate no reactive:
  - PC0
  - Name: HV Sub Generation Intermittent no RP charge
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Included: gennoreact
  - Excluded: dcp163
---
HV substation generation (GDP) half hourly single rate:
  - PC0
  - Name: HV Sub Generation (GDP area) Intermittent
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: gendom|portfolio
  - Excluded: dcp163
---
HV substation generation (GDT) half hourly single rate:
  - PC0
  - Name: HV Sub Generation (GDT area) Intermittent
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: gendom
  - Excluded: dcp163
---
HV substation generation half hourly:
  - PC0
  - Name: HV Sub Generation Non-Intermittent
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: gensub|portfolio
  - Excluded: dcp163
---
HV substation generation half hourly no reactive:
  - PC0
  - Name: HV Sub Generation Non-Intermittent no RP charge
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Included: gennoreact
  - Excluded: dcp163
---
HV substation generation (GDP) half hourly:
  - PC0
  - Name: HV Sub Generation (GDP area) Non-Intermittent
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: gendom|portfolio
  - Excluded: dcp163
---
HV substation generation (GDT) half hourly:
  - PC0
  - Name: HV Sub Generation (GDT area) Non-Intermittent
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: gendom
  - Excluded: dcp163
---
HV substation (132kV) generation half hourly single rate:
  - PC0
  - Name: HV Sub (132kV) Generation Intermittent
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: gensub132|portfolio
---
HV substation (132kV) generation (GDT) half hourly single rate:
  - PC0
  - Name: HV Sub Generation (GDT area) Intermittent
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: gensub132gendom
---
HV substation (132kV) generation half hourly:
  - PC0
  - Name: HV Sub (132kV) Generation Non-Intermittent
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: gensub132|portfolio
---
HV substation (132kV) generation (GDT) half hourly:
  - PC0
  - Name: HV Sub Generation (GDT area) Non-Intermittent
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: gensub132gendom
---
HV generation (wind tidal wave photovoltaic capacity):
  - Fixed charge p/MPAN/day
  - Generation capacity rate p/kW/day
  - Reactive power charge p/kVArh
  - Included: WS2|test
---
HV generation (hydro capacity):
  - Fixed charge p/MPAN/day
  - Generation capacity rate p/kW/day
  - Reactive power charge p/kVArh
  - Included: WS2
---
HV generation (non-intermittent non-CHP capacity):
  - Fixed charge p/MPAN/day
  - Generation capacity rate p/kW/day
  - Reactive power charge p/kVArh
  - Included: WS2
---
HV generation (CHP capacity):
  - Fixed charge p/MPAN/day
  - Generation capacity rate p/kW/day
  - Reactive power charge p/kVArh
  - Included: WS2
---
HV substation (33kV EDCM) generation half hourly single rate:
  - Name: Generation Category 1111
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: edcmgen
---
HV substation (132kV EDCM) generation half hourly single rate:
  - Name: Generation Category 1001
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: edcmgen
---
33kV generation half hourly single rate:
  - Name: Generation Category 1110
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: CE|CN|EDF|ENW|SPEN|SSE|WPD|WS2|Opt1|Opt2|Opt3|Opt4|Opt6|edcmgen
---
33kV substation generation half hourly single rate:
  - Name: Generation Category 1100
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: CE|CN|EDF|ENW|SPEN|SSE|WPD|WS2|Opt1|Opt2|Opt3|Opt4|Opt6|edcmgen
---
132kV generation half hourly single rate:
  - Name: Generation Category 1000
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: CE|CN|EDF|ENW|SPEN|SSE|WPD|WS2|test|t4|Opt1|Opt2|Opt3|Opt4|Opt6|edcmgen
---
GSP generation half hourly single rate:
  - Name: Generation Category 0000
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: edcmgen
---
33kV generation (GDT) half hourly single rate:
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: gensub132gendom|portfolio
---
33kV generation (GDT) half hourly:
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: gensub132gendom|portfolio
---
33kV generation half hourly:
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: simplegen
---
33kV substation generation (GDT) half hourly single rate:
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: gensub132gendom|portfolio
---
33kV substation generation (GDT) half hourly:
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: gensub132gendom|portfolio
---
132kV generation (GDT) half hourly single rate:
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: gensub132gendom|portfolio
---
132kV generation (GDT) half hourly:
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: gensub132gendom|portfolio
---
GSP generation (GDT) half hourly single rate:
  - Fixed charge p/MPAN/day
  - Unit rate 1 p/kWh
  - Reactive power charge p/kVArh
  - Included: gensub132gendom|portfolio
---
GSP generation (GDT) half hourly:
  - Fixed charge p/MPAN/day
  - Unit rates p/kWh
  - Reactive power charge p/kVArh
  - Included: gensub132gendom|portfolio
