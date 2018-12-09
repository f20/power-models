package CDCM;

# Copyright 2014-2018 Franck Latrémolière, Reckon LLP and others.
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
my @table1202 = Load <DATA>;

sub table1202 {
    my ($model) = @_;
    $table1202[0]{1202};
}

1;

__DATA__
---
1202:
  - _table: 1202. Consumption assumptions for illustrative customers
  - Business 23kVA: 110
    Business 5MVA: 410
    Business 690kVA: 310
    Business 69kVA: 210
    Continuous 23kVA: 120
    Continuous 5MVA: 420
    Continuous 690kVA: 320
    Continuous 69kVA: 220
    Demand 500kVA: 815
    Domestic 1550kWh: 710
    Domestic 3100kWh: 720
    Domestic 8400kWh: 740
    Generation 1MVA: 830
    LivingElecHeat 5MVA: 455
    LivingElecHeat 690kVA: 355
    LivingFuelHeat 5MVA: 465
    LivingFuelHeat 690kVA: 365
    Off-peak 23kVA: 130
    Off-peak 5MVA: 430
    Off-peak 690kVA: 330
    Off-peak 69kVA: 230
    _column: Order
  - Business 23kVA: '^(?:Small|LV).*Non[- ]Domestic(?: [UTN]|$)'
    Business 5MVA: '^(?:|(?:LD|Q)NO .*: )(?:Demand Category|HV.*HH)'
    Business 690kVA: '^(?:|(?:LD|Q)NO .*: )(?:LV|LV Sub|HV) HH Metered'
    Business 69kVA: '^(?:Small|LV) (?:Network )?(?:.*Non[- ]Domestic(?: [UTN]|$)|(?:Sub )?HH Metered)'
    Continuous 23kVA: '^(?:Small|LV).*Non[- ]Domestic(?: [UTN]|$)'
    Continuous 5MVA: '^(?:|(?:LD|Q)NO .*: )(?:Demand Category|HV.*HH)'
    Continuous 690kVA: '^(?:|(?:LD|Q)NO .*: )(?:LV|LV Sub|HV) HH Metered'
    Continuous 69kVA: '^(?:Small|LV) (?:Network )?(?:.*Non[- ]Domestic(?: [UTN]|$)|(?:Sub )?HH Metered)'
    Demand 500kVA: All-the-way metered demand
    Domestic 1550kWh: '(?:^|: )(?:LV Network Domestic Disabled|Domestic [UT])'
    Domestic 3100kWh: '(?:^|: )(?:LV Network Domestic Disabled|Domestic [UT])'
    Domestic 8400kWh: '(?:^|: )(?:LV Network Domestic Disabled|Domestic [UT])'
    Generation 1MVA: All-the-way generation
    LivingElecHeat 5MVA: '^(?:|(?:LD|Q)NO .*: )(?:Demand Category|HV.*HH)'
    LivingElecHeat 690kVA: '^(?:|(?:LD|Q)NO .*: )(?:LV|LV Sub|HV) HH Metered'
    LivingFuelHeat 5MVA: '^(?:|(?:LD|Q)NO .*: )(?:Demand Category|HV.*HH)'
    LivingFuelHeat 690kVA: '^(?:|(?:LD|Q)NO .*: )(?:LV|LV Sub|HV) HH Metered'
    Off-peak 23kVA: '^(?:Small|LV).*Non[- ]Domestic(?: [UTN]|$)'
    Off-peak 5MVA: '^(?:|(?:LD|Q)NO .*: )(?:Demand Category|HV.*HH)'
    Off-peak 690kVA: '^(?:|(?:LD|Q)NO .*: )(?:LV|LV Sub|HV) HH Metered'
    Off-peak 69kVA: '^(?:Small|LV) (?:Network )?(?:.*Non[- ]Domestic(?: [UTN]|$)|(?:Sub )?HH Metered)'
    _column: Tariff selection
  - Business 23kVA: 62
    Business 5MVA: 62
    Business 690kVA: 62
    Business 69kVA: 62
    Continuous 23kVA: 0
    Continuous 5MVA: 0
    Continuous 690kVA: 0
    Continuous 69kVA: 0
    Demand 500kVA: 0
    Domestic 1550kWh: 35
    Domestic 3100kWh: 35
    Domestic 8400kWh: 35
    Generation 1MVA: 0
    LivingElecHeat 5MVA: 35
    LivingElecHeat 690kVA: 35
    LivingFuelHeat 5MVA: 35
    LivingFuelHeat 690kVA: 35
    Off-peak 23kVA: 93.4
    Off-peak 5MVA: 93.4
    Off-peak 690kVA: 93.4
    Off-peak 69kVA: 93.4
    _column: Peak-time hours/week
  - Business 23kVA: 49
    Business 5MVA: 49
    Business 690kVA: 49
    Business 69kVA: 49
    Continuous 23kVA: 49
    Continuous 5MVA: 49
    Continuous 690kVA: 49
    Continuous 69kVA: 49
    Demand 500kVA: 49
    Domestic 1550kWh: 49
    Domestic 3100kWh: 49
    Domestic 8400kWh: 49
    Generation 1MVA: 49
    LivingElecHeat 5MVA: 49
    LivingElecHeat 690kVA: 49
    LivingFuelHeat 5MVA: 49
    LivingFuelHeat 690kVA: 49
    Off-peak 23kVA: 49
    Off-peak 5MVA: 49
    Off-peak 690kVA: 49
    Off-peak 69kVA: 49
    _column: Off-peak hours/week
  - Business 23kVA: 4.83
    Business 5MVA: 3500
    Business 690kVA: 483
    Business 69kVA: 48.3
    Continuous 23kVA: 6.21
    Continuous 5MVA: 4500
    Continuous 690kVA: 621
    Continuous 69kVA: 62.1
    Demand 500kVA: 450
    Domestic 1550kWh: 0.3245
    Domestic 3100kWh: 0.649
    Domestic 8400kWh: 1.2
    Generation 1MVA: 1000
    LivingElecHeat 5MVA: 1200
    LivingElecHeat 690kVA: 120
    LivingFuelHeat 5MVA: 1622.5
    LivingFuelHeat 690kVA: 162.25
    Off-peak 23kVA: 0
    Off-peak 5MVA: 0
    Off-peak 690kVA: 0
    Off-peak 69kVA: 0
    _column: Peak-time load (kW)
  - Business 23kVA: 1.5525
    Business 5MVA: 1125
    Business 690kVA: 155.25
    Business 69kVA: 15.525
    Continuous 23kVA: 6.21
    Continuous 5MVA: 4500
    Continuous 690kVA: 621
    Continuous 69kVA: 62.1
    Demand 500kVA: 450
    Domestic 1550kWh: 0.089
    Domestic 3100kWh: 0.178
    Domestic 8400kWh: 1.4
    Generation 1MVA: 1000
    LivingElecHeat 5MVA: 1400
    LivingElecHeat 690kVA: 140
    LivingFuelHeat 5MVA: 445
    LivingFuelHeat 690kVA: 44.5
    Off-peak 23kVA: 6.21
    Off-peak 5MVA: 4500
    Off-peak 690kVA: 621
    Off-peak 69kVA: 62.1
    _column: Off-peak load (kW)
  - Business 23kVA: 1.5525
    Business 5MVA: 1125
    Business 690kVA: 155.25
    Business 69kVA: 15.525
    Continuous 23kVA: 6.21
    Continuous 5MVA: 4500
    Continuous 690kVA: 621
    Continuous 69kVA: 62.1
    Demand 500kVA: 450
    Domestic 1550kWh: 0.1665
    Domestic 3100kWh: 0.333
    Domestic 8400kWh: 0.6
    Generation 1MVA: 1000
    LivingElecHeat 5MVA: 600
    LivingElecHeat 690kVA: 60
    LivingFuelHeat 5MVA: 832.5
    LivingFuelHeat 690kVA: 83.25
    Off-peak 23kVA: 6.21
    Off-peak 5MVA: 4500
    Off-peak 690kVA: 621
    Off-peak 69kVA: 62.1
    _column: Load at other times (kW)
  - Business 23kVA: 23
    Business 5MVA: 5000
    Business 690kVA: 690
    Business 69kVA: 69
    Continuous 23kVA: 23
    Continuous 5MVA: 5000
    Continuous 690kVA: 690
    Continuous 69kVA: 69
    Demand 500kVA: 500
    Domestic 1550kWh: 6
    Domestic 3100kWh: 9
    Domestic 8400kWh: 18
    Generation 1MVA: 1000
    LivingElecHeat 5MVA: 5000
    LivingElecHeat 690kVA: 690
    LivingFuelHeat 5MVA: 5000
    LivingFuelHeat 690kVA: 690
    Off-peak 23kVA: 23
    Off-peak 5MVA: 5000
    Off-peak 690kVA: 690
    Off-peak 69kVA: 69
    _column: Capacity (kVA)
