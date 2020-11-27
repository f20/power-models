package CDCM;

# Copyright 2014-2020 Franck Latrémolière and others.
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
    $table1202[0]{ $model->{tariffs}
          && $model->{tariffs} =~ /tcrbands/i ? '1202tcr' : '1202' };
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
    Business 5MVA: '^(?:|(?:LD|Q)NO .*: )(?:Demand Category|HV.*HH|HV Site Specific)'
    Business 690kVA: '^(?:|(?:LD|Q)NO .*: )(?:LV|LV Sub|HV) (?:HH Metered|Site Specific)'
    Business 69kVA: '^(?:Small|LV) (?:Network )?(?:.*Non[- ]Domestic(?: [UTN]|$)|(?:Sub )?(?:HH Metered|Site Specific))'
    Continuous 23kVA: '^(?:Small|LV).*Non[- ]Domestic(?: [UTN]|$)'
    Continuous 5MVA: '^(?:|(?:LD|Q)NO .*: )(?:Demand Category|HV.*HH|HV Site Specific)'
    Continuous 690kVA: '^(?:|(?:LD|Q)NO .*: )(?:LV|LV Sub|HV) (?:HH Metered|Site Specific)'
    Continuous 69kVA: '^(?:Small|LV) (?:Network )?(?:.*Non[- ]Domestic(?: [UTN]|$)|(?:Sub )?(?:HH Metered|Site Specific))'
    Demand 500kVA: All-the-way metered demand
    Domestic 1550kWh: '(?:^|: )(?:LV Network Domestic Disabled|Domestic [UT])'
    Domestic 3100kWh: '(?:^|: )(?:LV Network Domestic Disabled|Domestic [UT])'
    Domestic 8400kWh: '(?:^|: )(?:LV Network Domestic Disabled|Domestic [UT])'
    Generation 1MVA: All-the-way generation
    LivingElecHeat 5MVA: '^(?:|(?:LD|Q)NO .*: )(?:Demand Category|HV.*HH|HV Site Specific)'
    LivingElecHeat 690kVA: '^(?:|(?:LD|Q)NO .*: )(?:LV|LV Sub|HV) (?:HH Metered|Site Specific)'
    LivingFuelHeat 5MVA: '^(?:|(?:LD|Q)NO .*: )(?:Demand Category|HV.*HH|HV Site Specific)'
    LivingFuelHeat 690kVA: '^(?:|(?:LD|Q)NO .*: )(?:LV|LV Sub|HV) (?:HH Metered|Site Specific)'
    Off-peak 23kVA: '^(?:Small|LV).*Non[- ]Domestic(?: [UTN]|$)'
    Off-peak 5MVA: '^(?:|(?:LD|Q)NO .*: )(?:Demand Category|HV.*HH|HV Site Specific)'
    Off-peak 690kVA: '^(?:|(?:LD|Q)NO .*: )(?:LV|LV Sub|HV) (?:HH Metered|Site Specific)'
    Off-peak 69kVA: '^(?:Small|LV) (?:Network )?(?:.*Non[- ]Domestic(?: [UTN]|$)|(?:Sub )?(?:HH Metered|Site Specific))'
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
1202tcr:
  - _table: 1202. Consumption assumptions for illustrative customers
  - 10MVA high usage: 390
    10MVA no usage: 380
    40kVA high usage: 230
    40kVA no usage: 220
    750kVA high usage: 310
    750kVA no usage: 300
    Domestic high usage: 160
    Domestic mid usage: 150
    Domestic no usage: 140
    HV band 1/2 high usage: 330
    HV band 1/2 no usage: 320
    HV band 2/3 high usage: 350
    HV band 2/3 no usage: 340
    HV band 3/4 high usage: 370
    HV band 3/4 no usage: 360
    LV band 1/2 high usage: 250
    LV band 1/2 no usage: 240
    LV band 2/3 high usage: 270
    LV band 2/3 no usage: 260
    LV band 3/4 high usage: 290
    LV band 3/4 no usage: 280
    Non-domestic band 1/2: 180
    Non-domestic band 2/3: 190
    Non-domestic band 3/4: 200
    Non-domestic high usage: 210
    Non-domestic no usage: 170
    _column: Order
  - 10MVA high usage: '^(LV Sub|HV) Site Specific (No Residual|Band 4)'
    10MVA no usage: '^(LV Sub|HV) Site Specific (No Residual|Band 4)'
    40kVA high usage: '^(LV|LV Sub|HV) Site Specific (No Residual|Band 1)'
    40kVA no usage: '^(LV|LV Sub|HV) Site Specific (No Residual|Band 1)'
    750kVA high usage: '^(LV|LV Sub) Site Specific (No Residual|Band 4)'
    750kVA no usage: '^(LV|LV Sub) Site Specific (No Residual|Band 4)'
    Domestic high usage: '^Domestic Aggregated$'
    Domestic mid usage: '^Domestic Aggregated$'
    Domestic no usage: '^Domestic Aggregated$'
    HV band 1/2 high usage: '^HV Site Specific (No Residual|Band 1|Band 2)'
    HV band 1/2 no usage: '^HV Site Specific (No Residual|Band 1|Band 2)'
    HV band 2/3 high usage: '^HV Site Specific (No Residual|Band 2|Band 3)'
    HV band 2/3 no usage: '^HV Site Specific (No Residual|Band 2|Band 3)'
    HV band 3/4 high usage: '^HV Site Specific (No Residual|Band 3|Band 4)'
    HV band 3/4 no usage: '^HV Site Specific (No Residual|Band 3|Band 4)'
    LV band 1/2 high usage: '^(LV|LV Sub) Site Specific (No Residual|Band 1|Band 2)'
    LV band 1/2 no usage: '^(LV|LV Sub) Site Specific (No Residual|Band 1|Band 2)'
    LV band 2/3 high usage: '^(LV|LV Sub) Site Specific (No Residual|Band 2|Band 3)'
    LV band 2/3 no usage: '^(LV|LV Sub) Site Specific (No Residual|Band 2|Band 3)'
    LV band 3/4 high usage: '^(LV|LV Sub) Site Specific (No Residual|Band 3|Band 4)'
    LV band 3/4 no usage: '^(LV|LV Sub) Site Specific (No Residual|Band 3|Band 4)'
    Non-domestic band 1/2: '^Non-Domestic Aggregated (No Residual|Band 1|Band 2)'
    Non-domestic band 2/3: '^Non-Domestic Aggregated (No Residual|Band 2|Band 3)'
    Non-domestic band 3/4: '^(LV Site Specific Band 1|Non-Domestic Aggregated (No Residual|Band 3|Band 4))'
    Non-domestic high usage: '^(LV Site Specific Band 1|Non-Domestic Aggregated (No Residual|Band 4))'
    Non-domestic no usage: '^Non-Domestic Aggregated (No Residual|Band 1)'
    _column: Tariff selection
  - 10MVA high usage: ''
    10MVA no usage: ''
    40kVA high usage: ''
    40kVA no usage: ''
    750kVA high usage: ''
    750kVA no usage: ''
    Domestic high usage: 35
    Domestic mid usage: 35
    Domestic no usage: ''
    HV band 1/2 high usage: ''
    HV band 1/2 no usage: ''
    HV band 2/3 high usage: ''
    HV band 2/3 no usage: ''
    HV band 3/4 high usage: ''
    HV band 3/4 no usage: ''
    LV band 1/2 high usage: ''
    LV band 1/2 no usage: ''
    LV band 2/3 high usage: ''
    LV band 2/3 no usage: ''
    LV band 3/4 high usage: ''
    LV band 3/4 no usage: ''
    Non-domestic band 1/2: 60
    Non-domestic band 2/3: 60
    Non-domestic band 3/4: 60
    Non-domestic high usage: ''
    Non-domestic no usage: ''
    _column: Peak-time hours/week
  - 10MVA high usage: ''
    10MVA no usage: ''
    40kVA high usage: ''
    40kVA no usage: ''
    750kVA high usage: ''
    750kVA no usage: ''
    Domestic high usage: ''
    Domestic mid usage: 49
    Domestic no usage: ''
    HV band 1/2 high usage: ''
    HV band 1/2 no usage: ''
    HV band 2/3 high usage: ''
    HV band 2/3 no usage: ''
    HV band 3/4 high usage: ''
    HV band 3/4 no usage: ''
    LV band 1/2 high usage: ''
    LV band 1/2 no usage: ''
    LV band 2/3 high usage: ''
    LV band 2/3 no usage: ''
    LV band 3/4 high usage: ''
    LV band 3/4 no usage: ''
    Non-domestic band 1/2: ''
    Non-domestic band 2/3: ''
    Non-domestic band 3/4: ''
    Non-domestic high usage: ''
    Non-domestic no usage: ''
    _column: Off-peak hours/week
  - 10MVA high usage: ''
    10MVA no usage: ''
    40kVA high usage: ''
    40kVA no usage: ''
    750kVA high usage: ''
    750kVA no usage: ''
    Domestic high usage: 2.5
    Domestic mid usage: 0.649
    Domestic no usage: ''
    HV band 1/2 high usage: ''
    HV band 1/2 no usage: ''
    HV band 2/3 high usage: ''
    HV band 2/3 no usage: ''
    HV band 3/4 high usage: ''
    HV band 3/4 no usage: ''
    LV band 1/2 high usage: ''
    LV band 1/2 no usage: ''
    LV band 2/3 high usage: ''
    LV band 2/3 no usage: ''
    LV band 3/4 high usage: ''
    LV band 3/4 no usage: ''
    Non-domestic band 1/2: 0.6
    Non-domestic band 2/3: 2.1
    Non-domestic band 3/4: 4.2
    Non-domestic high usage: ''
    Non-domestic no usage: ''
    _column: Peak-time load (kW)
  - 10MVA high usage: ''
    10MVA no usage: ''
    40kVA high usage: ''
    40kVA no usage: ''
    750kVA high usage: ''
    750kVA no usage: ''
    Domestic high usage: ''
    Domestic mid usage: 0.178
    Domestic no usage: ''
    HV band 1/2 high usage: ''
    HV band 1/2 no usage: ''
    HV band 2/3 high usage: ''
    HV band 2/3 no usage: ''
    HV band 3/4 high usage: ''
    HV band 3/4 no usage: ''
    LV band 1/2 high usage: ''
    LV band 1/2 no usage: ''
    LV band 2/3 high usage: ''
    LV band 2/3 no usage: ''
    LV band 3/4 high usage: ''
    LV band 3/4 no usage: ''
    Non-domestic band 1/2: ''
    Non-domestic band 2/3: ''
    Non-domestic band 3/4: ''
    Non-domestic high usage: ''
    Non-domestic no usage: ''
    _column: Off-peak load (kW)
  - 10MVA high usage: 8500
    10MVA no usage: ''
    40kVA high usage: 34
    40kVA no usage: ''
    750kVA high usage: 637.5
    750kVA no usage: ''
    Domestic high usage: 1.5
    Domestic mid usage: 0.333
    Domestic no usage: ''
    HV band 1/2 high usage: 358.7
    HV band 1/2 no usage: ''
    HV band 2/3 high usage: 850
    HV band 2/3 no usage: ''
    HV band 3/4 high usage: 1530
    HV band 3/4 no usage: ''
    LV band 1/2 high usage: 68
    LV band 1/2 no usage: ''
    LV band 2/3 high usage: 127.5
    LV band 2/3 no usage: ''
    LV band 3/4 high usage: 196.35
    LV band 3/4 no usage: ''
    Non-domestic band 1/2: 0.3
    Non-domestic band 2/3: 1.05
    Non-domestic band 3/4: 2.15
    Non-domestic high usage: 58.65
    Non-domestic no usage: ''
    _column: Load at other times (kW)
  - 10MVA high usage: 10000
    10MVA no usage: 10000
    40kVA high usage: 40
    40kVA no usage: 40
    750kVA high usage: 750
    750kVA no usage: 750
    Domestic high usage: 23
    Domestic mid usage: 23
    Domestic no usage: 23
    HV band 1/2 high usage: 422
    HV band 1/2 no usage: 422
    HV band 2/3 high usage: 1000
    HV band 2/3 no usage: 1000
    HV band 3/4 high usage: 1800
    HV band 3/4 no usage: 1800
    LV band 1/2 high usage: 80
    LV band 1/2 no usage: 80
    LV band 2/3 high usage: 150
    LV band 2/3 no usage: 150
    LV band 3/4 high usage: 231
    LV band 3/4 no usage: 231
    Non-domestic band 1/2: 23
    Non-domestic band 2/3: 23
    Non-domestic band 3/4: 23
    Non-domestic high usage: 69
    Non-domestic no usage: 23
    _column: Capacity (kVA)
