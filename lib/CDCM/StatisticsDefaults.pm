package CDCM;

=head Copyright licence and disclaimer

Copyright 2014-2016 Franck Latrémolière, Reckon LLP and others.

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
  - Custom demand: 810
    Custom demand 2: 825
    Custom generation: 830
    Domestic: 720
    Domestic electric heat: 745
    Domestic low use: 705
    Large business: 310
    Large continuous: 325
    Large housing electric: 355
    Large housing standard: 365
    Large intermittent: 345
    Large off-peak: 335
    Medium business: 210
    Medium continuous: 225
    Medium housing electric: 255
    Medium housing standard: 265
    Medium intermittent: 245
    Medium off-peak: 235
    Small business: 110
    Small continuous: 125
    Small intermittent: 145
    Small off-peak: 135
    XL business: 415
    XL continuous: 420
    XL housing electric: 455
    XL housing standard: 465
    XL intermittent: 440
    XL off-peak: 430
    _column: Order
  - Custom demand: All-the-way demand
    Custom demand 2: All-the-way demand
    Custom generation: All-the-way generation
    Domestic: '(?:^|: )(?:LV Network Domestic|Domestic [UT])'
    Domestic electric heat: '(?:^|: )(?:LV Network Domestic|Domestic [UT])'
    Domestic low use: '(?:^|: )(?:LV Network Domestic|Domestic [UT])'
    Large business: '^(?:LV|LV Sub|HV|LDNO .*:) HH Metered'
    Large continuous: '^(?:LV|LV Sub|HV|LDNO .*:) HH Metered'
    Large housing electric: '^(?:LV|LV Sub|HV|LDNO .*:) HH Metered'
    Large housing standard: '^(?:LV|LV Sub|HV|LDNO .*:) HH Metered'
    Large intermittent: '^(?:LV|LV Sub|HV|LDNO .*:) HH Metered'
    Large off-peak: '^(?:LV|LV Sub|HV|LDNO .*:) HH Metered'
    Medium business: '^(?:Small|LV).*(?:Non[- ]Domestic(?: [UTN]|$)|HH Metered)'
    Medium continuous: '^(?:Small|LV).*(?:Non[- ]Domestic(?: [UTN]|$)|HH Metered)'
    Medium housing electric: '^(?:Small|LV).*(?:Non[- ]Domestic(?: [UTN]|$)|HH Metered)'
    Medium housing standard: '^(?:Small|LV).*(?:Non[- ]Domestic(?: [UTN]|$)|HH Metered)'
    Medium intermittent: '^(?:Small|LV).*(?:Non[- ]Domestic(?: [UTN]|$)|HH Metered)'
    Medium off-peak: '^(?:Small|LV).*(?:Non[- ]Domestic(?: [UTN]|$)|HH Metered)'
    Small business: '^(?:Small|LV).*Non[- ]Domestic(?: [UTN]|$)'
    Small continuous: '^(?:Small|LV).*Non[- ]Domestic(?: [UTN]|$)'
    Small intermittent: '^(?:Small|LV).*Non[- ]Domestic(?: [UTN]|$)'
    Small off-peak: '^(?:Small|LV).*Non[- ]Domestic(?: [UTN]|$)'
    XL business: '^(?:|LDNO .*: )(?:Demand Category|HV.*HH)'
    XL continuous: '^(?:|LDNO .*: )(?:Demand Category|HV.*HH)'
    XL housing electric: '^(?:|LDNO .*: )(?:Demand Category|HV.*HH)'
    XL housing standard: '^(?:|LDNO .*: )(?:Demand Category|HV.*HH)'
    XL intermittent: '^(?:|LDNO .*: )(?:Demand Category|HV.*HH)'
    XL off-peak: '^(?:|LDNO .*: )(?:Demand Category|HV.*HH)'
    _column: Tariff selection
  - Custom demand: 0
    Custom demand 2: 0
    Custom generation: 0
    Domestic: 35
    Domestic electric heat: 35
    Domestic low use: 35
    Large business: 62
    Large continuous: 0
    Large housing electric: 35
    Large housing standard: 35
    Large intermittent: 0
    Large off-peak: 93.4
    Medium business: 62
    Medium continuous: 0
    Medium housing electric: 35
    Medium housing standard: 35
    Medium intermittent: 0
    Medium off-peak: 93.4
    Small business: 62
    Small continuous: 0
    Small intermittent: 0
    Small off-peak: 93.4
    XL business: 62
    XL continuous: 0
    XL housing electric: 35
    XL housing standard: 35
    XL intermittent: 0
    XL off-peak: 93.4
    _column: Peak-time hours/week
  - Custom demand: 49
    Custom demand 2: 49
    Custom generation: 49
    Domestic: 49
    Domestic electric heat: 49
    Domestic low use: 49
    Large business: 49
    Large continuous: 49
    Large housing electric: 49
    Large housing standard: 49
    Large intermittent: 49
    Large off-peak: 49
    Medium business: 49
    Medium continuous: 49
    Medium housing electric: 49
    Medium housing standard: 49
    Medium intermittent: 49
    Medium off-peak: 49
    Small business: 49
    Small continuous: 49
    Small intermittent: 49
    Small off-peak: 49
    XL business: 49
    XL continuous: 49
    XL housing electric: 49
    XL housing standard: 49
    XL intermittent: 49
    XL off-peak: 49
    _column: Off-peak hours/week
  - Custom demand: 450
    Custom demand 2: 450
    Custom generation: 450
    Domestic: 0.649
    Domestic electric heat: 1
    Domestic low use: 0.3245
    Large business: 483
    Large continuous: 621
    Large housing electric: 100
    Large housing standard: 162.25
    Large intermittent: 276
    Large off-peak: 0
    Medium business: 48.3
    Medium continuous: 62.1
    Medium housing electric: 10
    Medium housing standard: 16.225
    Medium intermittent: 27.6
    Medium off-peak: 0
    Small business: 4.83
    Small continuous: 20.7
    Small intermittent: 2.76
    Small off-peak: 0
    XL business: 3500
    XL continuous: 4500
    XL housing electric: 1000
    XL housing standard: 1622.5
    XL intermittent: 2000
    XL off-peak: 0
    _column: Peak-time load (kW)
  - Custom demand: 450
    Custom demand 2: 450
    Custom generation: 450
    Domestic: 0.178
    Domestic electric heat: 1.2
    Domestic low use: 0.089
    Large business: 155.25
    Large continuous: 621
    Large housing electric: 120
    Large housing standard: 44.5
    Large intermittent: 276
    Large off-peak: 621
    Medium business: 15.525
    Medium continuous: 62.1
    Medium housing electric: 12
    Medium housing standard: 4.45
    Medium intermittent: 27.6
    Medium off-peak: 62.1
    Small business: 1.5525
    Small continuous: 20.7
    Small intermittent: 2.76
    Small off-peak: 6.21
    XL business: 1125
    XL continuous: 4500
    XL housing electric: 1200
    XL housing standard: 445
    XL intermittent: 2000
    XL off-peak: 4500
    _column: Off-peak load (kW)
  - Custom demand: 450
    Custom demand 2: 450
    Custom generation: 450
    Domestic: 0.333
    Domestic electric heat: 0.5
    Domestic low use: 0.1665
    Large business: 155.25
    Large continuous: 621
    Large housing electric: 50
    Large housing standard: 83.25
    Large intermittent: 276
    Large off-peak: 621
    Medium business: 15.525
    Medium continuous: 62.1
    Medium housing electric: 5
    Medium housing standard: 8.325
    Medium intermittent: 27.6
    Medium off-peak: 62.1
    Small business: 1.5525
    Small continuous: 20.7
    Small intermittent: 2.76
    Small off-peak: 6.21
    XL business: 1125
    XL continuous: 4500
    XL housing electric: 500
    XL housing standard: 832.5
    XL intermittent: 2000
    XL off-peak: 4500
    _column: Load at other times (kW)
  - Custom demand: 500
    Custom demand 2: 500
    Custom generation: 500
    Domestic: 9
    Domestic electric heat: 18
    Domestic low use: 6
    Large business: 690
    Large continuous: 690
    Large housing electric: 690
    Large housing standard: 690
    Large intermittent: 690
    Large off-peak: 690
    Medium business: 69
    Medium continuous: 69
    Medium housing electric: 69
    Medium housing standard: 69
    Medium intermittent: 69
    Medium off-peak: 69
    Small business: 23
    Small continuous: 23
    Small intermittent: 23
    Small off-peak: 23
    XL business: 5000
    XL continuous: 5000
    XL housing electric: 5000
    XL housing standard: 5000
    XL intermittent: 5000
    XL off-peak: 5000
    _column: Capacity (kVA)
