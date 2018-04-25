package CDCM;

=head Copyright licence and disclaimer

Copyright 2014-2017 Franck Latrémolière, Reckon LLP and others.

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
  - Domestic 1550: 710
    Domestic 3100: 720
    Domestic 8400: 740
    Large business: 310
    Large continuous: 320
    Large housing electric: 355
    Large housing standard: 365
    Large intermittent: 340
    Large off-peak: 330
    Medium business: 210
    Medium continuous: 225
    Medium intermittent: 245
    Medium off-peak: 235
    Other demand: 810
    Other generation: 830
    Small business: 110
    XL business: 410
    XL continuous: 420
    XL housing electric: 455
    XL housing standard: 465
    XL intermittent: 440
    XL off-peak: 430
    _column: Order
  - Domestic 1550: '(?:^|: )(?:LV Network Domestic Disabled|Domestic [UT])'
    Domestic 3100: '(?:^|: )(?:LV Network Domestic Disabled|Domestic [UT])'
    Domestic 8400: '(?:^|: )(?:LV Network Domestic Disabled|Domestic [UT])'
    Large business: '^(?:|(?:LD|Q)NO .*: )(?:LV|LV Sub|HV) HH Metered'
    Large continuous: '^(?:|(?:LD|Q)NO .*: )(?:LV|LV Sub|HV) HH Metered'
    Large housing electric: '^(?:|(?:LD|Q)NO .*: )(?:LV|LV Sub|HV) HH Metered'
    Large housing standard: '^(?:|(?:LD|Q)NO .*: )(?:LV|LV Sub|HV) HH Metered'
    Large intermittent: '^(?:|(?:LD|Q)NO .*: )(?:LV|LV Sub|HV) HH Metered'
    Large off-peak: '^(?:|(?:LD|Q)NO .*: )(?:LV|LV Sub|HV) HH Metered'
    Medium business: '^(?:Small|LV) (?:Network )?(?:Non[- ]Domestic(?: [UTN]|$)|(?:Sub )?HH Metered)'
    Medium continuous: '^(?:Small|LV) (?:Network )?(?:Non[- ]Domestic(?: [UTN]|$)|(?:Sub )?HH Metered)'
    Medium intermittent: '^(?:Small|LV) (?:Network )?(?:Non[- ]Domestic(?: [UTN]|$)|(?:Sub )?HH Metered)'
    Medium off-peak: '^(?:Small|LV) (?:Network )?(?:Non[- ]Domestic(?: [UTN]|$)|(?:Sub )?HH Metered)'
    Other demand: All-the-way metered demand
    Other generation: All-the-way generation
    Small business: '^(?:Small|LV).*Non[- ]Domestic(?: [UTN]|$)'
    XL business: '^(?:|(?:LD|Q)NO .*: )(?:Demand Category|HV.*HH)'
    XL continuous: '^(?:|(?:LD|Q)NO .*: )(?:Demand Category|HV.*HH)'
    XL housing electric: '^(?:|(?:LD|Q)NO .*: )(?:Demand Category|HV.*HH)'
    XL housing standard: '^(?:|(?:LD|Q)NO .*: )(?:Demand Category|HV.*HH)'
    XL intermittent: '^(?:|(?:LD|Q)NO .*: )(?:Demand Category|HV.*HH)'
    XL off-peak: '^(?:|(?:LD|Q)NO .*: )(?:Demand Category|HV.*HH)'
    _column: Tariff selection
  - Domestic 1550: 35
    Domestic 3100: 35
    Domestic 8400: 35
    Large business: 62
    Large continuous: 0
    Large housing electric: 35
    Large housing standard: 35
    Large intermittent: 0
    Large off-peak: 93.4
    Medium business: 62
    Medium continuous: 0
    Medium intermittent: 0
    Medium off-peak: 93.4
    Other demand: 0
    Other generation: 0
    Small business: 62
    XL business: 62
    XL continuous: 0
    XL housing electric: 35
    XL housing standard: 35
    XL intermittent: 0
    XL off-peak: 93.4
    _column: Peak-time hours/week
  - Domestic 1550: 49
    Domestic 3100: 49
    Domestic 8400: 49
    Large business: 49
    Large continuous: 49
    Large housing electric: 49
    Large housing standard: 49
    Large intermittent: 49
    Large off-peak: 49
    Medium business: 49
    Medium continuous: 49
    Medium intermittent: 49
    Medium off-peak: 49
    Other demand: 49
    Other generation: 49
    Small business: 49
    XL business: 49
    XL continuous: 49
    XL housing electric: 49
    XL housing standard: 49
    XL intermittent: 49
    XL off-peak: 49
    _column: Off-peak hours/week
  - Domestic 1550: 0.3245
    Domestic 3100: 0.649
    Domestic 8400: 1.2
    Large business: 483
    Large continuous: 621
    Large housing electric: 120
    Large housing standard: 162.25
    Large intermittent: 276
    Large off-peak: 0
    Medium business: 48.3
    Medium continuous: 62.1
    Medium intermittent: 27.6
    Medium off-peak: 0
    Other demand: 450
    Other generation: 10000
    Small business: 4.83
    XL business: 3500
    XL continuous: 4500
    XL housing electric: 1200
    XL housing standard: 1622.5
    XL intermittent: 2000
    XL off-peak: 0
    _column: Peak-time load (kW)
  - Domestic 1550: 0.089
    Domestic 3100: 0.178
    Domestic 8400: 1.4
    Large business: 155.25
    Large continuous: 621
    Large housing electric: 140
    Large housing standard: 44.5
    Large intermittent: 276
    Large off-peak: 621
    Medium business: 15.525
    Medium continuous: 62.1
    Medium intermittent: 27.6
    Medium off-peak: 62.1
    Other demand: 450
    Other generation: 10000
    Small business: 1.5525
    XL business: 1125
    XL continuous: 4500
    XL housing electric: 1400
    XL housing standard: 445
    XL intermittent: 2000
    XL off-peak: 4500
    _column: Off-peak load (kW)
  - Domestic 1550: 0.1665
    Domestic 3100: 0.333
    Domestic 8400: 0.6
    Large business: 155.25
    Large continuous: 621
    Large housing electric: 60
    Large housing standard: 83.25
    Large intermittent: 276
    Large off-peak: 621
    Medium business: 15.525
    Medium continuous: 62.1
    Medium intermittent: 27.6
    Medium off-peak: 62.1
    Other demand: 450
    Other generation: 10000
    Small business: 1.5525
    XL business: 1125
    XL continuous: 4500
    XL housing electric: 600
    XL housing standard: 832.5
    XL intermittent: 2000
    XL off-peak: 4500
    _column: Load at other times (kW)
  - Domestic 1550: 6
    Domestic 3100: 9
    Domestic 8400: 18
    Large business: 690
    Large continuous: 690
    Large housing electric: 690
    Large housing standard: 690
    Large intermittent: 690
    Large off-peak: 690
    Medium business: 69
    Medium continuous: 69
    Medium intermittent: 69
    Medium off-peak: 69
    Other demand: 500
    Other generation: 10000
    Small business: 23
    XL business: 5000
    XL continuous: 5000
    XL housing electric: 5000
    XL housing standard: 5000
    XL intermittent: 5000
    XL off-peak: 5000
    _column: Capacity (kVA)
