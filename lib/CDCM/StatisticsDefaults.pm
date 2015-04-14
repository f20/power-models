package CDCM;

=head Copyright licence and disclaimer

Copyright 2014-2015 Franck Latrémolière, Reckon LLP and others.

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
    $table1202[0]{ $_[1] };
}

1;

__DATA__
---
1202:
  - _table: 1202. Consumption assumptions for illustrative customers
  - Other demand 1: 815
    Other demand 2: 825
    Other generation: 835
    Domestic electric heat: 740
    Domestic low use: 700
    Domestic standard: 720
    Large business: 310
    Large continuous: 320
    Large housing electric: 355
    Large housing standard: 365
    Large intermittent: 345
    Large off-peak: 330
    Medium business: 210
    Medium continuous: 220
    Medium housing electric: 255
    Medium housing standard: 265
    Medium intermittent: 245
    Medium off-peak: 230
    Small business: 110
    Small continuous: 120
    Small intermittent: 145
    Small off-peak: 130
    XL business: 410
    XL continuous: 420
    XL housing electric: 455
    XL housing standard: 465
    XL intermittent: 445
    XL off-peak: 430
    _column: Order
  - Other demand 1: All-the-way demand
    Other demand 2: All-the-way demand
    Other generation: All-the-way generation
    Domestic electric heat: '(?:^|: )(?:LV Network Domestic|Domestic [UTN])'
    Domestic low use: '(?:^|: )(?:LV Network Domestic|Domestic [UTN])'
    Domestic standard: '(?:^|: )(?:LV Network Domestic|Domestic [UTN])'
    Large business: '^(?:LV|LV Sub|HV|LDNO .*:) HH Metered$'
    Large continuous: '^(?:LV|LV Sub|HV|LDNO .*:) HH Metered$'
    Large housing electric: '^(?:LV|LV Sub|HV|LDNO .*:) HH Metered$'
    Large housing standard: '^(?:LV|LV Sub|HV|LDNO .*:) HH Metered$'
    Large intermittent: '^(?:LV|LV Sub|HV|LDNO .*:) HH Metered$'
    Large off-peak: '^(?:LV|LV Sub|HV|LDNO .*:) HH Metered$'
    Medium business: '^(?:Small|LV).*(?:Non[- ]Domestic(?: [UTN]|$)|HH Metered$)'
    Medium continuous: '^(?:Small|LV).*(?:Non[- ]Domestic(?: [UTN]|$)|HH Metered$)'
    Medium housing electric: '^(?:Small|LV).*(?:Non[- ]Domestic(?: [UTN]|$)|HH Metered$)'
    Medium housing standard: '^(?:Small|LV).*(?:Non[- ]Domestic(?: [UTN]|$)|HH Metered$)'
    Medium intermittent: '^(?:Small|LV).*(?:Non[- ]Domestic(?: [UTN]|$)|HH Metered$)'
    Medium off-peak: '^(?:Small|LV).*(?:Non[- ]Domestic(?: [UTN]|$)|HH Metered$)'
    Small business: '^(?:Small|LV).*Non[- ]Domestic(?: [UTN]|$)'
    Small continuous: '^(?:Small|LV).*Non[- ]Domestic(?: [UTN]|$)'
    Small intermittent: '^(?:Small|LV).*Non[- ]Domestic(?: [UTN]|$)'
    Small off-peak: '^(?:Small|LV).*Non[- ]Domestic(?: [UTN]|$)'
    XL business: '^(?:|LDNO .*: )HV HH Metered$'
    XL continuous: '^(?:|LDNO .*: )HV HH Metered$'
    XL housing electric: '^(?:|LDNO .*: )HV HH Metered$'
    XL housing standard: '^(?:|LDNO .*: )HV HH Metered$'
    XL intermittent: '^(?:|LDNO .*: )HV HH Metered$'
    XL off-peak: '^(?:|LDNO .*: )HV HH Metered$'
    _column: Tariff selection
  - Other demand 1: 0
    Other demand 2: 0
    Other generation: 0
    Domestic electric heat: 35
    Domestic low use: 35
    Domestic standard: 35
    Large business: 66
    Large continuous: 0
    Large housing electric: 35
    Large housing standard: 35
    Large intermittent: 0
    Large off-peak: 0
    Medium business: 66
    Medium continuous: 0
    Medium housing electric: 35
    Medium housing standard: 35
    Medium intermittent: 0
    Medium off-peak: 0
    Small business: 66
    Small continuous: 0
    Small intermittent: 0
    Small off-peak: 0
    XL business: 66
    XL continuous: 0
    XL housing electric: 35
    XL housing standard: 35
    XL intermittent: 0
    XL off-peak: 0
    _column: Peak-time hours/week
  - Other demand 1: 0
    Other demand 2: 0
    Other generation: 0
    Domestic electric heat: 49
    Domestic low use: 49
    Domestic standard: 49
    Large business: 0
    Large continuous: 0
    Large housing electric: 49
    Large housing standard: 49
    Large intermittent: 0
    Large off-peak: 74.6666666666667
    Medium business: 0
    Medium continuous: 0
    Medium housing electric: 49
    Medium housing standard: 49
    Medium intermittent: 0
    Medium off-peak: 74.6666666666667
    Small business: 0
    Small continuous: 0
    Small intermittent: 0
    Small off-peak: 74.6666666666667
    XL business: 0
    XL continuous: 0
    XL housing electric: 49
    XL housing standard: 49
    XL intermittent: 0
    XL off-peak: 74.6666666666667
    _column: Off-peak hours/week
  - Other demand 1: 0
    Other demand 2: 0
    Other generation: 0
    Domestic electric heat: 1.1
    Domestic low use: 0.4
    Domestic standard: 0.8
    Large business: 350
    Large continuous: 0
    Large housing electric: 110
    Large housing standard: 200
    Large intermittent: 0
    Large off-peak: 0
    Medium business: 48.3
    Medium continuous: 0
    Medium housing electric: 11
    Medium housing standard: 20
    Medium intermittent: 0
    Medium off-peak: 0
    Small business: 16.1
    Small continuous: 0
    Small intermittent: 0
    Small off-peak: 0
    XL business: 3500
    XL continuous: 0
    XL housing electric: 1100
    XL housing standard: 2000
    XL intermittent: 0
    XL off-peak: 0
    _column: Peak-time load (kW)
  - Other demand 1: 0
    Other demand 2: 0
    Other generation: 0
    Domestic electric heat: 1.6
    Domestic low use: 0.125
    Domestic standard: 0.25
    Large business: 0
    Large continuous: 0
    Large housing electric: 160
    Large housing standard: 62.5
    Large intermittent: 0
    Large off-peak: 450
    Medium business: 0
    Medium continuous: 0
    Medium housing electric: 16
    Medium housing standard: 6.25
    Medium intermittent: 0
    Medium off-peak: 62.1
    Small business: 0
    Small continuous: 0
    Small intermittent: 0
    Small off-peak: 20.7
    XL business: 0
    XL continuous: 0
    XL housing electric: 1600
    XL housing standard: 625
    XL intermittent: 0
    XL off-peak: 4500
    _column: Off-peak load (kW)
  - Other demand 1: 0.5
    Other demand 2: 0.5
    Other generation: 0.5
    Domestic electric heat: 0.434817351598174
    Domestic low use: 0.194206621004566
    Domestic standard: 0.388413242009132
    Large business: 102.941176470588
    Large continuous: 450
    Large housing electric: 43.4817351598174
    Large housing standard: 97.1033105022831
    Large intermittent: 200
    Large off-peak: 0
    Medium business: 14.2058823529412
    Medium continuous: 62.1
    Medium housing electric: 4.34817351598174
    Medium housing standard: 9.71033105022831
    Medium intermittent: 27.6
    Medium off-peak: 0
    Small business: 4.73529411764706
    Small continuous: 20.7
    Small intermittent: 9.2
    Small off-peak: 0
    XL business: 1029.41176470588
    XL continuous: 4500
    XL housing electric: 434.817351598174
    XL housing standard: 971.033105022831
    XL intermittent: 2000
    XL off-peak: 0
    _column: Load at other times (kW)
  - Other demand 1: 0
    Other demand 2: 0
    Other generation: 0
    Domestic electric heat: 18
    Domestic low use: 6
    Domestic standard: 9
    Large business: 500
    Large continuous: 500
    Large housing electric: 500
    Large housing standard: 500
    Large intermittent: 500
    Large off-peak: 500
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
1202simple:
  - _table: 1202. Consumption assumptions for illustrative customers
  - Domestic low use: 1700
    Domestic standard: 1720
    Domestic high use: 1740
    _column: Order
  - Domestic low use: '^(?:LV Network Domestic|Domestic [UT])'
    Domestic standard: '^(?:LV Network Domestic|Domestic [UT])'
    Domestic high use: '^(?:LV Network Domestic|Domestic [UT])'
    _column: Tariff selection
  - Domestic low use: 35
    Domestic standard: 35
    Domestic high use: 35
    _column: Peak-time hours/week
  - Domestic low use: 49
    Domestic standard: 49
    Domestic high use: 49
    _column: Off-peak hours/week
  - Domestic low use: 0.4
    Domestic standard: 0.8
    Domestic high use: 1.313
    _column: Peak-time load (kW)
  - Domestic low use: 0.125
    Domestic standard: 0.25
    Domestic high use: 0.75
    _column: Off-peak load (kW)
  - Domestic low use: 0.194206621
    Domestic standard: 0.388413242
    Domestic high use: 0.75
    _column: Load at other times (kW)
  - Domestic low use: 6
    Domestic standard: 9
    Domestic high use: 12
    _column: Capacity (kVA)
