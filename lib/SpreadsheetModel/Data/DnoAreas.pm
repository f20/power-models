package SpreadsheetModel::Data::DnoAreas;

# Copyright 2010-2019 Reckon LLP and others.
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

sub normaliseDnoName {
    local @_ = @_ if defined wantarray;
    @_ = ('Untitled') unless @_;
    foreach (@_) {
        s/^CE-NEDL/NPG-Northeast/;
        s/^CE-YEDL/NPG-Yorkshire/;
        s/^CN-East/WPD-EastM/;
        s/^CN-West/WPD-WestM/;
        s/^EDFEN-/UKPN-/;
        s/^NP-/NPG-/;
        s/^SP-/SPEN-/;
        s/^SSE(PD)?-/SSEN-/;
        s/^WPD-Wales/WPD-SWales/;
        s/^WPD-West\b/WPD-SWest/;
    }
    wantarray ? @_ : $_[0];
}

sub dnoShortNames {
    (
        'ENWL',
        'NPG Northeast',
        'NPG Yorkshire',
        'SPEN SPD',
        'SPEN SPM',
        'SSEN SEPD',
        'SSEN SHEPD',
        'UKPN EPN',
        'UKPN LPN',
        'UKPN SPN',
        'WPD EastM',
        'WPD SWales',
        'WPD SWest',
        'WPD WestM',
    );
}

sub dnoLongNames {
    (
        'Electricity North West Limited',
        'Northern Powergrid Northeast',
        'Northern Powergrid Yorkshire',
        'SP Distribution',
        'SP Manweb',
        'Southern Electric Power Distribution',
        'Scottish Hydro Electric Power Distribution',
        'Eastern Power Networks',
        'London Power Networks',
        'South Eastern Power Networks',
        'WPD East Midlands',
        'WPD South Wales',
        'WPD South West',
        'WPD West Midlands',
    );
}

1;
