package Ancillary::DnoAreas;

=head Copyright licence and disclaimer

Copyright 2010-2014 Reckon LLP and others.

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

sub normaliseDnoName {
    local @_ = @_ if defined wantarray;
    foreach (@_) {
        s/^CE-NEDL/NPG-Northeast/;
        s/^CE-YEDL/NPG-Yorkshire/;
        s/^CN-East/WPD-EastM/;
        s/^CN-West/WPD-WestM/;
        s/^EDFEN-/UKPN-/;
        s/^NP-/NPG-/;
        s/^SP-/SPEN-/;
        s/^SSE-/SSEPD-/;
        s/^WPD-Wales/WPD-SWales/;
        s/^WPD-West\b/WPD-SWest/;
    }
    wantarray ? @_ : $_[0];
}

1;
