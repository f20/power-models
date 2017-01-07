package ModelM::PickBest;

=head Copyright licence and disclaimer

Copyright 2012-2017 Franck Latrémolière, Reckon LLP and others.

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
use POSIX;

sub score {

    my ( $class, $rule, $month ) = @_;
    $month ||= strftime( '%Y-%m', localtime );

    my $score = 0;

    # DCP 071
    $score += 100 if $rule->{dcp071} xor $month lt '2011-04';

    # DCP 095
    $score += 100 if $rule->{dcp095} xor $month lt '2011-10';

    # DCP 096
    $score += 100 if $rule->{dcp096} xor $month lt '2011-10';

    # DCP 118
    $score += 100 if $rule->{dcp118} && !$rule->{edcm} xor $month lt '2013-10';

    # Units input data
    $score += 40 if $rule->{calcUnits} xor $month lt '2014-01';

    # MEAV input data
    $score += 40 if $rule->{meav} xor $month lt '2014-01';

    # Net capex input data
    $score += 40 if $rule->{netCapex} xor $month lt '2014-01';

    # DCP 117 and DCP 231
    $score += 100 if $rule->{dcp117} xor $month lt '2015-10';
    $score += 50  if !$rule->{dcp117dcp118interaction} xor $month lt '2017-10';
    $score += 50  if !$rule->{dcp117weirdness} xor $month lt '2017-10';

    0
      and warn join ' ', $rule->{nickName} || $rule->{'.'} || $rule, $month,
      $score;

    $score;

}

1;
