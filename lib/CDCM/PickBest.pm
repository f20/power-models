package CDCM::PickBest;

=head Copyright licence and disclaimer

Copyright 2012-2018 Franck Latrémolière, Reckon LLP and others.

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

sub score {

    my ( $class, $rule, $month ) = @_;

    my $score = 0;

    # DCP 130
    $score += 10 if $rule->{tariffs} =~ /dcp130/i xor $month lt '2012-03';

    # DCP 132 and DCP 249
    $score += 10
      if $rule->{targetRevenue}
      && $rule->{targetRevenue} =~ /dcp132|dcp249/i xor $month lt '2012-03';

    # DCP 163
    $score += 10 if $rule->{tariffs} =~ /dcp163/i xor $month lt '2013-03';

    # Bungs
    $score += 10
      if $rule->{bung}
      and $month gt '2013-03' && $month le '2016-03'
      || $month ge '2018-02'  && $month le '2018-03';

    # Fiddle
    $score += 10 if $rule->{fiddle} and $month eq '2019-02';

    # DCP 179
    $score += 10 if $rule->{tariffs} =~ /pc34hh/i xor $month lt '2014-03';

    # DCP 227
    $score += 10
      if $rule->{agghhequalisation}
      && $rule->{agghhequalisation} =~ /rag/i xor $month lt '2016-03';

    # DCP 161
    $score += 10
      if $rule->{unauth}
      && $rule->{unauth} =~ /dayotex/i xor $month lt '2017-03';

    # DCP 249
    $score += 1
      if $rule->{targetRevenue}
      && $rule->{targetRevenue} =~ /dcp249/i xor $month lt '2017-10';

    # Otnei
    $score += 1 if $rule->{lvDiversityWrong} xor $month lt '2019-10';

    # DCP 268 avoidance
    $score *= 0.1 if $rule->{tariffGrouping};

    $score;

}

1;
