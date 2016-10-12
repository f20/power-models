package CDCM::PickBest;

=head Copyright licence and disclaimer

Copyright 2012-2016 Franck Latrémolière, Reckon LLP and others.

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
use POSIX;

sub score {

    my ( $class, $rule, $metadata ) = @_;
    my $month = $metadata->[1] || strftime( '%Y-%m-est', localtime );
    $month =~ s/^-*//;
    my $score = 0;

    # DCP 130
    $score += 100 if $rule->{tariffs} =~ /dcp130/i xor $month lt '2012-03';

    # DCP 161
    $score += 100
      if $rule->{unauth}
      && $rule->{unauth} =~ /dayotex/i xor $month lt '2017-03';

    # DCP 163
    $score += 50 if $rule->{tariffs} =~ /dcp163/i xor $month lt '2013-03';

    # DCP 179
    $score += 100 if $rule->{tariffs} =~ /pc34hh/i xor $month lt '2014-03';

    # DCP 227
    $score += 100
      if $rule->{agghhequalisation}
      && $rule->{agghhequalisation} =~ /rag/i xor $month lt '2016-03';

    # Bung
    $score += 10
      if $rule->{electionBung} && $month gt '2013-03' && $month lt '2016-03';

    # DCP 268
    $score += 900 if !$rule->{tariffGrouping};

    # Fun
    $score += 900 if !$rule->{pcd} xor $month lt '2020-03';

    0
      and warn join ' ', $rule->{nickName} || $rule->{'.'} || $rule, $month,
      $score;

    $score;

}

1;
