package ModelM::PickBest;

use warnings;
use strict;
use POSIX;

sub score {

    my ( $class, $rule, $metadata ) = @_;
    my $month = $metadata->[1] || strftime( '%Y-%m', localtime );
    $month =~ s/^-*//;
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
