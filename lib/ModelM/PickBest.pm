package ModelM::PickBest;

use warnings;
use strict;
use POSIX;

sub score {

    my ( $class, $rule, $metadata ) = @_;
    my $month = $metadata->[1] || strftime( '%Y-%m', localtime );
    my $score = 0;

    # DCP 071
    $score += 100 if $rule->{dcp071} xor $month lt '2011-04';

    # DCP 095
    $score += 100 if $rule->{dcp095} xor $month lt '2011-10';

    # DCP 096
    $score += 100 if $rule->{dcp096} xor $month lt '2011-10';

    # DCP 118
    $score += 100 if $rule->{dcp118} && !$rule->{edcm} xor $month lt '2013-10';

    0
      and warn join ' ', $rule->{nickName} || $rule->{'.'} || $rule, $month,
      $score;

    $score;

}

1;
