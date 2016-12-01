package EDCM2::PickBest;

use warnings;
use strict;
use POSIX;

sub score {

    my ( $class, $rule, $metadata ) = @_;
    my $month = $metadata->[1] || strftime( '%Y-%m', localtime );
    $month =~ s/^-*//;
    my $score = 0;

    # DCP 185
    $score += 100 if $rule->{dcp185} xor $month lt '2014-10';

    # DCP 189
    $score += 100 if $rule->{dcp189} xor $month lt '2014-10';

    # DCP 161
    $score += 100 if $rule->{dcp161} xor $month lt '2017-10';

    0
      and warn join ' ', $rule->{nickName} || $rule->{'.'} || $rule, $month,
      $score;

    $score;

}

1;
