package EDCM2::PickBest;

use warnings;
use strict;
use POSIX;

sub score {

    my ( $class, $rule, $metadata ) = @_;
    my $month = $metadata->[1] || strftime( '%Y-%m', localtime );
    my $score = 0;

    # DCP 185
    $score += 100 if $rule->{dcp185} xor $month lt '2014-10';

    warn join ' ', $rule->{nickName} || $rule->{'.'} || $rule, $month, $score;

    $score;

}

1;