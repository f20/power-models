package CDCM::Scoring;

use warnings;
use strict;
use POSIX;

sub score {
    my ( $class, $rule, $metadata ) = @_;
    my $month = $metadata->[1] || strftime( '%Y-%m', localtime );
    my $score = 0;
    $score += 100 if $rule->{tariffs} =~ /dcp130/i xor $month lt '2012-10';
    $score += 50  if $rule->{tariffs} =~ /dcp163/i xor $month lt '2013-10';
    0
      and warn(
        ( $rule->{nickName} || $rule->{'.'} || $rule ) . " $month: $score" );
    $score;
}

1;
