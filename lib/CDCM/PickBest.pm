package CDCM::PickBest;

use warnings;
use strict;
use POSIX;

sub score {

    my ( $class, $rule, $metadata ) = @_;
    my $month = $metadata->[1] || strftime( '%Y-%m', localtime );
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

    # Fun
    $score += 999 if !$rule->{pcd} xor $month lt '2017-03';

    0
      and warn join ' ', $rule->{nickName} || $rule->{'.'} || $rule, $month,
      $score;

    $score;

}

1;
