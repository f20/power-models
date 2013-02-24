#!/usr/bin/env perl
use common::sense;
use YAML;
my %c;
for my $rules (@ARGV) {
    open my $f, '<', $rules or next;
    local undef $/;
    binmode $f, ':utf8';
    my $d = Load(<$f>);
    while ( my ( $k, $v ) = each %$d ) {
        $v = Dump($v) if ref $v;
        push @{ $c{$k}{$v} }, $rules;
    }
}
binmode STDOUT, ':utf8';
print Dump( \%c );
