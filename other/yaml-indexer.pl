#!/usr/bin/env perl
use warnings;
use strict;
use utf8;
binmode STDIN,  ':utf8';
binmode STDOUT, ':utf8';
binmode STDERR, ':utf8';
use File::Spec::Functions qw(rel2abs catdir);
use File::Basename 'dirname';
my ( $homedir, $perl5dir );

BEGIN {
    $homedir = dirname( rel2abs( -l $0 ? ( readlink $0, dirname $0) : $0 ) );
    while (1) {
        $perl5dir = catdir( $homedir, 'lib' );
        last if -d catdir( $perl5dir, 'SpreadsheetModel' );
        my $parent = dirname $homedir;
        last if $parent eq $homedir;
        $homedir = $parent;
    }
}
use lib catdir( $homedir, 'cpan' ), $perl5dir;
use YAML qw(LoadFile DumpFile Dump);
use Digest::SHA qw(sha1_hex);
use Encode qw(encode_utf8);

chdir $homedir;
open LIST, q^find . -name '*.y*ml' |^;
my %obj;
while (<LIST>) {
    chomp;
    next if m#^\./map#;
    next if 1 && m#/Data-[0-9]{4}-[0-9]{2}#;
    my @a = LoadFile($_);
    if ( @a == 0 ) {
        warn "$_ contains no objects\n";
    }
    elsif ( @a == 1 ) {
        $obj{$_} = $a[0];
    }
    else {
        for ( my $no = 0 ; $no < @a ; ++$no ) { $obj{"$_/$no"} = $a[$no]; }
    }
}

my %map;
while ( my ( $f, $o ) = each %obj ) {
    unless ( ref $o eq 'HASH' ) {
        warn "$f is not a HASH";
        next;
    }
    my $cat = $o->{PerlModule} || '';
    next if 1 && !$cat;
    while ( my ( $k, $v ) = each %$o ) {
        if ( ref $v ) {
            my $hv = join '#', ref $v, sha1_hex( encode_utf8( Dump($v) ) );
            $map{$cat}{$k}{$hv}[0] = $v;
            $v = $hv;
        }
        push @{ $map{$cat}{$k}{$v} }, $f;
    }
}

while ( my ( $k, $v ) = each %map ) {
    DumpFile( $$, $v );
    rename $$, "map-full-$k.yml";
    DumpFile(
        $$,
        {
            map { ( $_ => [ keys %{ $v->{$_} } ] ) }
            grep { $_ ne '.' } keys %$v
        }
    );
    rename $$, "map-short-$k.yml";
}
