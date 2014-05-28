use warnings;
use strict;
use YAML qw(LoadFile DumpFile);
my $srcdir = '../Data-2014-02';
opendir D, $srcdir;
while ( my $file = readdir D ) {
    next if $file =~ /^\./;
    my ($s) = LoadFile("$srcdir/$file");
    my @o;
    foreach my $t ( grep { $_ eq '1061' } keys %$s ) {
        for ( my $c = 1 ; $c < @{ $s->{$t} } ; ++$c ) {
            foreach my $r ( grep { /Unrestricted/ } keys %{ $s->{$t}[$c] } ) {
                push @o, [ $t, $c, $r, $s->{$t}[$c]{$r} ];
            }
        }
    }
    for my $folder ( map { "Data-$_"; } qw(2015-02 2015-02alt) ) {
        my ($d) = LoadFile("$folder/$file") or next;
        $d->{ $_->[0] }[ $_->[1] ]{ $_->[2] } = $_->[3] foreach @o;
        DumpFile( "$folder/$file", $d );
    }
}
