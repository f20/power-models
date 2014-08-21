#!/usr/bin/env perl
use warnings;
use strict;
use YAML qw(LoadFile DumpFile);
my $srcFolder = '../Data-2014-02';
opendir SRCDHANDLE, $srcFolder;
while ( my $file = readdir SRCDHANDLE ) {
    next if $file =~ /^\./;
    my ($srcData) = LoadFile("$srcFolder/$file");
    my @dataItem;
    foreach my $tab ( grep { $_ eq '1061' } keys %$srcData ) {
        for ( my $col = 1 ; $col < @{ $srcData->{$tab} } ; ++$col ) {
            foreach my $row ( grep { /Unrestricted/ }
                keys %{ $srcData->{$tab}[$col] } )
            {
                push @dataItem,
                  [ $tab, $col, $row, $srcData->{$tab}[$col]{$row} ];
            }
        }
    }
    for my $dstFolder ( map { "Data-$_"; } qw(2015-02) ) {
        my ($dstData) = LoadFile("$dstFolder/$file") or next;
        $dstData->{ $_->[0] }[ $_->[1] ]{ $_->[2] } = $_->[3] foreach @dataItem;
        DumpFile( "$dstFolder/$file", $dstData );
    }
}
