#!/usr/bin/perl

=head Copyright licence and disclaimer

Copyright 2012 Reckon LLP and contributors. All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted
provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of
conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of
conditions and the following disclaimer in the documentation and/or other materials provided
with the distribution.

THIS SOFTWARE IS PROVIDED BY RECKON LLP AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED
WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL RECKON LLP OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;
OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT,
STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

=cut

use warnings;
use strict;
use utf8;
use YAML;
use File::Spec::Functions qw(rel2abs);
use File::Basename qw(dirname);
use Parse::Stata::DtaReader;

my %d;

foreach (@ARGV) {
    unless (/\.dta$/is) {
        warn "Ignored: $_";
        next;
    }
    open my $handle, '<', $_ or do {
        warn "Cannot open $_: $!";
        next;
    };
    warn "Reading $_\n";
    my $dta = Parse::Stata::DtaReader->new($handle);
    my ( @table, @column );
    for ( my $i = 1 ; $i < $dta->{nvar} ; ++$i ) {
        if ( $dta->{varlist}[$i] =~ /t([0-9]+)c([0-9]+)/ ) {
            $table[$i]  = $1;
            $column[$i] = $2;
        }
    }
    while ( my @row = $dta->readRow ) {
        my $book = $row[0];
        my $line = $table[1] ? 'Single-line CSV' : $row[1];
        $d{$book}{ $table[$_] }[ $column[$_] ]{$line} = $row[$_]
          foreach grep { $table[$_] } 1 .. $#table;
    }
}

while ( my ( $book, $data ) = each %d ) {
    my $yml = "$book.yml";
    $yml =~ s/(?:-LRIC|-LRICsplit|-FCP)?(-r[0-9]+)?\.yml$/.yml/is;
    warn "Writing $book data to $yml\n";
    open my $h, '>', $yml;
    binmode $h, ':utf8';
    print {$h} Dump $data;
}
