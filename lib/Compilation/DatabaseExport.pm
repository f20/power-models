package Compilation;

=head Copyright licence and disclaimer

Copyright 2009-2012 Reckon LLP and others.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice,
this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation
and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

=cut

use warnings;
use strict;
use utf8;

sub new {
    require DBI;
    my $databaseHandle = DBI->connect( 'dbi:SQLite:dbname=~$database.sqlite',
        '', '', { sqlite_unicode => 1, AutoCommit => 0, } )
      or die "Cannot open sqlite database: $!";
    bless \$databaseHandle, shift;
}

sub listModels {
    my ($self) = @_;
    my $findCo = $$self->prepare('select bid, filename from books');
    $findCo->execute;
    my @models;
    while ( my ( $bid, $filename ) = $findCo->fetchrow_array ) {
        next unless $filename =~ s/\.xlsx?$//is;
        local $_ = $filename;
        s/^M-//;
        s/^CE-NEDL/NPG-Northeast/;
        s/^CE-YEDL/NPG-Yorkshire/;
        s/^CN-East/WPD-EastM/;
        s/^CN-West/WPD-WestM/;
        s/^EDFEN/UKPN/;
        s/^NP-/NPG-/;
        s/^SP-/SPEN-/;
        s/^SSE-/SSEPD-/;
        s/^WPD-Wales/WPD-SWales/;
        s/^WPD-West\b/WPD-SWest/;
        my @a = /^(.+?)(-20[0-9]{2}-[0-9]{2})(-.*)$/s;
        @a = /^(.+?)(-20[0-9]{2}-[0-9]{2})?(-[^-]*)?$/s unless @a;
        push @models, [
            $bid,
            $filename,
            $_,
            map {
                local $_ = $_;
                tr/-/ /;
                s/^ //;
                $_;
            } grep { $_ } @a
        ];
    }
    sort { $a->[2] cmp $b->[2]; } @models;
}

sub DESTROY {
    ${ $_[0] }->disconnect;
}

1;
