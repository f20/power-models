package PowerModels::Extract::Json;

=head Copyright licence and disclaimer

Copyright 2011-2017 Franck Latrémolière and others. All rights reserved.

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
use Encode qw(decode_utf8);
use YAML;

my $jsonMachine;

sub jsonMachineMaker {
    return $jsonMachine if $jsonMachine;
    foreach (qw(JSON JSON::PP)) {
        return $jsonMachine = $_->new
          if eval "require $_";
    }
    die 'No JSON module';
}

sub jsonWriter {
    my ($arg) = @_;
    my $jsonMachine = jsonMachineMaker()->canonical(1)->utf8;
    my $options = { $arg =~ /min/i ? ( minimum => 1 ) : (), };
    sub {
        my ( $book, $workbook ) = @_;
        die unless $book;
        my $file = $book;
        $file =~ s/\.xl[a-z]+?$//is;
        my $tree;
        if ( -e $file ) {
            open my $h, '<', "$file.json";
            binmode $h;
            local undef $/;
            $tree = $jsonMachine->decode(<$h>);
        }
        require PowerModels::Extract::InputTables;
        my %trees =
          PowerModels::Extract::InputTables::extractInputData( $workbook,
            $tree, $options );
        while ( my ( $key, $value ) = each %trees ) {
            open my $h, '>', "$file$key.json";
            binmode $h;
            print {$h} $jsonMachine->encode($value);
        }
    };
}

1;
