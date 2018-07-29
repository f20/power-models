package PowerModels::Extract::Yaml;

=head Copyright licence and disclaimer

Copyright 2008-2017 Franck Latrémolière, Reckon LLP and others.

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
use Encode;

sub ymlWriter {
    my ($arg) = @_;
    my $options = { $arg =~ /min/i ? ( minimum => 1 ) : (), };
    my $sheetFilter;
    $sheetFilter = $1 if $arg =~ /sheet(.+)/;
    sub {
        my ( $book, $workbook ) = @_;
        die unless $book;
        my $file = $book;
        $file =~ s/\.xl[a-z]+?$//is;
        my $tree;
        require YAML;
        if ( my ($oldYaml) = grep { -f $_; } "$file.yml", "$file.yaml" ) {
            open my $h, '<', $oldYaml;
            binmode $h, ':utf8';
            local undef $/;
            $tree = YAML::Load(<$h>);
        }
        require PowerModels::Extract::InputTables;
        my %trees =
          PowerModels::Extract::InputTables::extractInputData( $workbook,
            $tree, $options );
        while ( my ( $key, $value ) = each %trees ) {
            open my $h, '>', "$file$key.yml";
            binmode $h, ':utf8';
            print $h YAML::Dump($value);
        }
      }, $sheetFilter
      ? (
        CellHandler => sub {
            my ( $workbook, $sheetIdx, $row, $col, $cell ) = @_;
            $workbook->worksheet($sheetIdx)->get_name ne $sheetFilter;
        }
      )
      : ();
}

1;

