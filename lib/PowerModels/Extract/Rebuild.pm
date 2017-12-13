package PowerModels::Extract::Rebuild;

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
use File::Glob qw(bsd_glob);
use File::Spec::Functions qw(abs2rel catdir catfile);
use Encode qw(decode_utf8);

sub rebuildWriter {
    my ( $arg, $runner ) = @_;
    sub {
        my ( $fileName, $workbook ) = @_;
        die unless $fileName;
        my ( $path, $core, $ext ) = $fileName =~ m#(.*/)?([^/]+)(\.xlsx?)$#is;
        $path = '' unless defined $path;
        my $tempFolder = $path . $core . '-' . $$ . '.tmp';
        my $sidecar    = $path . $core;
        mkdir $sidecar unless -e $sidecar;
        undef $sidecar unless -d $sidecar && -w _;

        unless ( defined $sidecar ) {
            $sidecar = $path . '~$' . $core;
            undef $sidecar unless -d $sidecar && -w _;
        }
        mkdir $tempFolder or die "Cannot create $tempFolder: $!";
        my $rulesFile;
        if ( defined $sidecar ) {
            $rulesFile = "$sidecar/%-$core.yml";
            undef $rulesFile unless -f $rulesFile;
            unless ( defined $rulesFile ) {
                $rulesFile = "$sidecar/%.yml";
                undef $rulesFile unless -f $rulesFile;
            }
        }
        require YAML;
        {
            my ( $h1, $h2 );
            open $h1, '>', "$tempFolder/index-$core.yml";
            binmode $h1, ':utf8';
            unless ( defined $rulesFile ) {
                open $h2, '>', $rulesFile = "$tempFolder/%-$core.yml";
                binmode $h2, ':utf8';
            }
            require PowerModels::Rules::FromWorkbook;
            foreach ( PowerModels::Rules::FromWorkbook::extractYaml($workbook) )
            {
                print $h1 $_;
                if ($h2) {
                    my $rules = YAML::Load($_);
                    delete $rules->{$_}
                      foreach 'template', grep { /^~/s } keys %$rules;
                    print {$h2} YAML::Dump($rules);
                }
            }
        }
        require PowerModels::Extract::InputTables;
        my %trees =
          PowerModels::Extract::InputTables::extractInputData($workbook);
        while ( my ( $key, $value ) = each %trees ) {
            open my $h, '>', "$tempFolder/$core$key.yml";
            binmode $h, ':utf8';
            print $h YAML::Dump($value);
        }
        $runner->makeModels(
            '-pickall',    '-single',
            '-purple',     "-folder=$tempFolder",
            "-template=%", lc $ext eq '.xls' ? '-xls' : '-xlsx',
            $rulesFile,    "$tempFolder/$core.yml"
        );
        if ( -s "$tempFolder/$core$ext" ) {
            rename "$path$core$ext", "$tempFolder/$core-old$ext"
              unless defined $sidecar;
            rename "$tempFolder/$core$ext", "$path$core$ext"
              or warn "Cannot move $tempFolder/$core$ext to $path$core$ext: $!";
        }
        if ( defined $sidecar ) {
            my $dh;
            opendir $dh, $tempFolder;
            foreach ( readdir $dh ) {
                next if /^\.\.?$/s;
                rename "$tempFolder/$_", "$sidecar/$_";
            }
            closedir $dh;
            rmdir $tempFolder;
        }
        else {
            rename $tempFolder, $path . "Z_Rebuild-$core-$$";
        }
      }
}

1;
