package SpreadsheetModel::Data::DataTools;

=head Copyright licence and disclaimer

Copyright 2011-2015 Franck Latrémolière and others. All rights reserved.

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

sub ymlDiff {

    my ( @names, @src, %stream );
    foreach my $fileName ( map { decode_utf8 $_} @_ ) {
        if ( $fileName =~ /^-+singlestream/i ) {
            $stream{'/'} = [];
            next;
        }
        next unless -f $fileName;
        my @obj = grep { ref $_ eq 'HASH'; } YAML::LoadFile($fileName);
        if ( @obj == 1 ) {
            push @names, $fileName;
        }
        else {
            push @names, map { "$fileName ($_)"; } 1 .. @obj;
        }
        $fileName =~ s#.*/##s;
        push @{ $stream{$fileName} }, @src .. $#names;
        push @src, @obj;
    }
    if ( my ($single) = grep { @{ $stream{$_} } == 1; } keys %stream ) {
        warn "Single stream (because of $single)";
        %stream = ( 'Single stream.yml' => [ 0 .. $#src ] );
    }

    while ( my ( $stream, $idar ) = each %stream ) {

        my ( %map, %combined );
        foreach my $datasetid (@$idar) {
            next unless ref $src[$datasetid] eq 'HASH';
            while ( my ( $tab, $dat ) = each %{ $src[$datasetid] } ) {
                next unless ref $dat eq 'ARRAY';
                for ( my $col = 0 ; $col < @$dat ; ++$col ) {
                    my $cd = $dat->[$col];
                    next unless ref $cd eq 'HASH';
                    while ( my ( $row, $v ) = each %$cd ) {
                        next unless defined $v;
                        push @{ $map{$tab}[$col]{$row}{$v} }, $datasetid;
                        $combined{$tab}[$col]{$row} = $v;
                    }
                }
            }
        }

        my ( @diff, @addd, %common, %unconflicted );
        while ( my ( $tab, $dat ) = each %map ) {
            for ( my $col = 0 ; $col < @$dat ; ++$col ) {
                while ( my ( $row, $set ) = each %{ $dat->[$col] } ) {
                    my @k = keys %$set;
                    if ( @k == 1 ) {
                        my $v     = $k[0];
                        my $vidar = $set->{$v};
                        $unconflicted{$tab}[$col]{$row} = $v;
                        if ( @$vidar == @$idar ) {
                            $common{$tab}[$col]{$row} = $v;
                        }
                        else {
                            $addd[$_]{$tab}[$col]{$row} = $v foreach @$vidar;
                        }
                    }
                    else {
                        foreach my $v (@k) {
                            $diff[$_]{$tab}[$col]{$row} = $v
                              foreach @{ $set->{$v} };
                        }
                    }
                }
            }
        }

        foreach (@$idar) {
            _ymlDump( $names[$_], ' changed', $diff[$_] );
            _ymlDump( $names[$_], ' added',   $addd[$_] );
        }
        _ymlDump( $stream, ' common',       \%common );
        _ymlDump( $stream, ' accumulated',  \%combined );
        _ymlDump( $stream, ' unconflicted', \%unconflicted );

    }

}

sub ymlMerge {
    my (%results);
    foreach my $fileName ( sort map { decode_utf8 $_} @_ ) {
        $fileName = './' . $fileName unless $fileName =~ m#/#;
        next unless -f $fileName;
        my ( $path, $core, $ext ) = $fileName =~ m#(.*/)([^/]+)(\.ya?ml)$#is
          or next;
        $core =~ s/[0-9+~]//g;
        $results{ $path . $core . $ext } ||= {};
        my $counter = '';
        foreach ( grep { ref $_ eq 'HASH'; } YAML::LoadFile($fileName) ) {
            $results{ $path . $core . $ext } =
              { %{ $results{ $path . $core . $ext } ||= {} }, %$_ };
        }
    }
    while ( my ( $k, $v ) = each %results ) {
        _ymlDump( $k, '', $v );
    }
}

sub ymlSplit {
    foreach my $fileName ( map { decode_utf8 $_} @_ ) {
        $fileName = './' . $fileName unless $fileName =~ m#/#;
        next unless -f $fileName;
        my ( $path, $core, $ext ) = $fileName =~ m#(.*/)([^/]+)(\.ya?ml)$#is
          or next;
        my $counter = '';
        foreach ( grep { ref $_ eq 'HASH'; } YAML::LoadFile($fileName) ) {
            while ( my ( $k, $v ) = each %$_ ) {
                _ymlDump( "$path+$k+$core$counter$ext", '', { $k => $v } );
            }
            --$counter;
        }
    }
}

sub _ymlDump {
    my ( $file, $tag, $data ) = @_;
    my $path = '';
    my $ext  = '';
    $path = $1 if $file =~ s#(.*/)##s;
    $ext  = $1 if $file =~ s/(\.[a-zA-Z][a-zA-Z0-9_+-]+)$//s;
    my $tmp = $path . '~$' . $file . $tag . $$ . '.tmp';
    YAML::DumpFile( $tmp, $data );
    my $final = $path . $file . $tag . $ext;
    rename $tmp, $final;
    warn "$final\n";
}

1;
