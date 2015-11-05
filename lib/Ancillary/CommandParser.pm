package Ancillary::CommandParser;

=head Copyright licence and disclaimer

Copyright 2011-2014 Franck Latrémolière and others. All rights reserved.

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

sub factory {
    bless [], $_[0];
}

sub readCommand {
    my $a = shift @{ $_[0] };
    $a ? @$a : ();
}

sub dispatch {
    my $self = shift;
    if ( local $_ = $_[0] ) {
        return push @$self, [ makeFolder => @_[ 1 .. $#_ ] ] if /folder$/i;
        return push @$self, [ makeModels => @_[ 1 .. $#_ ] ] if /models$/i;
        return push @$self, [ fillDatabase => @_[ 1 .. $#_ ] ]
          if /fillDatabase$/i;
        return push @$self, [ useDatabase => @_[ 1 .. $#_ ] ]
          if /useDatabase$/i;
        return push @$self, [ R       => @_[ 1 .. $#_ ] ] if /^R$/i;
        return push @$self, [ sampler => @_[ 1 .. $#_ ] ] if /^sampler$/i;
    }
    return push @$self, [ ymlDiff  => @_ ] if grep { /-+ya?mldiff/si } @_;
    return push @$self, [ ymlIndex => @_ ] if grep { /-*ya?mlindex/si } @_;
    return push @$self, [ ymlMerge => @_ ] if grep { /-+ya?mlmerge/si } @_;
    return push @$self, [ ymlSplit => @_ ] if grep { /-+ya?mlsplit/si } @_;
    return push @$self, [ makeModels => @_ ]
      if grep { /\.(?:ya?ml|json|dta)$/si } @_;
    return push @$self, [ fillDatabase => @_ ]
      if grep { /\.xl\S+$/si || /^-+prune=/si; } @_;
    return push @$self, [ makeModels => @_ ] if grep { /[*?]/; } @_;

    if ( grep { /\.txt$/i } @_ ) {
        $self->interpret($_) foreach grep { -s $_; } @_;
    }
    else {
        push @$self, [ useDatabase => @_ ];
    }
}

sub interpret {
    my ( $self, $file ) = @_;
    my $fh;
    if ( ref $file ) {
        $fh = $file;
    }
    elsif ( -f $file && -r _ ) {
        open $fh, '<', $file;
        unless ($fh) {
            warn "Cannot open $file";
            return;
        }
        local $_ = $file;
        s#.*/##s;
        s/\.te?xt$//i;
        push @$self, [ makeFolder => "_$_" ];
    }
    binmode $fh, ':utf8';
    local $/ = "\n";
    my @buffer = grep { chomp; $_; } <$fh>;
    undef $fh;
    my @current;
    my $indent = qr(\s+);
    while (@buffer) {
        local $_ = shift @buffer;
        if (/^[;#]/s) {
            push @$self, [ comment => $_ ];
            next;
        }
        elsif (s/^($indent)//s) {
            $indent = $1;
            push @current, $_;
            next;
        }
        $self->dispatch(@current) if @current;
        $indent = qr(\s+);
        @current = /\S\s\S/ ? split /\s+/s : $_;
    }
    $self->dispatch(@current) if @current;
}

1;
