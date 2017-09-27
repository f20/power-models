package PowerModels::CLI::CommandParser;

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
use File::Spec;

sub new {
    bless [], $_[0];
}

sub run {
    my ( $parser, $runner ) = @_;
    foreach (@$parser) {
        my ( $verb, @objects ) = @$_;
        $runner->log( $verb, @objects );
        $runner->$verb(@objects);
    }
}

my %normalisedVerb;
$normalisedVerb{ lc $_ } = $_ foreach qw(
  R
  Rcode
  combineRulesets
  makeFolder
  makeModels
  useDatabase
  useModels
  ymlDiff
  ymlMerge
  ymlSplit
);

sub acceptCommand {
    my $self = shift;
    if ( my $verb = $_[0] ) {
        if ( $verb =~ /(\S+)::\S/ ) {
            eval "require $1";
            if ($@) {
                warn "require $1: $@";
            }
            else {
                return push @$self, [@_];
            }
        }
        if ( $verb = $normalisedVerb{ lc $verb } ) {
            return push @$self, [ $verb => @_[ 1 .. $#_ ] ];
        }
    }
    if ( grep { /\.txt$/i } @_ ) {
        $self->acceptScript($_) foreach grep { -s $_; } @_;
        return;
    }
    return push @$self, [ makeModels => @_ ]
      if grep { /\.(?:ya?ml|json|dta|csv)$/si || /[*?]/; } @_;
    return push @$self, [ useModels => @_ ]
      if grep { /\.xl\S+$/si || /^-+prune=/si; } @_;
    return push @$self, [ useDatabase => @_ ]
      if grep {
             /^tscs/i
          || /^csv/i
          || /^all/i
          || /\//
          || /^[0-9]/
          || /^-+(?:base|dcp|change)=/i;
      } @_;
    warn "Ignored: @_";
}

sub acceptScript {
    my ( $self, $fileOrFilehandle ) = @_;
    my $fh;
    if ( ref $fileOrFilehandle ) {
        $fh = $fileOrFilehandle;
    }
    elsif ( -f $fileOrFilehandle && -r _ ) {
        open $fh, '<', $fileOrFilehandle;
        unless ($fh) {
            warn "Cannot open $fileOrFilehandle";
            return;
        }
        ( undef, undef, local $_ ) = File::Spec->splitpath($fileOrFilehandle);
        s/\.te?xt$//i;
        push @$self, [ makeFolder => $_ ];
    }
    binmode $fh, ':utf8';
    local $/ = "\n";
    my @buffer = grep { chomp; $_; } <$fh>;
    undef $fh;
    my @current;
    while (@buffer) {
        local $_ = shift @buffer;
        if (/^\s*[;#]/s) {
            push @$self, [ comment => $_ ];
        }
        elsif (s/^\s+//s) {
            push @current, $_;
        }
        else {
            $self->acceptCommand(@current) if @current;
            @current = /\S\s\S/ ? split /\s+/s : $_;
        }
    }
    $self->acceptCommand(@current) if @current;
}

1;
