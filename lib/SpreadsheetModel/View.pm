package SpreadsheetModel::View;

# Copyright 2008-2023 Franck Latrémolière and others.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice,
# this list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
# this list of conditions and the following disclaimer in the documentation
# and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
# EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
# THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Stack;
our @ISA = qw(SpreadsheetModel::Stack);

sub check {
    my ($self) = @_;
    return 'No sources in view' unless @{ $self->{sources} };
    my $n = $self->{sources}[0]{name};
    $self->{name} = new SpreadsheetModel::Label($n);
    if ( $#{ $self->{sources} } == 0 && $self->{sources}[0]{columns} ) {
        $self->{rows} = $self->{sources}[0]{rows} unless defined $self->{rows};
        return 'Mismatched rows in columnset to dataset conversion'
          unless $self->{rows} == $self->{sources}[0]{rows};
        my $ncol = 0;
        $ncol += 1 + $_->lastCol foreach @{ $self->{sources}[0]{columns} };
        return
          'Mismatched number of columns in columnset to dataset conversion'
          unless @{ $self->{cols}{list} } == $ncol;
        return;
    }
    return $self->SUPER::check;
}

sub wsWrite {
    my ( $self, @args ) = @_;
    die 'Cannot write this type of View'
      unless $#{ $self->{sources} } == 0 && $self->{sources}[0]{columns};
    $self->{sources}[0]{columns}[0]->wsWrite(@args);
}

sub objectType {
    'Select from table';
}

sub wsUrl {
    my $self = shift;
    $self->SUPER::wsUrl(@_) || $self->{sources}[0]->wsUrl(@_);
}

sub addForwardLink {
    my $self = shift;
    $self->SUPER::addForwardLink(@_);
    (
        $_->{location}
          && ref $_->{location} eq 'SpreadsheetModel::Columnset'
        ? $_->{location}
        : $_
    )->addForwardLink(@_)
      foreach @{ $self->{sources} };
}

1;
