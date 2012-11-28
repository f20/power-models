package PCFM;

=head Copyright licence and disclaimer

Copyright 2012 Reckon LLP and others. All rights reserved.

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
use SpreadsheetModel::Shortcuts ':all';

# Compatibility layer
sub new { init(@_)->loadModules->makeTables->useDataset; }

sub useDataset {
    my ( $model, $dataset ) = @_;
    if ( $dataset ||= $model->{dataset} ) {
        $_->{dataset} = $dataset foreach @{ $model->{inputTables} };
    }
    $model;
}

sub init {
    my $class = shift;
    bless {@_}, $class;
}

sub loadModules {
    my ($model) = @_;
    my $require = $model->{require};
    foreach my $module ( 'HIDAM::Sheets', 'HIDAM::Inputs',
        ref $require ? @$require : $require ? $require : () )
    {
        eval "require $module";
        die "Cannot load $module: $@" if $@;
    }
    $model;
}

sub makeTables {
    my ($model) = @_;
    $model->{inputTables} = [];
    $model;
}

1;
