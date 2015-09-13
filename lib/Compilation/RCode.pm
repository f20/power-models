package Compilation::RCode;

=head Copyright licence and disclaimer

Copyright 2015 Franck Latrémolière, Reckon LLP and others.

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

sub rCode {
    my $self   = shift;
    my $script = {};
    map { $self->$_($script); } @_;
}

sub maps4202cs {
    my ( $self, $script ) = @_;
    require Compilation::RCode::PriceMaps;
    Compilation::RCode::PriceMaps->maps4202ts($script);
}

sub maps4202ts {
    my ( $self, $script ) = @_;
    require Compilation::RCode::PriceMaps;
    Compilation::RCode::PriceMaps->maps4202ts($script);
}

sub mapCdcmEdcm {
    my ( $self, $script ) = @_;
    require Compilation::RCode::PriceMaps;
    Compilation::RCode::PriceMaps->mapCdcmEdcm($script);
}

sub treemap {
    my ( $self, $script ) = @_;
    require Compilation::RCode::Treemap;
    Compilation::RCode::Treemap->treemapWithCategories;
}

sub treemapWithCategories {
    my ( $self, $script ) = @_;
    require Compilation::RCode::Treemap;
    Compilation::RCode::Treemap->treemapWithCategories;
}

sub treemapByCategory {
    my ( $self, $script ) = @_;
    require Compilation::RCode::Treemap;
    Compilation::RCode::Treemap->treemapWithCategories(1);
}

sub treemapWithComponents {
    my ( $self, $script ) = @_;
    require Compilation::RCode::Treemap;
    Compilation::RCode::Treemap->treemapWithComponents;
}

sub treemapByComponent {
    my ( $self, $script ) = @_;
    require Compilation::RCode::Treemap;
    Compilation::RCode::Treemap->treemapWithComponents(1);
}

sub bandConsumption {
    my ( $self, $script ) = @_;
    require Compilation::RCode::Multi;
    Compilation::RCode::Multi->bandConsumption($script);
}

sub coincidenceWaterfall {
    my ( $self, $script ) = @_;
    require Compilation::RCode::Multi;
    Compilation::RCode::Multi->coincidenceWaterfall($script);
}

1;
