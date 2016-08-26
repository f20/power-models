package SpreadsheetModel::Data::RCode;

=head Copyright licence and disclaimer

Copyright 2015-2016 Franck Latrémolière, Reckon LLP and others.

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

=head Dependencies

This assumes a working version of R with the following packages:
* RSQLite and its many dependencies
* sp
* plotrix
* shape
* treemap

The following commands might help:
	R -e 'install.packages(c("RSQLite", "sp", "plotrix", "shape", "treemap"), repos = "http://mirror.mdx.ac.uk/R/", dependencies = TRUE)'
	R CMD INSTALL <package-file> <package-file> ...

Packages are installed in ~/Library/R/3.2/library/ if using R 3.2 on Mac OS X.

=cut

sub rCode {
    my $self   = shift;
    my $script = {};
    map { $self->$_($script); } @_;
}

sub maps4202cs {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::PriceMaps;
    SpreadsheetModel::Data::RCode::PriceMaps->maps4202cs($script);
}

sub maps4202ts {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::PriceMaps;
    SpreadsheetModel::Data::RCode::PriceMaps->maps4202ts($script);
}

sub margins4203 {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::PriceMaps;
    SpreadsheetModel::Data::RCode::PriceMaps->margins4203($script);
}

*margins = \&margins4203;

sub mapCdcmEdcm {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::PriceMaps;
    SpreadsheetModel::Data::RCode::PriceMaps->mapCdcmEdcm($script);
}

sub treemapWithCategories {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::Treemap;
    SpreadsheetModel::Data::RCode::Treemap->treemapWithCategories;
}

*treemap = \&treemapWithCategories;

sub treemapByCategory {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::Treemap;
    SpreadsheetModel::Data::RCode::Treemap->treemapWithCategories(1);
}

sub treemapWithComponents {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::Treemap;
    SpreadsheetModel::Data::RCode::Treemap->treemapWithComponents;
}

sub treemapByComponent {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::Treemap;
    SpreadsheetModel::Data::RCode::Treemap->treemapWithComponents(1);
}

sub treemap1020ByLevel {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::Treemap;
    SpreadsheetModel::Data::RCode::Treemap->treemap1020;
}

sub treemap1020ByCompany {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::Treemap;
    SpreadsheetModel::Data::RCode::Treemap->treemap1020(1);
}

sub treemap2706ByLevel {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::Treemap;
    SpreadsheetModel::Data::RCode::Treemap->treemap2706;
}

sub treemap2706ByCompany {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::Treemap;
    SpreadsheetModel::Data::RCode::Treemap->treemap2706(1);
}

sub bandConsumption {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::Multi;
    SpreadsheetModel::Data::RCode::Multi->bandConsumption($script);
}

sub coincidenceWaterfall {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::Multi;
    SpreadsheetModel::Data::RCode::Multi->coincidenceWaterfall($script);
}

1;
