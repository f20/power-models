﻿package SpreadsheetModel::Data::RCode;

=head Copyright licence and disclaimer

Copyright 2015-2017 Franck Latrémolière, Reckon LLP and others.

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

This code needs a working version of R callable from the command line as "R",
with the shape package installed.

Treeemaps need the treemap package, and any charts based on data from
a SQLite database need the RSQLite package and its dependencies.

The following command should trigger the installation of shape and treemap:
	R -e 'install.packages(c("shape", "treemap"), repos = "http://mirror.mdx.ac.uk/R/", dependencies = TRUE)'

Alternatively, download the package files manually and run:
	R CMD INSTALL <package-file> <package-file> ...

Installing treemap might require GNU Fortran to be setup first.

=cut

sub rCode {
    my $self   = shift;
    my $script = {};
    map { $self->$_($script); } @_;
}

sub maps3701rate3ts {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::MapsFromDatabase;
    SpreadsheetModel::Data::RCode::MapsFromDatabase->maps3701rate3ts($script);
}

sub maps3701reactivets {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::MapsFromDatabase;
    SpreadsheetModel::Data::RCode::MapsFromDatabase->maps3701reactivets($script);
}

sub maps4202cs {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::MapsFromDatabase;
    SpreadsheetModel::Data::RCode::MapsFromDatabase->maps4202cs($script);
}

sub maps4202ts {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::MapsFromDatabase;
    SpreadsheetModel::Data::RCode::MapsFromDatabase->maps4202ts($script);
}

sub margins4203 {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::MapsFromDatabase;
    SpreadsheetModel::Data::RCode::MapsFromDatabase->margins4203($script);
}

*margins = \&margins4203;

sub mapCdcmEdcm {
    my ( $self, $script ) = @_;
    require SpreadsheetModel::Data::RCode::MapsFromDatabase;
    SpreadsheetModel::Data::RCode::MapsFromDatabase->mapCdcmEdcm($script);
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

eval "require SpreadsheetModel::Data::RCode::MapsOther";
eval "require SpreadsheetModel::Data::RCode::MultiHard";

1;
