package SpreadsheetModel::Export::Controller;

=head Copyright licence and disclaimer

Copyright 2008-2015 Franck Latrémolière, Reckon LLP and others.

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

use constant {
    WE_BASELOC => 0,
    WE_WBOOK   => 1,
    WE_LOC     => 2,
    WE_MODEL   => 3,
};

sub new {
    my ( $class, $dumpLoc, @extras ) = @_;
    $dumpLoc =~ s/\.xlsx?$//i;
    bless [ $dumpLoc, @extras ], $class;
}

sub setModel {
    $_[0][WE_LOC]   = $_[0][WE_BASELOC] . $_[1];
    $_[0][WE_MODEL] = $_[2];
}

sub ExportHtml {
    require SpreadsheetModel::Export::Html;
    mkdir $_[0][WE_LOC];
    chmod 0770, $_[0][WE_LOC];
    SpreadsheetModel::Export::Html::writeHtml( $_[0][WE_MODEL]{logger},
        $_[0][WE_LOC] . '/' );
}

sub ExportText {
    require SpreadsheetModel::Export::Text;
    SpreadsheetModel::Export::Text::writeText( $_[0][WE_MODEL],
        $_[0][WE_LOC] . '-' );
}

sub ExportTablemap {
    require SpreadsheetModel::Export::TableMap;
    SpreadsheetModel::Export::TableMap::updateTableMap( $_[0][WE_MODEL],
        $_[0][WE_LOC] );
}

sub ExportRtf {
    require SpreadsheetModel::Export::Rtf;
    SpreadsheetModel::Export::Rtf::write( $_[0][WE_MODEL], $_[0][WE_LOC] );
}

sub ExportGraphviz {
    require SpreadsheetModel::Export::Graphviz;
    my $dir = $_[0][WE_LOC] . '-graphs';
    mkdir $dir;
    chmod 0770, $dir;
    SpreadsheetModel::Export::Graphviz::writeGraphs(
        $_[0][WE_MODEL]{logger}{objects},
        $_[0][WE_WBOOK], $dir . '/' );
}

sub ExportObjects {
    my @objects = grep { defined $_ } @{ $_[0][WE_MODEL]{logger}{objects} };
    my $objNames = join( "\n",
        $_[0][WE_MODEL]{logger}{realRows}
        ? @{ $_[0][WE_MODEL]{logger}{realRows} }
        : map { "$_->{name}" } @objects );
    (
        $objNames,
        map { UNIVERSAL::can( $_, 'getCore' ) ? $_->getCore : "$_"; } @objects
    );
}

sub ExportYaml {
    require YAML;
    open my $fh, '>', $_[0][WE_LOC] . $$;
    binmode $fh, ':utf8';
    my ( $objNames, @coreObj ) = $_[0]->ExportObjects;
    print {$fh} YAML::Dump(
        {
            '.' => $objNames,
            map { ( ref $_ ? $_->{name} : $_, $_ ); } @coreObj
        }
    );
    close $fh;
    rename $_[0][WE_LOC] . $$, $_[0][WE_LOC] . '.yaml';
}

sub ExportPerl {
    require Data::Dumper;
    my ( $objNames, @coreObj ) = $_[0]->ExportObjects;
    my %counter;
    local $_ = Data::Dumper->new( [ $objNames, @coreObj ] )->Indent(1)->Names(
        [
            'tableNames',
            map {
                my $n =
                  ref $_
                  ? $_->{name}
                  : $_;
                $n =~ s/[^a-z0-9]+/_/gi;
                $n =~ s/^([0-9]+)[0-9]{2}/$1/s;
                "t$n" . ( $counter{$n}++ ? "_$counter{$n}" : '' );
            } @coreObj
        ]
    )->Dump;
    s/\\x\{([0-9a-f]+)\}/chr (hex ($1))/eg;
    open my $fh, '>', $_[0][WE_LOC] . $$;
    binmode $fh, ':utf8';
    print {$fh} $_;
    close $fh;
    rename $_[0][WE_LOC] . $$, $_[0][WE_LOC] . '.pl';
}

1;
