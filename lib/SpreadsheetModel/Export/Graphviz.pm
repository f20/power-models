package SpreadsheetModel::Export::Graphviz;

# Copyright 2008-2017 Franck Latrémolière, Reckon LLP and others.
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

my $dotCL = 'dot';
0 and $dotCL = '/usr/local/bin/dot';
0 and $dotCL = '/Applications/Graphviz.app/Contents/MacOS/dot';
0 and $dotCL = 'DYLD_INSERT_LIBRARIES=/usr/X11/lib/libpixman-1.0.dylib dot';

sub writeGraphs {    # $logger->{objects} is a good $objectList
    my ( $objectList, $wbook, $pathPrefix ) = @_;
    $pathPrefix = '' unless defined $pathPrefix;
    my $obi = 0;
    my @sheets;
    my @sheetSubgraph;
    my @sheetName;
    foreach (@$objectList) {
        if ( $_->{name} =~ /^([0-9]+)([0-9]{2}|[a-z]+)\./ ) {
            push @{ $sheets[$1] }, $_;
            $_->{dotSheet} = $1;
            $sheetSubgraph[$1] = $_->{location}
              if $_->{location} && $wbook->{ $_->{location} };
            $sheetSubgraph[$1] ||= "cluster$1";
            $sheetName[$1] ||= $_->{$wbook}{worksheet}->get_name();
        }
    }

    my @dots;

    push @dots,
      join "\n",
      'digraph g {',
      'graph [rankdir=LR, size="8,8", concentrate=true, nodesep="0.2", '
      . 'ranksep="0.4",fontname="Verdanab", fontsize="24", fontcolor="#666666"];',
      'node [label="\\N",shape=ellipse, style=filled, fontname=Verdana, '
      . 'color="#0066cc", fillcolor=white, fontcolor=black, fontsize="20"];',
      'edge [arrowtail=none, fontname=Verdana, color="#ff6633", '
      . 'fontcolor="#99b3cc", fontsize=18, style="setlinewidth(3)", arrowsize="2"];',
      (
        map {
            $sheets[$_]
              ? (
                "subgraph $sheetSubgraph[$_] {",
                qq%label="$sheetName[$_]";%,
                (
                    map {
                        _dotLabel(
                            $_->{dotName} =
                              't' . ++$obi,
                            $_->{name},
                            ref $_ eq 'SpreadsheetModel::Dataset'
                            ? ( rect => '#ccffff' )
                            : ref $_ eq 'SpreadsheetModel::Constant'
                            ? ( rect => '#e9e9e9' )
                            : ( Mrecord => '#ffffcc' )
                        );
                    } @{ $sheets[$_] }
                ),
                '}'
              )
              : ()
        } 0 .. $#sheets
      ),
      (
        map {
            my $ob = $_;
            $ob->{sourceLines}
              ? map {
                $_->{usedIn}[ $ob->{dotSheet} ] = 1;
                $_->{dotName}
                  && $ob->{dotName} ? "$_->{dotName} -> $ob->{dotName};" : ();
              } map {
                !$_->{sources}
                  || $#{ $_->{sources} }
                  || $_->{cols} != $_->{sources}[0]{cols}
                  || $_->{rows} != $_->{sources}[0]{rows}
                  ? $_
                  : $_->{sources}[0]
              } map {
                UNIVERSAL::isa( $_, 'SpreadsheetModel::View' )
                  ? $_->{sources}[0]
                  : $_;
              } @{ $ob->{sourceLines} }
              : ();
          } grep {
            $_->isa('SpreadsheetModel::Dataset')
              || $_->isa('SpreadsheetModel::Columnset')
          } @$objectList
      ),
      '}';

    foreach my $shno ( grep { !$_ || $sheets[$_] } 0 .. $#sheets ) {
        my %dedup;
        push @dots,
          join "\n", "digraph g$shno {",
          'graph [rankdir="TD", size="8,8", concentrate="'
          . ( $sheets[$shno] ? 'false' : 'false' )
          . '", nodesep="0.2", '
          . 'ranksep="0.4", fontname="Verdanab", fontsize="32", fontcolor="#666666", label="'
          . ( $sheets[$shno] ? "$sheetName[$shno]" : '' ) . '"];',
          'node [label="\\N",shape=ellipse, style=filled, fontname=Verdana,'
          . ' color="#0066cc", fillcolor=white, fontcolor=black, fontsize="32"];',
          'edge [arrowtail=none, fontname=Verdana, color="#ff6633", '
          . 'fontcolor="#99b3cc", fontsize=18, style="setlinewidth(3)", arrowsize="2"];',
          (
            map {
                   !$sheets[$_] ? ()
                  : $_ != $shno && $sheetSubgraph[$_] =~ /^cluster/
                  ? _dotLabel( "s$_", $sheetName[$_], ellipse => '#ffffff' )
                  : (
                    "subgraph $sheetSubgraph[$_] {",
                    'label="";',    # qq%label="$sheetName[$_]";%,
                    (
                        map {
                            _dotLabel( $_->{dotName}, $_->{name},
                                ref $_ eq 'SpreadsheetModel::Dataset'
                                ? ( rect => '#ccccff' )
                                : ref $_ eq 'SpreadsheetModel::Constant'
                                ? ( rect => '#cccccc' )
                                : ( Mrecord => '#ffffcc' ) );
                          } $_ == $shno
                        ? @{ $sheets[$_] }
                        : grep { $_->{usedIn}[$shno] } @{ $sheets[$_] }
                    ),
                    '}'
                  )
            } 0 .. $#sheets
          ),
          (
            grep { my $ok = !exists $dedup{$_}; undef $dedup{$_}; $ok; } map {
                my $obdn =
                  (      $_->{dotSheet} == $shno
                      || $sheetSubgraph[ $_->{dotSheet} ] !~ /^cluster/ )
                  ? $_->{dotName}
                  : "s$_->{dotSheet}";
                my $ob = $_;
                $ob->{sourceLines}
                  ? map {
                         !defined $_->{dotSheet}
                      || !$_->{usedIn}[$shno] && $obdn !~ /^s/ ? ()
                      : $_->{dotSheet} == $shno ? "$_->{dotName} -> $obdn;"
                      : $sheetSubgraph[ $_->{dotSheet} ] !~ /^cluster/ ? (
                        $_->{usedIn}[$shno] && $ob->{dotSheet} == $shno
                        ? "$_->{dotName} -> $obdn;"
                        : ()
                      )
                      : "s$_->{dotSheet}" eq $obdn ? ()
                      : "s$_->{dotSheet} -> $obdn;"
                  } map {
                    !$_->{sources}
                      || UNIVERSAL::isa( $_->{location},
                        'SpreadsheetModel::CalcBlock' )
                      || $#{ $_->{sources} }
                      || $_->{cols} != $_->{sources}[0]{cols}
                      || $_->{rows} != $_->{sources}[0]{rows}
                      ? $_
                      : $_->{sources}[0]
                  } map {
                    ref $_ eq 'SpreadsheetModel::View' ? $_->{sources}[0] : $_;
                  } @{ $ob->{sourceLines} }
                  : ();
              } grep {
                $_->isa('SpreadsheetModel::Dataset')
                  || $_->isa('SpreadsheetModel::Columnset')
              } @$objectList
          ),
          '}';
    }

    open my $hh, '>', $pathPrefix . 'index.html';
    binmode $hh, ':utf8';
    my $prefix = $pathPrefix;
    $prefix =~ s#^.*/##s;
    foreach ( 0 .. $#dots ) {
        my $name;
        $name = $1 if $dots[$_] =~ /graph\s*\[.*label="(.+?)"/m;
        undef $name if $name && $name =~ /[^ ,.0-9=@-~-]/s;
        $name ||= $_;
        $name ||= 'Everything';
        while ( -e "$pathPrefix$name.txt" ) { $name .= '_'; }
        open my $dh, '>', "$pathPrefix$name.txt";
        binmode $dh, ':utf8';
        print $dh $dots[$_];
        close $dh;
        `$dotCL -Tpdf -o '$pathPrefix$name.pdf' '$pathPrefix$name.txt'`;
        $dots[$_] =~ s/size="[0-9\.]*, *[0-9\.]*"/size="24,24"/;
        open $dh, "| $dotCL -Tpng -o '$pathPrefix$name.png'";
        binmode $dh, ':utf8';
        print $dh $dots[$_];
        close $dh;
        my $att = '';
        {
            open $dh, '<', "$pathPrefix$name.png";
            read $dh, local $_, 32;
            close $dh;
            my ( $width, $height ) = unpack( "NN", $1 ) if /IHDR(........)/;
            if ( $width && $height ) {
                my $max = $width > $height ? $width : $height;
                $max    = 800 if $max < 800;
                $width  = int( $width * 800 / $max );
                $height = int( $height * 800 / $max );
                $att    = qq% style="width:${width}px;heigth:${height}px"%;
            }
        }
        print $hh "<p><strong>$name</strong></p>",
          qq%<p><a href="$prefix$name.pdf">%,
          qq%<img src="$prefix$name.png" alt="$name"$att />%,
          '<br />PDF version of the above graph</a></p><hr />';
    }
    close $hh;

}

sub _dotLabel {
    my ( $o, $name, $shape, $fillcolour ) = @_;
    $name =~ s/"/'/g;
    1 while $name =~ s/([^\\]{15,}?) (.{12,})/$1\\n$2/;
    qq#$o [label="$name", shape="$shape", fillcolor="$fillcolour"];#;
}

1;
