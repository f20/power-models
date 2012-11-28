package SpreadsheetModel::Logger;

=head Copyright licence and disclaimer

Copyright 2008-2011 Reckon LLP and others. All rights reserved.

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

require SpreadsheetModel::Object;
our @ISA = qw(SpreadsheetModel::Object);

use Spreadsheet::WriteExcel::Utility;

sub log {
    my $self = shift;
    push @{ $self->{objects} }, grep {
             $_->{location}
          || !$_->{sources}
          || $#{ $_->{sources} }
          || $_->{cols} != $_->{sources}[0]{cols}
          || $_->{rows} != $_->{sources}[0]{rows}

          # this is probably a heuristic for some more precise conditions
          # improved in June 2009

    } @_;
}

sub check {
    my ($self) = @_;
    $self->{lines} = [ SpreadsheetModel::Object::splitLines( $self->{lines} ) ]
      if $self->{lines};
    $self->{objects} = [];
    return;
}

sub lastCol {
    3;
}

sub lastRow {
    $_[0]->{realRows} ? $_[0]->{realRows} - 1 : $#{ $_[0]->{objects} };
}

sub xmlElement {
    my ( $e, $c ) = splice @_, 0, 2;
    my %a = %{ ref( $_[0] ) eq 'HASH' ? $_[0] : +{@_} };
    my $z = "<$e";
    while ( my ( $k, $v ) = each %a ) {
        $z .= qq% $k="$v"%;
    }
    defined $c ? "$z>$c</$e>" : "$z />";
}

sub xmlEscape {
    local @_ = @_ if defined wantarray;
    for (@_) {
        if ( defined $_ ) {
            s/&/&amp;/g;
            s/</&lt;/g;
            s/>/&gt;/g;
            s/"/&quot;/g;
        }
    }
    wantarray ? @_ : $_[0];
}

sub xmlFlatten {
    return join '',
      map { ref $_ eq 'ARRAY' ? xmlFlatten(@$_) : xmlEscape($_) } @_
      if ref $_[0] eq 'ARRAY';
    return xmlEscape( $_[1] ) unless $_[0];
    my ( $e, $c ) = splice @_, 0, 2;
    my %a = %{ ref( $_[0] ) eq 'HASH' ? $_[0] : +{@_} };
    my $z = "<$e";
    while ( my ( $k, $v ) = each %a ) {
        xmlEscape $v;
        $z .= qq% $k="$v"%;
    }
    defined $c
      ? "$z>" . ( ref $c ? xmlFlatten($c) : xmlEscape($c) ) . "</$e>"
      : "$z />";
}

sub htmlCode {
    my ( $self, $wbook, $title ) = @_;
    my %htmlWriter;
    my %html;
    foreach my $pot (qw(Inputs Calculations Ancillary)) {
        $html{$pot}       = '';
        $htmlWriter{$pot} = sub {
            $html{$pot} .= xmlFlatten(@$_) foreach @_;
            "$pot.html";
        };
    }
    $_->htmlWrite( \%htmlWriter, $htmlWriter{Calculations} )
      foreach @{ $self->{objects} };
    %html;
}

sub _dotLabel {
    my ( $o, $name, $shape, $fillcolour ) = @_;
    $name =~ s/"/'/g;
    1 while $name =~ s/([^\\]{15,}?) (.{12,})/$1\\n$2/;
    qq#$o [label="$name", shape="$shape", fillcolor="$fillcolour"];#;
}

sub dotCode {
    my ( $self, $wbook, $title ) = @_;
    my $obi = 0;
    my @sheets;
    my @sheetSubgraph;
    my @sheetName;

    foreach ( @{ $self->{objects} } ) {
        if ( $_->{name} =~ /^([0-9]+)([0-9]{2}|[a-z]+)\./ ) {
            push @{ $sheets[$1] }, $_;
            $_->{dotSheet} = $1;
            $sheetSubgraph[$1] = $_->{location}
              if $_->{location} && $wbook->{ $_->{location} };
            $sheetSubgraph[$1] ||= "cluster$1";
            $sheetName[$1] ||= $_->{$wbook}{worksheet}->get_name();
        }
    }

    (
        join "\n",
        'digraph g {',
        'graph [rankdir=LR, size="16,10.5", concentrate=true, nodesep="0.2", '
          . 'ranksep="0.4",fontname="verdanab", fontsize="24", fontcolor="#666666", label="$title"];',
        'node [label="\\N",shape=ellipse, style=filled, fontname=verdana, '
          . 'color="#0066cc", fillcolor=white, fontcolor=black, fontsize="20"];',
        'edge [arrowtail=none, fontname=verdana, color="#ff6633", '
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
                    $_->{dotName}  ||= 'z' . ( 0 + $_ );
                    $ob->{dotName} ||= 'z' . ( 0 + $ob );
                    "$_->{dotName} -> $ob->{dotName};"
                  } map {
                    !$_->{sources}
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
              } @{ $self->{objects} }
        ),
        '}'
      ),

      map {
        my $shno = $_;
        my $graphTitle =
          $sheets[$shno] ? "$_. $sheetName[$_]" : 'Overview';    # $title;
        join "\n", "graph g$_ {",
          'graph [rankdir=TD, size="10.5,8", concentrate=true, nodesep="0.2", '
          . 'ranksep="0.4", fontname="verdanab", fontsize="32", fontcolor="#666666", label="$graphTitle"];',
          'node [label="\\N",shape=ellipse, style=filled, fontname=verdana,'
          . ' color="#0066cc", fillcolor=white, fontcolor=black, fontsize="32"];',
          'edge [arrowtail=none, fontname=verdana, color="#ff6633", '
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
            map {
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
              } @{ $self->{objects} }
          ),
          '}'
      } grep { !$_ || $sheets[$_] } 0 .. $#sheets;
}

sub htmlForGraphs {
    my ( $self, $wbook, $title, $dotName, $dotCommandLine ) = @_;
    $dotCommandLine ||=
      -e '/usr/X11/lib/libpixman-1.0.dylib'
      ? 'DYLD_INSERT_LIBRARIES=/usr/X11/lib/libpixman-1.0.dylib dot'
      : '/Applications/Graphviz.app/Contents/MacOS/dot';
    my $html = '';
    my @dots = $self->dotCode( $wbook, $title );
    foreach ( 0 .. $#dots ) {
        open my $dh, '>', "$dotName$_.dot";
        binmode $dh, ':utf8';
        print $dh $dots[$_];
        undef $dh;
        `$dotCommandLine -Tpdf -o $dotName$_.pdf $dotName$_.dot`;

`cat $dotName$_.dot | sed '1,\$ s/size="[0-9\.]*, *[0-9\.]*"/size="6,4"/' | $dotCommandLine -Tpng -o $dotName$_.png`;
        my ($title) = ( $dots[$_] =~ /label="(.*?)"/ );
        $title ||= 'Graph';
        $html .= <<EOH;
<p><strong>$title</strong></p>
<p><a href="graph$_.pdf"><img src="graph$_.png" alt="graph $_" /><br />The above graph in PDF format</a></p>
<hr />
EOH
    }
    $html;
}

sub _number {
    local $_ = $_[0]->{name};
    return 'Z' unless /^([0-9a-z]+)/;
    $1;
}

sub wsWrite {
    my ( $self, $wb, $ws, $row, $col ) = @_;
    ( $row, $col ) = ( ( $ws->{nextFree} ||= -1 ) + 1, 0 )
      unless defined $row && defined $col;
    $ws->write( $row++, $col, "$self->{name}", $wb->getFormat('caption') );
    my $numFormat0 = $wb->getFormat('0softnz');
    my $numFormat1 = $wb->getFormat('0.000soft');
    my $textFormat = $wb->getFormat('text');
    my $linkFormat = $wb->getFormat('link');
    if ( $self->{lines} ) {
        $ws->write( $row++, $col, "$_", $textFormat )
          foreach @{ $self->{lines} };
    }

    my @h = ( 'Worksheet', 'Data table', 'Type of table' );
    push @h, 'Dimensions', 'Count', 'Average' if $wb->{logAll};

    $ws->write( $row, $col + $_, "$h[$_]", $wb->getFormat('th') ) for 0 .. $#h;
    $row++;

    my @objectList = sort {
        ( $a->{$wb}{worksheet}{sheetNumber} || 666 )
          <=> ( $b->{$wb}{worksheet}{sheetNumber} || 666 )
    } grep { $_->{$wb}{worksheet} } @{ $self->{objects} };

    my $r = 0;
    my %columnsetDone;
    foreach my $obj (@objectList) {

        my $cset;

        unless ( $wb->{logAll} ) {
            $cset = $obj->{location};
            undef $cset unless ref $cset eq 'SpreadsheetModel::Columnset';
            if ($cset) {
                next if exists $columnsetDone{$cset};
                undef $columnsetDone{$cset};
            }
        }

        my ( $wo, $ro, $co ) = @{ $obj->{$wb} }{qw(worksheet row col)};
        my $ty = $cset ? $cset->objectType : $obj->objectType;
        my $ce = xl_rowcol_to_cell( $ro, $co );
        my $wn = $wo->get_name;
        my $na = $cset ? "$cset->{name}" : "$obj->{name}";
        0 and $ws->set_row( $row + $r, undef, undef, 1 ) unless $na;
        $ws->write_url( $row + $r, $col + 1, "internal:'$wn'!$ce", $na,
            $linkFormat );
        $ws->write_string( $row + $r, $col + 2, $ty, $textFormat );
        $ws->write_string( $row + $r, $col,     $wn, $textFormat );

        if ( $wb->{logAll} && $obj->isa('SpreadsheetModel::Dataset') ) {
            my ( $wss, $rows, $cols ) = $obj->wsWrite( $wb, $ws );
            my $wsn = $wss->get_name;
            my $c1 =
              Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $rows,
                $cols, 0, 0 );
            my $c2 = Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                $rows + $obj->lastRow,
                $cols + $obj->lastCol,
                0, 0
            );
            my $range = "'$wsn'!$c1:$c2";
            $ws->write( $row + $r, $col + 3,
                ( 1 + $obj->lastRow ) . ' × ' . ( 1 + $obj->lastCol ),
                $textFormat );
            $ws->write( $row + $r, $col + 4, "=COUNT($range)",   $numFormat0 );
            $ws->write( $row + $r, $col + 5, "=AVERAGE($range)", $numFormat1 );
        }
        ++$r;
    }

    $self->{realRows} = $r;
    $ws->autofilter( $row - 1, $col, $row + $r - 1, $col + 2 );
    0 and $ws->filter_column( $col, 'x <> ""' );

    $ws->{nextFree} = $row + $r
      unless $ws->{nextFree} > $row + $r;
}

1;
