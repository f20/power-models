package SpreadsheetModel::ExportHtml;

=head Copyright licence and disclaimer

Copyright 2008-2014 Franck Latrémolière, Reckon LLP and others.

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

sub writeHtml {    # $logger->{objects} is a good $objectList
    my ( $logger, $pathPrefix ) = @_;
    $pathPrefix = '' unless defined $pathPrefix;
    my %htmlWriter;
    my @end;
    foreach my $pot (qw(Inputs Calculations Ancillary)) {
        my $file  = "$pathPrefix$pot.html";
        my $tfile = $pathPrefix . '~$' . $$ . '.' . $pot . '.html';
        open my $fh, '>', $tfile;
        binmode $fh, ':utf8';
        print $fh
          '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8" /><title>'
          . $pot
          . '</title><head><body>';
        my $url = $file;
        $url =~ s/^.*\///s;
        $htmlWriter{$pot} = sub {
            print {$fh} map { _xmlFlatten(@$_); } @_;
            $url;
        };
        push @end, sub {
            print $fh '</body></html>';
            close $fh;
            rename $tfile, $file;
        };
    }
    $_->htmlWrite( \%htmlWriter, $htmlWriter{Calculations} )
      foreach grep { defined $_ } @{ $logger->{objects} };
    $_->() foreach @end;
}

sub _xmlElement {
    my ( $e, $c ) = splice @_, 0, 2;
    my %a = %{ ref( $_[0] ) eq 'HASH' ? $_[0] : +{@_} };
    my $z = "<$e";
    while ( my ( $k, $v ) = each %a ) {
        $z .= qq% $k="$v"%;
    }
    defined $c ? "$z>$c</$e>" : "$z />";
}

sub _xmlEscape {
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

sub _xmlFlatten {
    return join '',
      map { ref $_ eq 'ARRAY' ? _xmlFlatten(@$_) : _xmlEscape($_) } @_
      if ref $_[0] eq 'ARRAY';
    return _xmlEscape( $_[1] ) unless $_[0];
    my ( $e, $c ) = splice @_, 0, 2;
    my %a = %{ ref( $_[0] ) eq 'HASH' ? $_[0] : +{@_} };
    my $z = "<$e";
    while ( my ( $k, $v ) = each %a ) {
        _xmlEscape $v;
        $z .= qq% $k="$v"%;
    }
    defined $c
      ? "$z>" . ( ref $c ? _xmlFlatten($c) : _xmlEscape($c) ) . "</$e>"
      : "$z />";
}

1;
