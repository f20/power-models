package SpreadsheetModel::ExportText;

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

sub writeText {
    my ( $options, $pathPrefix ) = @_;
    if ( my $yaml = $options->{yaml} ) {
        my $file  = "${pathPrefix}Rules.txt";
        my $tfile = $pathPrefix . $$ . '.txt';
        open my $fh, '>', $tfile;
        binmode $fh, ':utf8';
        print $fh $yaml;
        rename $tfile, $file;
    }
    my $logger = $options->{logger};
    $pathPrefix = '' unless defined $pathPrefix;
    my %writer;
    my @end;
    foreach my $pot (qw(Datasets Labelsets Tables)) {
        my $file  = "$pathPrefix$pot.txt";
        my $tfile = $pathPrefix  . $$ . '.' . $pot . 'txt';
        open my $fh, '>', $tfile;
        binmode $fh, ':utf8';
        my $url = $file;
        $url =~ s/^.*\///s;
        $writer{$pot} = sub {
            print {$fh} map { _flatten(@$_); } @_;
            $url;
        };
        push @end, sub {
            close $fh;
            rename $tfile, $file;
        };
    }
    $writer{Inputs} = $writer{Calculations} = $writer{Datasets};
    $writer{Ancillary} = $writer{Labelsets};
    my @objects = grep { defined $_ } @{ $logger->{objects} };
    $writer{Tables}->(
        [
            undef => join "\n",
            $logger->{realRows}
            ? @{ $logger->{realRows} }
            : map { "$_->{name}" } @objects
        ]
    );
    $_->htmlWrite( \%writer, $writer{Calculations} ) foreach @objects;
    $_->() foreach @end;
}

sub _flatten {
    return join '', map { ref $_ eq 'ARRAY' ? _flatten(@$_) : $_ } @_
      if ref $_[0] eq 'ARRAY';
    return $_[1] unless $_[0];
    my ( $e, $c ) = splice @_, 0, 2;
    ( $e eq 'fieldset' ? "-\n" : '' )
      . ( !defined $c ? '' : ref $c ? _flatten($c) : $c )
      . ( $e eq 'legend' || $e eq 'div' || $e eq 'p' ? "\n" : '' );
}

1;
