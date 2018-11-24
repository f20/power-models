package SpreadsheetModel::Export::Rtf;

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
use Encode qw(encode);

sub write {
    my ( $options, $logger, $name ) = @_;
    my ( %blob, %writer );
    foreach my $pot (qw(Calculations Inputs Ancillary)) {
        $blob{$pot}   = '';
        $writer{$pot} = sub {
            $blob{$pot} .= _flatten(@$_) foreach @_;
        };
    }
    my @objects = grep { defined $_ } @{ $logger->{objects} };
    $_->htmlWrite( \%writer, $writer{Calculations} ) foreach @objects;
    my $tfile = $name . $$ . '.doc';
    open my $fh, '>', $tfile;
    binmode $fh;
    print $fh <<'EORTFH';
{\rtf1\ansi\ansicpg1252\cocoartf1265\cocoasubrtf200
{\fonttbl\f0\fswiss\fcharset0 Helvetica;\f1\fmodern\fcharset0 Courier;}
{\colortbl;\red255\green255\blue255;\red51\green51\blue51;}
\paperw11900\paperh16840\margl1440\margr1440\margt1800\vieww10800\viewh8400\viewkind0
EORTFH
    local $_ = $name;
    s#.*/##s;
    print $fh '\pard\ri0\sa240\f0\b\fs28\cf2\outlinelevel0\keepn ' . "$_\\\n";
    print $fh _rtfCode( '' => $options->{yaml} );
    print $fh _rtfCode(
        Tables => join "\n",
        $logger->{realRows}
        ? @{ $logger->{realRows} }
        : map { "$_->{name}\n" } @objects
    );
    print $fh _rtfCode( Inputs       => $blob{Inputs} );
    print $fh _rtfCode( Calculations => $blob{Calculations} );
    print $fh _rtfCode( Ancillary    => $blob{Ancillary} );
    print $fh '}';
    close $fh;
    rename $tfile, "$name.doc";
}

sub _rtfCode {
    ( my $h2, local $_ ) = @_;
    return '' unless $_;
    s/[\r\n]*$/\n\n/s;
    s/\r?\n/\\\n/gs;
    (
        $h2
        ? '\pard\ri0\sa240\f0\b\fs22\cf2\outlinelevel1\keepn\pagebb '
          . encode( 'iso-8859-1', $h2 ) . "\\\n"
        : ''
      )
      . '\pard\ri0\f1\b0\fs16\cf0 '
      . encode( 'iso-8859-1', $_ );
}

sub _pmarks {
    ( local $_ ) = @_;
    s/\r?\n/¶/gs;
    $_;
}

sub _flatten {
    return join '', map { ref $_ eq 'ARRAY' ? _flatten(@$_) : $_ } @_
      if ref $_[0] eq 'ARRAY';
    return _pmarks( $_[1] ) unless $_[0];
    my ( $e, $c ) = splice @_, 0, 2;
    ( !defined $c ? '' : ref $c ? _flatten($c) : _pmarks($c) )
      . ( $e eq 'legend' || $e eq 'div' || $e eq 'p' ? "\n" : '' )
      . ( $e eq 'fieldset' ? "\n" : '' );
}

1;
