package SpreadsheetModel::Constant;

# Copyright 2008-2018 Franck Latrémolière, Reckon LLP and others.
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
require SpreadsheetModel::Dataset;
use SpreadsheetModel::Label;
use Spreadsheet::WriteExcel::Utility;
our @ISA = qw(SpreadsheetModel::Dataset);

sub populateCore {
    my ($self) = @_;
    $self->{core}{$_} = $self->{$_}
      foreach grep { exists $self->{$_}; } qw(data);
}

sub dataset {
    return;
}

sub check {
    $_[0]{defaultFormat}        ||= '0.000con';
    $_[0]{defaultMissingFormat} ||= 'unavailable';
    return "No data in constant $_[0]{name}" unless 'ARRAY' eq ref $_[0]{data};
    $_[0]{arithmetic} = '[ ' . join(
        ( $_[0]{byrow} ? ",\n" : ', ' ),
        map {
                !defined $_       ? 'undef'
              : $_ eq ''          ? q['']
              : ref $_ eq 'ARRAY' ? (
                '[ '
                  . join( ', ',
                    map { !defined $_ ? 'undef' : $_ eq '' ? q[''] : "$_" }
                      @$_ )
                  . ' ]'
              )
              : "$_"
        } @{ $_[0]{data} }
    ) . ' ]';
    $_[0]->SUPER::check;
}

sub objectType {
    $_[0]{specialType} || 'Fixed data';
}

1;
