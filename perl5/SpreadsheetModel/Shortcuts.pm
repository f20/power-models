package SpreadsheetModel::Shortcuts;

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

use SpreadsheetModel::Label;

use SpreadsheetModel::Labelset;

use SpreadsheetModel::Notes;

use SpreadsheetModel::Dataset;
use SpreadsheetModel::Columnset;

use SpreadsheetModel::Arithmetic;
use SpreadsheetModel::GroupBy;
use SpreadsheetModel::Stack;
use SpreadsheetModel::SumProduct;

use SpreadsheetModel::Miscellaneous;

require Exporter;
our @ISA = qw(Exporter);
our @EXPORT_OK =
  qw(Arithmetic Columnset Constant Dataset GroupBy Label Labelset Notes Stack SumProduct View);
our %EXPORT_TAGS = ( 'all' => \@EXPORT_OK );

# use goto in these subroutines in order to preserve caller

sub Arithmetic {
    unshift @_, 'SpreadsheetModel::Arithmetic';
    goto \&SpreadsheetModel::Object::new;
}

sub Columnset {
    unshift @_, 'SpreadsheetModel::Columnset';
    goto \&SpreadsheetModel::Object::new;
}

sub Constant {
    unshift @_, 'SpreadsheetModel::Constant';
    goto &SpreadsheetModel::Object::new;
}

sub Dataset {
    unshift @_, 'SpreadsheetModel::Dataset';
    goto &SpreadsheetModel::Object::new;
}

sub GroupBy {
    unshift @_, 'SpreadsheetModel::GroupBy';
    goto &SpreadsheetModel::Object::new;
}

sub Label {
    new SpreadsheetModel::Label(@_);
}

sub Labelset {
    unshift @_, 'SpreadsheetModel::Labelset';
    goto &SpreadsheetModel::Object::new;
}

sub Notes {
    unshift @_, 'SpreadsheetModel::Notes';
    goto &SpreadsheetModel::Object::new;
}

sub Stack {
    unshift @_, 'SpreadsheetModel::Stack';
    goto &SpreadsheetModel::Object::new;
}

sub SumProduct {
    unshift @_, 'SpreadsheetModel::SumProduct';
    goto &SpreadsheetModel::Object::new;
}

sub View {
    unshift @_, 'SpreadsheetModel::View';
    goto &SpreadsheetModel::Object::new;
}

1;
