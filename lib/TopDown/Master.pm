package TopDown;

=head Copyright licence and disclaimer

Copyright 2014 Franck Latrémolière, Reckon LLP and others.

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
use SpreadsheetModel::Shortcuts ':all';

use TopDown::Customers;
use TopDown::Pots;
use TopDown::Routeing;
use TopDown::Usage;
use TopDown::Rates;
use TopDown::Tariffs;
use TopDown::Sheets;

sub new {
    my $class = shift;
    my $model = bless { inputTables => [], @_ }, $class;
    my @needFinish;

    push @needFinish, my $customers = TopDown::Customers->new($model);

    push @needFinish, my $pots = TopDown::Pots->new($model);

    push @needFinish,
      my $routeing = TopDown::Routeing->new( $model, $customers, $pots );

    push @needFinish,
      my $usage = TopDown::Usage->new( $model, $customers, $pots, $routeing );

    push @needFinish, my $rates = TopDown::Rates->new( $model, $usage, $pots );

    push @needFinish,
      my $tariffs =
      TopDown::Tariffs->new( $model, $customers, $rates, $routeing );

    $tariffs->revenueSummary($customers);

    $_->finish($model) foreach @needFinish;
    $model;
}

1;
