package EUoS;

=head Copyright licence and disclaimer

Copyright 2012-2013 Franck Latrémolière, Reckon LLP and others.

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
use EUoS::Sheets;
use EUoS::Setup;
use EUoS::Customers;
use EUoS::Usage;
use EUoS::Charging;
use EUoS::Tariffs;

sub new {

    my $class = shift;
    my $model = bless { inputTables => [], @_ }, $class;

    my $setup = EUoS::Setup->new($model);

    my $customers = EUoS::Customers->new( $model, $setup );

    my $usage = EUoS::Usage->new( $model, $setup, $customers );

    my $charging = EUoS::Charging->new( $model, $setup, $usage );

    foreach ( # the order matters! (affects column order)
        qw(
        usetMatchAssets
        usetBoundaryCosts
        usetRunningCosts
        )
      )
    {
        my $usetName = $model->{$_};
        next unless $usetName;
        my $doNotApply = 0;
        $doNotApply = 1 if $usetName =~ s/ \(information only\)$//i;
        $charging->$_( $usage->totalUsage( $customers->totalDemand($usetName) ),
            $doNotApply );
    }

    my $tariffs = EUoS::Tariffs->new( $model, $setup, $usage, $charging );

    if ( my $usetName = $model->{usetRevenues} ) {
        $tariffs->revenues( $customers->totalDemand($usetName) );
    }

    $tariffs->revenues( $customers->detailedVolumes,
        'Notional revenue by customer', 1 );

    $_->finish foreach $setup, $usage, $charging, $customers, $tariffs;

    $model;

}

1;
