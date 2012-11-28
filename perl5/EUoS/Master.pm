package EUoS;

=head Copyright licence and disclaimer

Copyright 2012 Reckon LLP and others. All rights reserved.

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

    my $setup     = EUoS::Setup->new($model);
    my $customers = EUoS::Customers->new( $model, $setup );
    my $usage     = EUoS::Usage->new( $model, $setup, $customers );
    my $charging  = EUoS::Charging->new( $model, $setup, $usage );

    my %customers;
    while ( my ( $scenario, $exclusions ) =
        each %{ $model->{scenarioExclude} } )
    {
        $customers{$scenario} =
          $customers->scenario( sub { $_[0] !~ /$exclusions/i; }, $scenario );
    }

    my %totalUsage;
    foreach (qw(matchAssets matchRunning matchBoundary)) {
        next unless my $scenario = $model->{$_};
        my $usage = $totalUsage{$scenario} ||=
          $usage->totalUsage( $customers{$scenario} );
        $charging->$_( $usage, $model->{ $_ . 'DoNotApply' } );
    }

    my $tariffs = EUoS::Tariffs->new( $model, $setup, $usage, $charging );

    $tariffs->revenues( $customers{ $model->{revenues} } )
      if $model->{revenues};

    $tariffs->revenues( $customers->detailedVolumes,
        'Notional revenue by customer from the application of UoS tariffs', 1 );

    $_->finish foreach $setup, $usage, $charging, $customers, $tariffs;

    $model;

}

1;
