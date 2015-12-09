package Elec;

=head Copyright licence and disclaimer

Copyright 2012-2014 Franck Latrémolière, Reckon LLP and others.

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
use Elec::Sheets;
use Elec::Setup;
use Elec::Customers;
use Elec::Usage;
use Elec::Charging;
use Elec::Tariffs;
use Elec::Supply;

sub new {
    my $class     = shift;
    my $model     = bless { inputTables => [], @_ }, $class;
    my $setup     = Elec::Setup->new($model);
    my $customers = Elec::Customers->new( $model, $setup );
    my $usage     = Elec::Usage->new( $model, $setup, $customers );
    my $charging  = Elec::Charging->new( $model, $setup, $usage );

    foreach (    # the order affects the column order in input data
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

    my $tariffs = Elec::Tariffs->new( $model, $setup, $usage, $charging );

    if ( my $usetName = $model->{usetRevenues} ) {
        $tariffs->revenues( $customers->totalDemand($usetName) );
    }

    my $supplyTariffs;
    if ( my $usetName = $model->{usetEnergy} ) {
        $supplyTariffs = Elec::Supply->new( $model, $setup, $tariffs,
            $charging->energyCharge->{arguments}{A1} );
        $supplyTariffs->revenues( $customers->totalDemand($usetName) );
        $supplyTariffs->margin( $customers->totalDemand($usetName) )
          if $model->{energyMargin};
        $customers->{compareppu} = Dataset(
            $model->{table1653} ? () : ( number => 1599 ),
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            name     => 'Comparison p/kWh',
            rows     => $customers->userLabelset,
            data     => [ map { 10 } @{ $customers->userLabelset->{list} } ]
        ) if $model->{compareppu};
        if ( $model->{compareppu} || $model->{showppu} ) {
            my $volumes = $customers->individualDemand($usetName);
            $supplyTariffs->revenues(
                $volumes,
                $customers->{compareppu} || $model->{showppu},
                undef, undef,
                $tariffs->revenueCalculation(
                    $customers->individualDemand(
                        $model->{usetRevenues} || $usetName
                    )
                ),
                $supplyTariffs->marginCalculation($volumes)
            );
            $tariffs->revenues(
                $customers->detailedVolumes,
                $customers->{compareppu} || $model->{showppu},
                1, 'Notional revenue by customer',
            ) if $model->{notionalRevenue};
        }
    }

    $charging->detailedAssets(
        $usage->totalUsage( $customers->detailedVolumes )->{source} )
      if $model->{detailedAssets};

    $_->finish($model)
      foreach grep { $_; } $setup, $usage, $charging, $customers, $tariffs,
      $supplyTariffs;

    $model;

}

1;
