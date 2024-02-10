package Elec;

# Copyright 2012-2023 Franck Latrémolière and others.
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

sub serviceMap {
    my ($model) = @_;
    my @servMap = (
        sheets => 'Elec::Sheets',
        setup  => 'Elec::Setup',
    );
    push @servMap, tariffs  => 'Elec::Tariffs';
    push @servMap, assets   => 'Elec::AssetDetail' if $model->{assetDetail};
    push @servMap, charging => 'Elec::Charging';
    push @servMap, checksum => 'SpreadsheetModel::Checksum'
      if $model->{checksums};
    push @servMap, customers => 'Elec::CustomersTyped' if $model->{ulist};
    push @servMap, customers => 'Elec::Customers' unless $model->{ulist};
    push @servMap,
      interpolator => $model->{interpolator} =~ /ramp/i
      ? 'Elec::DemandRamping'
      : 'Elec::Interpolator'
      if $model->{interpolator};
    push @servMap,
      groupedDemand => 'Elec::CustomersGrouped',
      averagedUsage => 'Elec::UsageAveraged'
      if $model->{usetRouteGrouper};
    push @servMap, summaries => 'Elec::Summaries'
      if $model->{usetUoS} || $model->{compareppu} || $model->{showppu};
    push @servMap, supply    => 'Elec::TariffsSupply' if $model->{usetEnergy};
    push @servMap, timebands => 'Elec::TimebandSets'  if $model->{timebandSets};
    push @servMap, timebands => 'Elec::Timebands'     if $model->{timebands};
    push @servMap, usage     => 'Elec::Usage';
    push @servMap, profits   => 'Elec::Profits' if $model->{showProfits};
    @servMap;
}

sub requiredModulesForRuleset {
    my ( $class, $model ) = @_;
    my %serviceMap = serviceMap($model);
    values %serviceMap;
}

sub register {
    my ( $model, $object ) = @_;
    push @{ $model->{finishList} }, $object;
    $object;
}

sub new {

    my $class = shift;
    my $model = bless { inputTables => [], finishList => [], @_ }, $class;
    ${ $model->{sharingObjectRef} }->registerModel($model)
      if ref $model->{sharingObjectRef}
      && UNIVERSAL::can( ${ $model->{sharingObjectRef} }, 'registerModel' );
    $model->{dataset}{datasetCallback}->($model)
      if $model->{dataset} && $model->{dataset}{datasetCallback};

    # Set-up

    my %serviceMap = $model->serviceMap;
    my $setup      = $serviceMap{setup}->new($model);
    $setup->registerTimebands( $serviceMap{timebands}->new( $model, $setup ) )
      if $serviceMap{timebands};
    $model->{interpolator} =
      $serviceMap{interpolator}->new( $model, $setup, $model->{interpolator} )
      if $model->{interpolator};
    my $customers = $serviceMap{customers}->new( $model, $setup );
    my $usage     = $serviceMap{usage}->new( $model, $setup, $customers );
    my $assets    = $serviceMap{assets}->new( $model, $setup )
      if $serviceMap{assets};

    # Matching and grouping

    my $charging =
      $serviceMap{charging}->new( $model, $setup, $usage, $assets );

    if ( my $usetName = $model->{'usetMatchUsage'} ) {
        $usage = $usage->matchTotalUsage( $customers->totalDemand($usetName) );
    }

    # Legacy fudge
    $model->{'usetNonAssetCosts'} ||= $model->{'usetBoundaryCosts'};

    # This will append a column to input data table for each unique uset.
    my $applyChargingMatchingRule = sub {
        my ($rulename) = @_;
        return unless my $usetName = $model->{$rulename};
        my $applicationOptions = '';
        $applicationOptions = $1 if $usetName =~ s/(\s*\(.*\))$//i;
        $charging->$rulename(
            $usage->totalUsage( $customers->totalDemand($usetName) ),
            $applicationOptions );
    };

    $applyChargingMatchingRule->('usetMatchAssetDetail');

    if ( my $usetName = $model->{'usetRouteGrouper'} ) {
        $customers =
          $serviceMap{groupedDemand}->new( $model, $setup, $customers );
        $usage = $serviceMap{averagedUsage}->new(
            $model, $setup, $customers, ' (averaged)',
            usage => $usage,
            uset  => $usetName
        );
        $charging =
          $serviceMap{charging}->new( $model, $setup, $usage, $assets );
    }

    $applyChargingMatchingRule->('usetMatchAssets');

    $applyChargingMatchingRule->('usetNonAssetCosts');

    $applyChargingMatchingRule->('usetRunningCosts');

    # Use of system tariff calculation

    my $tariffs =
      $serviceMap{tariffs}->new( $model, $setup, $usage, $charging );

    $tariffs->showAverageUnitRateTable($customers)
      if $serviceMap{timebands} && $model->{showAverageUnitRateTable};

    # Revenues, supply tariffs, reporting and statistics
    # NB: the order of feeding arguments to $customers->totalDemand will
    # determine the column order for proportions in the input data table.

    my $distributionRevenue;
    if ( my $usetName = $model->{usetRevenues} ) {
        if ( $model->{showppu} ) {
            $serviceMap{summaries}->new( $model, $setup )
              ->setupByGroup( $customers, $usetName )
              ->addRevenueComparison($tariffs);
            $serviceMap{summaries}->new( $model, $setup )
              ->setupWithActiveCustomers( $customers, $usetName )
              ->addRevenueComparison($tariffs)
              ->addDetailedAssets( $charging, $usage );
        }
        else {
            $tariffs->revenues( $customers->totalDemand($usetName) );
        }
        $distributionRevenue =
          $tariffs->revenueCalculation( $customers->totalDemand($usetName) );
    }

    if ( my $usetName = $model->{usetEnergy} ) {
        my $supplyTariffs =
          $serviceMap{supply}
          ->new( $model, $setup, $tariffs, $charging->energyCost );
        $supplyTariffs->revenues( $customers->totalDemand($usetName) );
        my $energyMarginTotal =
          $supplyTariffs->margin( $customers->totalDemand($usetName),
            $serviceMap{profits} && $model->{usetRunningCosts} )
          if $model->{energyMargin};
        $serviceMap{profits}->new(
            $model,
            $energyMarginTotal,
            $distributionRevenue,
            $tariffs,
            $customers->totalDemand( $model->{usetRunningCosts} ),
            $charging->costItems,
        ) if $serviceMap{profits} && $model->{usetRunningCosts};
        $serviceMap{summaries}->new( $model, $setup )
          ->setupWithActiveCustomers( $customers, $usetName )
          ->addRevenueComparison(
            $supplyTariffs,
            [ revenueCalculation => $tariffs ],
            [ marginCalculation  => $supplyTariffs ],
        )->addDetailedAssets( $charging, $usage )
          if $model->{compareppu} || $model->{showppu};
    }

    if ( my $usetName = $model->{usetUoS} ) {
        $serviceMap{summaries}->new( $model, $setup )
          ->setupByGroup( $customers, $usetName )
          ->addRevenueComparison($tariffs)
          ->addDetailedAssets( $charging, $usage );
        if ( $model->{detailsByCustomer} ) {
            my $method =
              $model->{detailsByCustomer} =~ /all/i
              ? 'setupWithAllCustomers'
              : 'setupWithActiveCustomers';
            $serviceMap{summaries}->new( $model, $setup )
              ->$method( $customers, $usetName )
              ->addRevenueComparison($tariffs)
              ->addDetailedAssets( $charging, $usage );
        }
    }

    # Finish

    $_->finish($model) foreach @{ $model->{finishList} };

    $model;

}

1;
