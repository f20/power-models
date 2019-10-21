﻿package Elec;

# Copyright 2012-2019 Franck Latrémolière, Reckon LLP and others.
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
    my @modules = (
        sheets => 'Elec::Sheets',
        setup  => 'Elec::Setup',
    );
    push @modules, tariffs  => 'Elec::Tariffs';
    push @modules, assets   => 'Elec::AssetDetail' if $model->{assetDetail};
    push @modules, charging => 'Elec::Charging';
    push @modules, checksum => 'SpreadsheetModel::Checksum'
      if $model->{checksums};
    push @modules, customers => 'Elec::CustomersTyped' if $model->{ulist};
    push @modules, customers => 'Elec::Customers' unless $model->{ulist};
    push @modules, interpolator => 'Elec::Interpolator'
      if $model->{interpolator};
    push @modules, summaries => 'Elec::Summaries'
      if $model->{usetUoS} || $model->{compareppu} || $model->{showppu};
    push @modules, supply    => 'Elec::Supply'       if $model->{usetEnergy};
    push @modules, timebands => 'Elec::TimebandSets' if $model->{timebandSets};
    push @modules, timebands => 'Elec::Timebands'    if $model->{timebands};
    push @modules, usage     => 'Elec::Usage';
    @modules;
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

    my $class      = shift;
    my $model      = bless { inputTables => [], finishList => [], @_ }, $class;
    my %serviceMap = $model->serviceMap;
    my $setup      = $serviceMap{setup}->new($model);
    $setup->registerTimebands( $serviceMap{timebands}->new( $model, $setup ) )
      if $serviceMap{timebands};
    $model->{interpolator} = $serviceMap{interpolator}->new( $model, $setup )
      if $model->{interpolator};
    my $customers = $serviceMap{customers}->new( $model, $setup );
    my $usage = $serviceMap{usage}->new( $model, $setup, $customers );
    my $assets = $serviceMap{assets}->new( $model, $setup )
      if $serviceMap{assets};
    my $charging =
      $serviceMap{charging}->new( $model, $setup, $usage, $assets );

    # Matching activities
    # Note that the order of feeding arguments to $customers->totalDemand can
    # affect the column order for proportions included in the input data table.

    if ( my $usetName = $model->{usetMatchUsage} ) {
        $usage = $usage->matchTotalUsage( $customers->totalDemand($usetName) );
    }

    $model->{usetNonAssetCosts} ||= $model->{usetBoundaryCosts};   # Legacy only
    foreach (qw(usetMatchAssets usetNonAssetCosts usetRunningCosts)) {
        next unless my $usetName = $model->{$_};
        my $doNotApply = $usetName =~ s/ \(information only\)$//i;
        $charging->$_( $usage->totalUsage( $customers->totalDemand($usetName) ),
            $doNotApply );
    }

    # Use of system tariff calculation

    my $tariffs =
      $serviceMap{tariffs}->new( $model, $setup, $usage, $charging );

    $tariffs->showAverageUnitRateTable($customers)
      if $serviceMap{timebands} && $model->{showAverageUnitRateTable};

    # Revenues, supply tariffs, reporting and statistics

    if ( my $usetName = $model->{usetRevenues} ) {
        if ( $model->{showppu} ) {
            $serviceMap{summaries}->new( $model, $setup )
              ->setupByGroup( $customers, $usetName )
              ->addRevenueComparison($tariffs);
        }
        else {
            $tariffs->revenues( $customers->totalDemand($usetName) );
        }
    }

    if ( my $usetName = $model->{usetEnergy} ) {
        my $supplyTariffs =
          $serviceMap{supply}->new( $model, $setup, $tariffs,
            $charging->energyCharge->{arguments}{A1} );
        $supplyTariffs->revenues( $customers->totalDemand($usetName) );
        $supplyTariffs->margin( $customers->totalDemand($usetName) )
          if $model->{energyMargin};
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
