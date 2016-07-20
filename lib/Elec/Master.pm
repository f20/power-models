package Elec;

=head Copyright licence and disclaimer

Copyright 2012-2016 Franck Latrémolière, Reckon LLP and others.

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

sub requiredModulesForRuleset {
    my ( $class, $model ) = @_;
    my @modules = (
        'Elec::Setup',    'Elec::Sheets',
        'Elec::Charging', 'Elec::Customers',
        'Elec::Tariffs',  'Elec::Usage',
    );
    push @modules, 'Elec::Supply' if $model->{usetEnergy};
    push @modules, 'Elec::Summaries'
      if $model->{usetUoS} || $model->{compareppu} || $model->{showppu};
    @modules;
}

sub new {
    my $class     = shift;
    my $model     = bless { inputTables => [], finishList => [], @_ }, $class;
    my $setup     = Elec::Setup->new($model);
    my $customers = Elec::Customers->new( $model, $setup );
    my $usage     = Elec::Usage->new( $model, $setup, $customers );
    my $charging  = Elec::Charging->new( $model, $setup, $usage );

    foreach (  # NB: this order affects the column order in the input data table
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

    my $summary;
    if ( my $usetName = $model->{usetEnergy} ) {
        my $supplyTariffs = Elec::Supply->new( $model, $setup, $tariffs,
            $charging->energyCharge->{arguments}{A1} );
        $supplyTariffs->revenues( $customers->totalDemand($usetName) );
        $supplyTariffs->margin( $customers->totalDemand($usetName) )
          if $model->{energyMargin};
        $summary =
          Elec::Summaries->new( $model, $setup )
          ->setupWithTotals( $customers, $usetName )->summariseTariffs(
            $supplyTariffs,
            [ revenueCalculation => $tariffs ],
            [ marginCalculation  => $supplyTariffs ],
          ) if $model->{compareppu} || $model->{showppu};
    }

    elsif ( $usetName = $model->{usetUoS} ) {
        $summary =
          Elec::Summaries->new( $model, $setup )
          ->setupWithDisabledCustomers( $customers, $usetName )
          ->summariseTariffs($tariffs);
    }

    $summary->addDetailedAssets( $charging, $usage ) if $summary;

    $_->finish($model) foreach @{ $model->{finishList} };

    $model;

}

sub register {
    my ( $model, $object ) = @_;
    push @{ $model->{finishList} }, $object;
    $object;
}

1;
