package Elec::Summaries;

# Copyright 2012-2022 Franck Latrémolière and others.
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
use Elec::SummariesComparison;

sub new {
    my ( $class, $model, $setup, ) = @_;
    $model->register( bless { model => $model, setup => $setup, }, $class );
}

sub setupByGroup {
    my ( $self, $customers, $usetName ) = @_;
    $self->{usetName} = $usetName;
    $self->{volumes}  = $customers->totalDemand($usetName);
    $self->{comparison} =
      Elec::SummariesComparison->new( $self->{model}, $self->{setup}, undef,
        'revenueTables' );
    $self->setupComparisonPpu(
        Elec::Summaries::CustomersPlaceholder->new( $self->{volumes}[0]{rows} )
    ) if $self->{model}{compareppu} && $self->{model}{compareppu} =~ /group/i;
    $self;
}

sub setupWithActiveCustomers {
    my ( $self, $customers, $usetName ) = @_;
    $self->{usetName} = $usetName;
    $self->{names}    = $customers->names;
    $self->{volumes}  = $customers->individualDemandUsed($usetName);
    $self->{comparison} =
      Elec::SummariesComparison->new( $self->{model}, $self->{setup} );
    $self->setupComparisonPpu($customers);
}

sub setupWithAllCustomers {
    my ( $self, $customers, $usetName ) = @_;
    $self->{names}   = $customers->names;
    $self->{volumes} = $customers->individualDemand;
    $self->{comparison} =
      Elec::SummariesComparison->new( $self->{model}, $self->{setup},
        $customers->totalDemand($usetName)->[0]{matrix} );
    $self->{comparison}
      ->useAlternativeRowset( $customers->userLabelsetRegrouped )
      if UNIVERSAL::can( $customers, 'userLabelsetRegrouped' );
    $self->setupComparisonPpu($customers);
}

sub setupComparisonPpu {
    my ( $self, $customers ) = @_;
    $self->{model}{compareppu} =~ /tariff/
      ? $self->{comparison}->addComparisonTariff( $customers, $self->{setup} )
      : $self->{comparison}->addComparisonPpu($customers)
      if $self->{model}{compareppu};
    $self;
}

sub addDetailedAssets {
    my ( $self, $charging, $usage ) = @_;
    $charging->detailedAssets(
        $usage->detailedUsage( $self->{volumes} ),
        $self->{usetName}
        ? ( showTotals => 1, userName => $self->{usetName} )
        : ()
    ) if $self->{model}{detailedAssets};
    $self;
}

sub addRevenueComparison {
    my ( $self, $tariffs, @methodObjectPairs ) = @_;
    $self->{totalRevenue} = $self->{comparison}->revenueComparison(
        $tariffs,
        $self->{volumes},
        $self->{names},
        map {
            my ( $method, $object ) = @$_;
            $object->$method( $self->{volumes} );
        } @methodObjectPairs
    );
    $self;
}

sub finish {
}

package Elec::Summaries::CustomersPlaceholder;
use SpreadsheetModel::Shortcuts ':all';
sub new          { bless \$_[1], $_[0]; }
sub userLabelset { ${ $_[0] }; }
sub addColumnset { shift; Columnset(@_); }

1;
