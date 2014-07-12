package ModelM;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others.
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
use ModelM::Options;
use ModelM::Sheets;
use ModelM::Inputs;
use ModelM::Meav;
use ModelM::NetCapex;
use ModelM::Expenditure;
use ModelM::Allocation;
use ModelM::Discounts;

sub requiredModulesForRuleset {
    my ( $class, $ruleset ) = @_;
    $ruleset->{dcp095}   ? qw(ModelM::Dcp095) : (),
      $ruleset->{dcp117} ? qw(ModelM::Dcp117) : (),
      $ruleset->{dcp118} ? qw(ModelM::Dcp118) : (),
      $ruleset->{edcm}   ? qw(ModelM::Edcm)   : ();
}

sub setUpMultiModelSharing {
    my ( $module, $mmsRef, $options, $oaRef ) = @_;
    require ModelM::MultiModel;
    $options->{multiModelSharing} = $$mmsRef ||= ModelM::MultiModel->new;
}

sub new {
    my $class = shift;
    my $model = bless { inputTables => [], @_ }, $class;

    my $allocLevelset = Labelset(
        name => 'Allocation levels',
        list => [ split /\n/, <<END_OF_LIST ] );
LV
HV/LV
HV
EHV&132
END_OF_LIST
    my ( $totalReturn, $totalDepreciation, $totalOperating, ) =
      $model->totalDpcr;
    my ( $revenue, $incentive, $pension, ) = $model->oneYearDpcr;
    my ( $units, ) = $model->units($allocLevelset);
    my ( $allocationRules, $capitalised, $directIndicator, ) = @{
        $model->{multiModelSharing}
        ? ( $model->{multiModelSharing}{optionsColumns} ||=
              $model->allocationRules )
        : $model->allocationRules
    };
    my ( $expenditure, ) = $model->expenditure( $allocationRules->{rows} );
    my ( $netCapexPercentages, ) = $model->netCapexPercentages($allocLevelset);
    my ( $meavPercentages, )     = $model->meavPercentages($allocLevelset);

    ( $netCapexPercentages, $meavPercentages, ) =
      $model->adjust118( $netCapexPercentages, $meavPercentages, )
      if $model->{dcp118};

    # because Numbers for iPad does not accept a SUMIF across sheets.
    $expenditure = Stack( sources => [$expenditure] );

    my ( $afterAllocation, $direct, $tableForColumnset ) =
      $model->expenditureAlloc(
        $allocLevelset,   $allocationRules, $capitalised,
        $directIndicator, $expenditure,     $meavPercentages,
      );

    push @{ $model->{calcTables} }, $allocationRules,
      Columnset(
        columns => [ $expenditure, $tableForColumnset ],
        name    => 'Expenditure data'
      );

    ( $afterAllocation, $allocLevelset, $netCapexPercentages, $units ) =
      $model->realloc95( $afterAllocation, $allocationRules,
        $netCapexPercentages, $units, )
      if $model->{dcp095};

    my ( $alloc, ) = $model->allocation(
        $afterAllocation,     $allocLevelset, $allocationRules,
        $capitalised,         $expenditure,   $incentive,
        $netCapexPercentages, $revenue,       $totalDepreciation,
        $totalOperating,      $totalReturn,   $units,
    );

    push @{ $model->{calcTables} }, $alloc,
      $model->{fixedIndirectPercentage} ? () : $direct;

    my $dcp071 = $model->{dcp071} || $model->{dcp071A};
    my ( $lvSplit, $hvSplit ) = $model->splits;
    my $discountCall = $model->{dcp095} ? 'discounts95' : 'discounts';
    $model->$discountCall( $alloc, $allocLevelset, $dcp071, $direct, $hvSplit,
        $lvSplit, );

    $model->discountEdcm( $alloc, $direct ) if $model->{edcm};

    $model;
}

1;
