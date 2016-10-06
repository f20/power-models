package ModelM;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others.
Copyright 2012-2015 Franck Latrémolière, Reckon LLP and others.

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

sub requiredModulesForRuleset {

    my ( $class, $ruleset ) = @_;

    my @modules = (

        'ModelM::Allocation',
        'ModelM::Expenditure',
        'ModelM::Inputs',
        'ModelM::Meav',
        'ModelM::NetCapex',
        'ModelM::Options',
        'ModelM::Sheets',
        'ModelM::Units',

          $ruleset->{dcp095} ? 'ModelM::Dcp095'
        : $ruleset->{edcm} && $ruleset->{edcm} =~ /only/ ? ()
        : 'ModelM::Discounts',

        $ruleset->{dcp118} ? 'ModelM::Dcp118' : (),

        $ruleset->{edcm} ? 'ModelM::Edcm' : (),

        $ruleset->{dcp117}
          && $ruleset->{dcp117} !~ /201[34]/ ? 'ModelM::Dcp117_2012' : (),

        $ruleset->{checksums} ? 'SpreadsheetModel::Checksum' : ()

    );

    push @modules, $class->requiredModulesForRuleset($_)
      foreach ref $ruleset->{AdditionalRules} eq 'ARRAY'
      ? @{ $ruleset->{AdditionalRules} }
      : ref $ruleset->{AdditionalRules} eq 'HASH' ? $ruleset->{AdditionalRules}
      :                                             ();

    @modules;

}

sub new {
    my $class = shift;
    my $model = bless {
        objects => { inputTables => [] },
        dataset => {},
        @_,
    }, $class;

    die 'Cannot build an orange model without a suitable disclaimer'
      if $model->{colour}
      && $model->{colour} =~ /orange/i
      && !($model->{extraNotice}
        && length( $model->{extraNotice} ) > 299
        && $model->{extraNotice} =~ /DCUSA/ );

    my $extras = delete $model->{AdditionalRules};
    $model->run;
    foreach (
          ref $extras eq 'ARRAY' ? @$extras
        : ref $extras eq 'HASH'  ? $extras
        :                          ()
      )
    {
        (
            bless {
                objects => $model->{objects},
                dataset => $model->{dataset},
                %$_,
            },
            $class
        )->run;
    }
    $model;
}

sub setUpMultiModelSharing {
    my ( $module, $mmsRef, $options, $oaRef ) = @_;
    require ModelM::MultiModel;
    $options->{multiModelSharing} = $$mmsRef ||= ModelM::MultiModel->new;
}

sub run {
    my ($model) = @_;
    my $allocLevelset = $model->{objects}{allocLevelset} ||= Labelset(
        name => 'Allocation levels',
        list => [qw(LV HV/LV HV EHV&132)]
    );
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

    my ( $afterAllocation, $direct, ) = $model->expenditureAlloc(
        $allocLevelset,   $allocationRules, $capitalised,
        $directIndicator, $expenditure,     $meavPercentages,
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

    unless ( $model->{edcm} && $model->{edcm} =~ /only/ ) {
        my $dcp071 = $model->{dcp071} || $model->{dcp071A};
        my $discountCall = $model->{dcp095} ? 'discounts95' : 'discounts';
        $model->$discountCall( $alloc, $allocLevelset, $dcp071, $direct,
            $model->hvSplit, $model->lvSplit, );
    }

    $model->discountEdcm( $alloc, $direct ) if $model->{edcm};

}

1;
