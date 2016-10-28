package Elec::Setup;

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
use SpreadsheetModel::Shortcuts ':all';

sub new {
    my ( $class, $model ) = @_;
    $model->register( bless { model => $model }, $class );
}

sub daysInYear {
    my ($self) = @_;
    $self->{daysInYear} ||= Dataset(
        name       => 'Days in the charging year',
        validation => {
            validate => 'decimal',
            criteria => 'between',
            minimum  => 365,
            maximum  => 366,
        },
        data          => [365],
        defaultFormat => '0hard'
    );
}

sub annuityRate {
    my ($self) = @_;
    return $self->{annuityRate} if $self->{annuityRate};

    my $annuitisationPeriod = Dataset(
        name       => 'Annualisation period (years)',
        validation => {
            validate => 'decimal',
            criteria => 'between',
            minimum  => 0,
            maximum  => 999_999,
        },
        data          => [45],
        defaultFormat => '0hard',
    );

    my $rateOfReturn = Dataset(
        name       => 'Rate of return',
        validation => {
            validate      => 'decimal',
            criteria      => 'between',
            minimum       => 0,
            maximum       => 4,
            input_title   => 'Rate of return:',
            input_message => 'Percentage',
            error_message => 'The rate of return must be'
              . ' a non-negative percentage value.'
        },
        defaultFormat => '%hard',
        data          => [0.103]
    );

    $self->{annuityRate} = Arithmetic(
        name          => 'Annuity rate',
        defaultFormat => '%soft',
        arithmetic    => '=PMT(A1,A2,-1)',
        arguments     => {
            A1 => $rateOfReturn,
            A2 => $annuitisationPeriod,
        }
    );

}

sub registerTimebands {
    my ( $self, $timebands ) = @_;
    $self->{timebands} = $timebands;
}

sub timebandList {
    my ($self) = @_;
    $self->{timebands} ? @{ $self->{timebands}->timebandSet->{list} } : ('All');
}

sub timebandNumber {
    my ($self) = @_;
    $self->{timebandNumber} ||= () = $self->timebandList;
}

sub tariffComponents {
    my ($self) = @_;
    $self->{tariffComponents} ||= [
        ( map { "$_ p/kWh" } $self->timebandList ),
        'Fixed p/day',
        'Capacity p/kVA/day',
        $self->{model}{reactive} ? 'Excess reactive p/kVArh' : (),
    ];
}

sub digitsRounding {
    my ($self) = @_;
    $self->{model}{noRounding}
      ? []
      : [
        ( map { 3 } 1 .. $self->timebandNumber ),
        2, 2, $self->{model}{reactive} ? 3 : (),
      ];
}

sub volumeComponents {
    my ($self) = @_;
    $self->{volumeComponents} ||= [
        ( map { "$_ kWh" } $self->timebandList ),
        'Supply points',
        'Capacity kVA', $self->{model}{reactive} ? 'Excess reactive kVArh' : (),
    ];
}

sub usageTypes {
    my ($self) = @_;
    $self->{usageTypes} ||= $self->{model}{usageTypes}
      || [
        'Boundary capacity kVA',
        'Ring capacity kVA',
        'Transformer capacity kVA',
        'Low voltage network capacity kVA',
        'Metering switchgear for ring supply',
        'Low voltage metering switchgear',
        'Low voltage service 100 Amp',
        'Energy consumption kW',
      ];
}

sub usageSet {
    my ($self) = @_;
    $self->{usageSet} ||= Labelset(
        name => 'Network usage categories',
        list => $self->usageTypes,
    );
}

sub nonAssetUsageSet {
    my ($self) = @_;
    $self->{nonAssetUsageSet} ||= Labelset(
        name => 'Usage to allocate non-asset costs',
        list => [
            $self->usageSet->{list}[0],
            $self->usageSet->{list}[1]
              && $self->usageSet->{list}[1] =~
              /^(?:Boundary|Indirect|Non[ -]?asset)/i
            ? $self->usageSet->{list}[1]
            : (),
        ]
    );
}

sub energyUsageSet {
    my ($self) = @_;
    die join '|', map { defined $_ ? $_ : 'undef'; } caller
      if $self->{model}{noEnergy};
    my $listr = $self->usageSet->{list};
    $self->{energyUsageSet} ||= Labelset(
        name => 'Energy usage',
        list => [ $listr->[$#$listr] ]
    );
}

sub assetUsageSet {
    my ($self) = @_;
    my $listr = $self->usageSet->{list};
    $self->{assetUsageSet} ||= Labelset(
        name => 'Asset usage',
        list => [
            @{$listr}[
              @{ $self->nonAssetUsageSet->{list} }
              .. ( $#$listr - $self->{model}{noEnergy} ? 0 : 1 )
            ]
        ]
    );
}

sub finish {
    my ($self) = @_;
    return if $self->{generalInputDataTable};
    return
      unless my @columns = (
        $self->{daysInYear} || (),
        $self->{annuityRate}
        ? @{ $self->{annuityRate}{arguments} }{qw(A1 A2)}
        : ()
      );
    $self->{generalInputDataTable} |= Columnset(
        name     => 'Financial and general input data',
        number   => 1510,
        appendTo => $self->{model}{inputTables},
        dataset  => $self->{model}{dataset},
        columns  => \@columns,
    );
}

1;
