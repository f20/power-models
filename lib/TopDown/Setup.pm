package TopDown::Setup;

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

sub new {
    my ( $class, $model ) = @_;
    bless { model => $model }, $class;
}

sub daysInYear {
    my ($self) = @_;
    $self->{daysInYear} ||= Dataset(
        name       => 'Days in the charging year',
        validation => {
            validate => 'decimal',
            criteria => 'between',
            minimum  => 1,
            maximum  => 999,
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
        defaultFormat => '%softnz',
        arithmetic    => '=PMT(IV1,IV2,-1)',
        arguments     => {
            IV1 => $rateOfReturn,
            IV2 => $annuitisationPeriod,
        }
    );

}

sub tariffComponents {
    my ($self) = @_;
    $self->{tariffComponents} ||=
      [ 'Unit p/kWh', 'Fixed p/day', 'Capacity p/kVA/day', ];    # hard coded
}

sub digitsRounding {
    [ 3, 0, 2, ];
}

sub volumeComponents {
    my ($self) = @_;
    $self->{volumeComponents} ||=
      [ 'Units kWh', 'Supply points', 'Capacity kVA', ];
}

sub finish {
    my ($self) = @_;
    return if $self->{generalInputDataTable};
    return
      unless my @columns = (
        $self->{daysInYear} || (),
        $self->{annuityRate}
        ? @{ $self->{annuityRate}{arguments} }{qw(IV1 IV2)}
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
