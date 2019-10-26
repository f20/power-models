package Elec::Usage;

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
use SpreadsheetModel::Shortcuts ':all';

sub new {
    my ( $class, $model, $setup, $customers, $timebands, $suffix ) = @_;
    $model->register(
        bless {
            model     => $model,
            setup     => $setup,
            customers => $customers,
            timebands => $setup->{timebands},
            suffix    => $suffix,
        },
        $class
    );
}

sub usageRates {

    my ($self) = @_;
    return $self->{usageRates} if $self->{usageRates};
    my ( $model, $setup, $customers ) = @{$self}{qw(model setup customers)};

    my $allBlank = [
        map {
            [ map { '' } $customers->tariffSet->indices ]
        } $setup->usageSet->indices
    ];

    my $unitsRouteingFactor = Dataset(
        name => 'Network usage of 1kW of '
          . ( $self->{setup}{timebands} ? '' : 'average ' )
          . 'consumption',
        rows     => $customers->tariffSet,
        cols     => $setup->usageSet,
        number   => 1531,
        appendTo => $model->{inputTables},
        dataset  => $model->{dataset},
        data     => $allBlank,
    );

    $self->{usageRates} = [
        $self->{timebands}
        ? ( map { [ $unitsRouteingFactor, $_ ] }
              @{ $self->{timebands}->bandFactors } )
        : $unitsRouteingFactor,
        $model->{fixedUsageRules}
        ? Constant(
            name => 'Network usage of an exit point',
            rows => $customers->tariffSet,
            cols => $setup->usageSet,
            data => [
                map {
                    /33kV metering breaker/
                      ? [ map { /^33kV/ ? 1 : 0; }
                          @{ $customers->tariffSet->{list} } ]
                      : /HV metered source breaker/
                      ? [ map { /^HV Sub/ ? 1 : 0; }
                          @{ $customers->tariffSet->{list} } ]
                      : /HV metered secondary switchgear/
                      ? [ map { /^HV Sub/ ? 0 : /^HV/ ? 1 : 0; }
                          @{ $customers->tariffSet->{list} } ]
                      : /LV >100A metered service/
                      ? [ map { /^LV/ ? 1 : 0; }
                          @{ $customers->tariffSet->{list} } ]
                      : /LV <100A metered service/
                      ? [ map { /^Small LV/ ? 1 : 0; }
                          @{ $customers->tariffSet->{list} } ]
                      : [ map { 0; } $customers->tariffSet->indices ];
                } @{ $setup->usageSet->{list} }
            ],
          )
        : Dataset(
            name     => 'Network usage of an exit point',
            rows     => $customers->tariffSet,
            cols     => $setup->usageSet,
            number   => 1532,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            data     => $allBlank,
        ),
        $model->{fixedUsageRules}
        ? Constant(
            name => 'Network usage of 1kVA of agreed capacity',
            rows => $customers->tariffSet,
            cols => $setup->usageSet,
            data => [
                map {
                    /Indirect costs|Boundary charge|33kV$/
                      ? [ map { /^33kV/ ? 1 : /^HV Sub/ ? 1 : 0; }
                          @{ $customers->tariffSet->{list} } ]
                      : /33kV\/HV/ ? [ map { /^HV/ ? 1 : 0; }
                          @{ $customers->tariffSet->{list} } ]
                      : /^HV$/
                      ? [ map { /^HV Sub/ ? 0 : /^HV/ ? 1 : /^LV Sub/ ? 1 : 0; }
                          @{ $customers->tariffSet->{list} } ]
                      : /HV\/LV/ ? [ map { /^LV/ ? 1 : 0; }
                          @{ $customers->tariffSet->{list} } ]
                      : /^LV$/ ? [ map { /^LV Sub/ ? 0 : /^LV/ ? 1 : 0; }
                          @{ $customers->tariffSet->{list} } ]
                      : [ map { 0; } $customers->tariffSet->indices ];
                } @{ $setup->usageSet->{list} }
            ],
          )
        : Dataset(
            name     => 'Network usage of 1kVA of agreed capacity',
            rows     => $customers->tariffSet,
            cols     => $setup->usageSet,
            number   => 1533,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            data     => $allBlank,
        ),
        $self->{model}{reactive} ? Dataset(
            name     => 'Network usage of 1kVAr reactive consumption',
            rows     => $customers->tariffSet,
            cols     => $setup->usageSet,
            number   => 1534,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            data     => $allBlank,
        ) : (),
    ];

}

sub matchTotalUsage {

    my ( $self, $volumes ) = @_;
    my ( $model, $setup, $customers, $timebands ) =
      @{$self}{qw(model setup customers timebands)};

    my $targetUsage =
        $self->{model}{interpolator}
      ? $self->{model}{interpolator}->targetUsage( $self->{setup}->usageSet )
      : Dataset(
        name          => 'Target network usage',
        defaultFormat => '0hard',
        cols          => $self->{setup}->usageSet,
        number        => 1539,
        appendTo      => $model->{inputTables},
        dataset       => $model->{dataset},
        data          => [ map { ''; } $setup->usageSet->indices ],
      );

    my $adjustableCapacityUsageRate = $model->{fixedUsageRules}
      ? Constant(
        name => 'Adjustable element of network usage'
          . ' of 1kVA of agreed capacity',
        rows => $customers->tariffSet,
        cols => $setup->usageSet,
        data => [
            map {
                /Indirect costs|Boundary charge/
                  ? [ map { /^HV Sub/ ? 0 : /^HV/ ? 1 : /^LV/ ? 1 : 0; }
                      @{ $customers->tariffSet->{list} } ]
                  : [ map { 0; } $customers->tariffSet->indices ];
            } @{ $setup->usageSet->{list} }
        ],
      )
      : Dataset(
        name => 'Adjustable element of network usage'
          . ' of 1kVA of agreed capacity',
        rows     => $customers->tariffSet,
        cols     => $setup->usageSet,
        number   => 1537,
        appendTo => $model->{inputTables},
        dataset  => $model->{dataset},
        data     => [
            map {
                [ map { '' } $customers->tariffSet->indices ]
            } $setup->usageSet->indices
        ],
      );

    my $baselineUsage = $self->totalUsage($volumes);
    my $usageRates    = $self->usageRates;
    my ($capacityIndex) =
      grep {
        UNIVERSAL::isa( $usageRates->[$_], 'SpreadsheetModel::Dataset' )
          && $usageRates->[$_]{name} =~ /agreed capacity/i;
      } 0 .. $#$usageRates
      or return $self;

    my $adjustableUsage = SumProduct(
        name          => 'Total adjustable element of network usage',
        defaultFormat => '0soft',
        matrix        => $adjustableCapacityUsageRate,
        vector        => $volumes->[$capacityIndex],
    );

    my $factorBase = Arithmetic(
        name       => 'Factor to apply to base element of network usage',
        arithmetic => '=IF(ISNUMBER(A5),IF(A6,IF(A1<A3,1,A4/A2),1),1)',
        arguments  => {
            A1 => $baselineUsage,
            A2 => $baselineUsage,
            A3 => $targetUsage,
            A4 => $targetUsage,
            A5 => $targetUsage,
            A6 => $targetUsage,
        },
    );

    my $factorAdjE = Arithmetic(
        name       => 'Factor to apply to adjustable element of network usage',
        arithmetic => '=IF(A1,MAX(0,A3-A2)/A11,0)',
        arguments  => {
            A1  => $adjustableUsage,
            A11 => $adjustableUsage,
            A2  => $baselineUsage,
            A3  => $targetUsage,
        },
    );

    my $adjustedUsage =
      __PACKAGE__->new( $model, $setup, $customers, $timebands, ' (adjusted)' );

    $adjustedUsage->{usageRates} = [
        map {
            $_ == $capacityIndex
              ? Arithmetic(
                name => 'Adjusted '
                  . lcfirst( $usageRates->[$capacityIndex]->objectShortName ),
                arithmetic => '=A1*A2+A3*A4',
                arguments  => {
                    A1 => $usageRates->[$capacityIndex],
                    A2 => $factorBase,
                    A3 => $adjustableCapacityUsageRate,
                    A4 => $factorAdjE,
                },
              )
              : $usageRates->[$_];
        } 0 .. $#$usageRates
    ];

    $adjustedUsage;

}

sub detailedUsage {
    my ( $self, $volumes ) = @_;
    return $self->{detailedUsage}{ 0 + $volumes }
      if $self->{detailedUsage}{ 0 + $volumes };
    my $labelTail =
      $volumes->[0]{usetName} ? " for $volumes->[0]{usetName}" : '';
    $labelTail .= $self->{suffix} if defined $self->{suffix};
    my $usageRates = $self->usageRates;
    my @type =
      map {
            'ARRAY' eq ref $usageRates->[$_] ? 3
          : $_ == 0                          ? 1
          : $self->{model}{reactive} && $_ == $#$usageRates ? 1
          :                                                   0;
      } 0 .. $#$usageRates;
    $self->{detailedUsage}{ 0 + $volumes } = Arithmetic(
        name       => 'Network usage' . $labelTail,
        rows       => $volumes->[0]{rows},
        cols       => $self->{setup}->usageSet,
        arithmetic => '=('
          . join( '+',
            map { "A1$_*A2$_" . ( $type[$_] == 3 ? "*A3$_" : '' ); }
            grep { $type[$_]; } 0 .. $#type )
          . ')/24/A6+'
          . join( '+', map { "A1$_*A2$_"; } grep { !$type[$_]; } 0 .. $#type ),
        arguments => {
            A6 => $self->{setup}->daysInYear,
            map {
                $type[$_] == 3
                  ? (
                    "A1$_" => $usageRates->[$_][0],
                    "A3$_" => $usageRates->[$_][1],
                    "A2$_" => $volumes->[$_]
                  )
                  : (
                    "A1$_" => $usageRates->[$_],
                    "A2$_" => $volumes->[$_]
                  );
            } 0 .. $#type,
        },
        defaultFormat => '0soft',
        names         => $volumes->[0]{names},
    );
}

sub totalUsage {
    my ( $self, $volumes ) = @_;
    return $self->{totalUsage}{ 0 + $volumes }
      if $self->{totalUsage}{ 0 + $volumes };
    my $labelTail =
      $volumes->[0]{usetName} ? " for $volumes->[0]{usetName}" : '';
    $labelTail .= $self->{suffix} if defined $self->{suffix};
    $self->{totalUsage}{ 0 + $volumes } = GroupBy(
        defaultFormat => '0soft',
        name          => 'Total network usage' . $labelTail,
        rows          => 0,
        cols          => $self->detailedUsage($volumes)->{cols},
        source        => $self->detailedUsage($volumes),
    );
}

sub finish { }

1;
