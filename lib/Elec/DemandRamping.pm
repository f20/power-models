package Elec::DemandRamping;

# Copyright 2021 Franck Latrémolière and others.
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
use base 'Elec::Interpolator';
use SpreadsheetModel::Shortcuts ':all';
require Spreadsheet::WriteExcel::Utility;

sub demandRampingRowset {
    my ($self) = @_;
    $self->{demandRampingRowset} ||= Labelset(
        list => [ 1 .. ( $self->{model}{numRowsDemand} || 24 ) ],
        defaultFormat => 'thitem',
    );
}

sub demandRampingInputsAndFactors {

    my ($self) = @_;

    return $self->{demandRampingInputsAndFactors}
      if $self->{demandRampingInputsAndFactors};

    my $inputRowset = $self->demandRampingRowset;

    my $name = Dataset(
        name          => 'Name',
        rows          => $inputRowset,
        defaultFormat => 'texthard',
        data          => [ map { ''; } @{ $inputRowset->{list} } ],
    );
    my $tariff = Dataset(
        name          => 'Applicable tariff',
        rows          => $inputRowset,
        defaultFormat => 'texthard',
        data          => [ map { ''; } @{ $inputRowset->{list} } ],
    );
    my $startDate = Dataset(
        name          => 'Start date',
        rows          => $inputRowset,
        defaultFormat => 'datehard',
        data => [ map { '=DATE(2019,4,1)'; } @{ $inputRowset->{list} } ],
    );
    my $probability = Dataset(
        name          => 'Probability',
        rows          => $inputRowset,
        defaultFormat => '%hard',
        data          => [ map { 1; } @{ $inputRowset->{list} } ],
    );
    my $rampingPattern = Dataset(
        name          => 'Ramping pattern',
        rows          => $inputRowset,
        defaultFormat => [
            base       => 'stephard',
            num_format => '[Black]"Pattern #"0',
        ],
        data => [ map { 1; } @{ $inputRowset->{list} } ],
    );
    my ( $factorP, $factorC, $factorU ) =
      map { $self->rampingFactor( $_, $startDate, $rampingPattern ); }
      'supply points', 'capacity', 'units';
    my @demands = map {
        Dataset(
            rows          => $inputRowset,
            defaultFormat => '0hard',
            name          => $_,
            data          => [ map { ''; } @{ $inputRowset->{list} } ],
        );
    } @{ $self->{setup}->volumeComponents };
    $self->{demandInputColumns} =
      [ $name, $tariff, $startDate, $probability, $rampingPattern, @demands, ];
    $self->{demandRampingInputsAndFactors} =
      [ $tariff, $probability, $factorP, $factorC, $factorU, @demands, ];

}

sub rampingPatternRowset {
    my ($self) = @_;
    $self->{rampingPatterRowset} ||= Labelset(
        list => [ 1 .. ( $self->{model}{numRampingPatterns} || 4 ) ],
        defaultFormat => [
            base       => 'thitem',
            num_format => '[Black]"Pattern #"0',
        ]
    );
}

sub rampingFactor {

    my ( $self, $label, $startDate, $rampingPattern ) = @_;

    my $rowset = $self->rampingPatternRowset;

    my $startLevel = Dataset(
        name => "Initial level for $label (relative to ultimate level)",
        defaultFormat => '%hard',
        rows          => $rowset,
        data          => [ map { 0.5; } @{ $rowset->{list} } ]
    );

    my $startTime = Dataset(
        name          => "Start time for $label ramp (months)",
        defaultFormat => '0hard',
        rows          => $rowset,
        data          => [ map { 12; } @{ $rowset->{list} } ]
    );

    my $endTime = Dataset(
        name          => "End time for $label ramp (months)",
        defaultFormat => '0hard',
        rows          => $rowset,
        data          => [ map { 24; } @{ $rowset->{list} } ]
    );

    push @{ $self->{rampingPatternColumns} }, $startLevel, $startTime, $endTime;

    my $rampStartDate = Arithmetic(
        name          => "Starting date for $label ramping",
        defaultFormat => 'datesoft',
        arithmetic =>
          '=IF(A1,DATE(YEAR(A21),MONTH(A22)+INDEX(A3_A4,A11),DAY(A23)),A24)',
        arguments => {
            A1    => $rampingPattern,
            A11   => $rampingPattern,
            A21   => $startDate,
            A22   => $startDate,
            A23   => $startDate,
            A24   => $startDate,
            A3_A4 => $startTime,
        },
    );

    my $rampEndDate = Arithmetic(
        name          => "End date for $label ramping",
        defaultFormat => 'datesoft',
        arithmetic =>
          '=IF(A1,DATE(YEAR(A21),MONTH(A22)+INDEX(A3_A4,A11),DAY(A23)),A24)',
        arguments => {
            A1    => $rampingPattern,
            A11   => $rampingPattern,
            A21   => $startDate,
            A22   => $startDate,
            A23   => $startDate,
            A24   => $startDate,
            A3_A4 => $endTime,
        },
    );

    my $rampStartLevel = Arithmetic(
        name          => "Starting level for $label ramping",
        defaultFormat => '%soft',
        arithmetic    => '=IF(A1,INDEX(A3_A4,A11),1)',
        arguments     => {
            A1    => $rampingPattern,
            A11   => $rampingPattern,
            A3_A4 => $startLevel,
        },
    );

    my $daysInitialLevel = Arithmetic(
        name          => 'Days in year at initial level',
        defaultFormat => '0soft',
        arithmetic    => '=MAX(A52,MIN(A21,1+A61))-MIN(1+A62,MAX(A51,A1))',
        arguments     => {
            A1  => $startDate,
            A21 => $rampStartDate,
            A51 => $self->firstDay,
            A52 => $self->firstDay,
            A61 => $self->lastDay,
            A62 => $self->lastDay,
            A71 => $rampStartLevel,
        },
    );

    my $daysFinalLevel = Arithmetic(
        name          => 'Days in year at final level',
        defaultFormat => '0soft',
        arithmetic    => '=1+A62-MAX(A51,MIN(A31,1+A61))',
        arguments     => {
            A1  => $startDate,
            A11 => $startDate,
            A31 => $rampEndDate,
            A51 => $self->firstDay,
            A61 => $self->lastDay,
            A62 => $self->lastDay,
            A71 => $rampStartLevel,
        },
    );

    my $rampingDays = Arithmetic(
        name          => 'How many days of ramping in the year?',
        defaultFormat => '0soft',
        arithmetic    => '=MAX(A52,MIN(1+A61,A31))-MIN(1+A62,MAX(A51,A1))',
        arguments     => {
            A1  => $rampStartDate,
            A31 => $rampEndDate,
            A51 => $self->firstDay,
            A52 => $self->firstDay,
            A61 => $self->lastDay,
            A62 => $self->lastDay,
        },
    );

    my $earliestRampDay = Arithmetic(
        name => 'Which day of the ramp does ramping in the year start?',
        defaultFormat => '0soft',
        arithmetic    => '=IF(A1,MAX(A51-A11,0),0)',
        arguments     => {
            A1  => $rampingDays,
            A11 => $rampStartDate,
            A51 => $self->firstDay,
        },
    );

    my $factor = Arithmetic(
        name          => "Scaling factor for $label ramping",
        defaultFormat => '%soft',
        arithmetic    => '=(A1*A71+A2+'
          . 'A3*(A72+(1-A73)*(A4+(A31+1)/2)/(A6-A5+1)))/A9',
        arguments => {
            A1  => $daysInitialLevel,
            A2  => $daysFinalLevel,
            A3  => $rampingDays,
            A31 => $rampingDays,
            A4  => $earliestRampDay,
            A41 => $earliestRampDay,
            A5  => $rampStartDate,
            A6  => $rampEndDate,
            A71 => $rampStartLevel,
            A72 => $rampStartLevel,
            A73 => $rampStartLevel,
            A9  => $self->daysInYear,
        },
    );

    push @{ $self->{columnsets} },
      Columnset(
        name    => "Ramping calculations for $label",
        columns => [
            $rampStartDate,    $rampEndDate,    $rampStartLevel,
            $daysInitialLevel, $daysFinalLevel, $rampingDays,
            $earliestRampDay,  $factor,
        ]
      );

    $factor;

}

sub totalDemand {
    my ( $self, $usetName, $wantedRowset ) = @_;
    return $self->{totalDemand}{$wantedRowset}{$usetName}
      if $self->{totalDemand}{$wantedRowset}{$usetName};
    my ( $tariff, $probability, $factorP, $factorC, $factorU, @demands ) =
      @{ $self->demandRampingInputsAndFactors };
    my $prop;
    push @{ $self->{scenarioProportions} },
      $prop = Dataset(
        name          => "Proportion in $usetName",
        rows          => $self->demandRampingRowset,
        defaultFormat => '%hard',
        data          => [ map { 1; } @{ $self->demandRampingRowset->{list} } ],
      ) unless $usetName eq 'all users';
    my @columns = map {
        $self->forecastAggregator(
            $wantedRowset,
            undef,
            $prop ? [ $prop, $probability ] : $probability,
            $tariff,
            $_->{name} =~ /point/i      ? $factorP
            : $_->{name} =~ /capacity/i ? $factorC
            : $factorU,
            $_
        );
    } @demands;
    if ( $self->{setup}{timebands} ) {
        push @columns,
          Arithmetic(
            name          => 'Total units kWh',
            defaultFormat => '0soft',
            arithmetic    => '='
              . join( '+', map { "A$_"; } 1 .. $self->{setup}->timebandNumber ),
            arguments => {
                map { ( "A$_" => $columns[ $_ - 1 ] ); }
                  1 .. $self->{setup}->timebandNumber
            },
          );
    }
    push @{ $self->{model}{volumeTables} },
      Columnset(
        name    => "Forecast demands for $usetName",
        columns => \@columns,
      );
    $self->{totalDemand}{$wantedRowset}{$usetName} = \@columns;
}

sub finish {
    my ($self) = @_;
    $self->SUPER::finish;
    Columnset(
        name     => 'Ramping patterns',
        number   => 1517,
        appendTo => $self->{model}{inputTables},
        dataset  => $self->{model}{dataset},
        columns  => $self->{rampingPatternColumns},
    ) if $self->{rampingPatternColumns};
}

