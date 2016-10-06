package Elec::Timebands;

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
    my ( $class, $model, $setup, ) = @_;
    $model->register(
        bless {
            model => $model,
            setup => $setup,
        },
        $class
    );
}

sub timebandSet {
    my ($self) = @_;
    return 0 unless $self->{model}{timebands};
    $self->{timebandSet} ||=
      Labelset( name => 'Timebands', list => $self->{model}{timebands} );
}

sub hours {
    my ($self) = @_;
    return $self->{hours} if $self->{hours};
    my $hours = Dataset(
        name          => 'Typical annual hours by time band',
        defaultFormat => '0.0hard',
        rows          => $self->timebandSet,
        number        => 1568,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        data          => [ map { 1000 } @{ $self->timebandSet->{list} } ],
    );
    my $totalHours = GroupBy(
        name          => 'Total annual hours across all time bands',
        defaultFormat => '0.0soft',
        source        => $hours
    );
    $self->{hours} = Arithmetic(
        name          => 'Hours by time band in the charging year',
        defaultFormat => '0.0soft',
        arithmetic    => '=A1/A2*A6*24',
        arguments     => {
            A1 => $hours,
            A2 => $totalHours,
            A6 => $self->{setup}->daysInYear,
        },
    );
}

sub peakingProbabilities {
    my ($self) = @_;
    return $self->{peakingProbabilities} if $self->{peakingProbabilities};
    my $peakingProbabilities = Dataset(
        name          => 'Peaking probabilities',
        defaultFormat => '%hard',
        rows          => $self->timebandSet,
        cols          => $self->{setup}->usageSet,
        number        => 1569,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        data          => [
            map {
                [ map { 'spread' } @{ $self->timebandSet->{list} } ]
            } @{ $self->{setup}->usageSet->{list} }
        ],
    );

    my $totalProbability = GroupBy(
        name          => 'Total known peak probability',
        cols          => $self->{setup}->usageSet,
        defaultFormat => '%soft',
        source        => Arithmetic(
            name          => 'Known peaking probabilities',
            defaultFormat => '%soft',
            arithmetic    => '=IF(ISNUMBER(A1),A11,0)',
            arguments     => {
                A1  => $peakingProbabilities,
                A11 => $peakingProbabilities,
            },
        ),
    );
    my $relevantHours = Arithmetic(
        name          => 'Hours for spreading',
        defaultFormat => '0.0soft',
        arithmetic    => '=IF(ISNUMBER(A1),0,A2)',
        arguments     => {
            A1 => $peakingProbabilities,
            A2 => $self->hours,
        },
    );
    my $hoursToSpread = GroupBy(
        name          => 'Total hours for spreading',
        cols          => $self->{setup}->usageSet,
        defaultFormat => '0.0soft',
        source        => $relevantHours,
    );
    $peakingProbabilities = Arithmetic(
        name          => 'Filled in peaking probabilities',
        defaultFormat => '%soft',
        arithmetic    => '=IF(A41,IF(A31,(1-A5)*A32/A4,A1),'
          . 'IF(ABS(A51-1)>1e-9,A11,(1-A52)/A42))',
        arguments => {
            A1  => $peakingProbabilities,
            A11 => $peakingProbabilities,
            A31 => $relevantHours,
            A32 => $relevantHours,
            A4  => $hoursToSpread,
            A41 => $hoursToSpread,
            A42 => $hoursToSpread,
            A5  => $totalProbability,
            A51 => $totalProbability,
            A52 => $totalProbability,
        },
    );
    $self->{peakingProbabilities} = $peakingProbabilities;
}

sub bandFactors {
    my ($self) = @_;
    return $self->{bandFactors} if $self->{bandFactors};
    my $targetUtilisation = Dataset(
        name          => 'Target average capacity utilisation',
        defaultFormat => '%hard',
        rows          => $self->timebandSet,
        cols          => $self->{setup}->usageSet,
        number        => 1570,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        data          => [
            map {
                [ map { 1 } @{ $self->timebandSet->{list} } ]
            } @{ $self->{setup}->usageSet->{list} }
        ],
    );
    $self->{bandFactors} ||= [
        map {
            Arithmetic(
                name       => "$_ time band capacity estimation factors",
                rows       => Labelset( list => [$_] ),
                cols       => $self->{setup}->usageSet,
                arithmetic => '=A4*A6*24/A5/IF(A81,A82,1)',
                arguments  => {
                    A4  => $self->peakingProbabilities,
                    A5  => $self->hours,
                    A6  => $self->{setup}->daysInYear,
                    A81 => $targetUtilisation,
                    A82 => $targetUtilisation,
                },
            );
        } @{ $self->timebandSet->{list} }
    ];
}

sub finish {
    my ($self) = @_;
    push @{ $self->{model}{bandTables} }, $self->{hours} if $self->{hours};
    push @{ $self->{model}{bandTables} }, @{ $self->{bandFactors} }
      if $self->{bandFactors};
}

1;
