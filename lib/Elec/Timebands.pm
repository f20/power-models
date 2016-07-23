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
    return $self->{timebandSet} if $self->{timebandSet};
    my @timebands;
    my @coreBand;
    my %id;
    foreach ( @{ $self->{model}{timebands} } ) {
        unless ( 'HASH' eq ref $_ ) {
            die "Duplicate time band $_" if exists $self->{timebandRules}{$_};
            push @timebands, $_;
            $coreBand[ $id{$_} = $#timebands ] = 1;
            return;
        }
        while ( my ( $set, $bands ) = each %$_ ) {
            my @sharedBand;
            my @ownBand;
            foreach (@$bands) {
                if ( my $id = $self->{timebandRules}{$_} ) {
                    $sharedBand[$id] = 1;
                }
                else {
                    push @timebands, $_;
                    $ownBand[ $id{$_} = $#timebands ] = 1;
                }
            }
            push @{ $self->{timebandSets} }, [ $set, \@sharedBand, \@ownBand ];
        }
    }
    unshift @{ $self->{timebandSets} }, [ 'Core time bands', [], \@coreBand ];
    $self->{timebandSet} = Labelset( name => 'Timebands', list => \@timebands );
}

sub hours {
    my ($self) = @_;
    return $self->{hours} if $self->{hours};
    my $timebandSet = $self->timebandSet;
    my $hours       = Dataset(
        name          => 'Typical annual hours by time band',
        defaultFormat => '0.0hard',
        rows          => $timebandSet,
        number        => 1568,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        data          => [ map { 1000 } @{ $timebandSet->{list} } ],
    );
    unless ( $self->{timebandSets} ) {
        my $totalHours = GroupBy(
            name          => 'Total annual hours across all time bands',
            defaultFormat => '0.0soft',
            source        => $hours
        );
        return $self->{hours} = Arithmetic(
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
    my @columns = ($hours);
    foreach ( @{ $self->{timebandSets} } ) {
        my ( $set, $sharedId, $ownId ) = @$_;
        push @columns,
          my $toScale = Constant(
            name          => "$set: to scale",
            rows          => $timebandSet,
            data          => [ \@data ],
            defaultFormat => '0con',
          );
    }
    Columnset(
        name    => 'Normalise hours to fit charging year',
        columns => \@columns,
    );
    $self->{hours} = $hours;
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
                [ map { .5 } @{ $self->timebandSet->{list} } ]
            } @{ $self->{setup}->usageSet->{list} }
        ],
    );
    unless ( $self->{timebandSets} ) {
        my $totalProb = GroupBy(
            name => Label(
                'Sum of peaking probabilities',
                'Sum of peaking probabilities (expected to be 100%)'
            ),
            cols          => $self->{setup}->usageSet,
            defaultFormat => '%soft',
            source        => $peakingProbabilities
        );
        return $self->{peakingProbabilities} = Arithmetic(
            name          => 'Normalised peaking probabilties',
            defaultFormat => '%soft',
            arithmetic    => '=IF(A2,A1/A3,A5/A6/24)',
            arguments     => {
                A1 => $peakingProbabilities,
                A2 => $totalProb,
                A3 => $totalProb,
                A5 => $self->hours,
                A6 => $self->{setup}->daysInYear,
            },
        );
    }
    die 'Not implemented';
}

sub bandFactors {
    my ($self) = @_;
    return $self->{bandFactors} if $self->{bandFactors};
    my $timebandSet          = $self->timebandSet;
    my $hours                = $self->hours;
    my $peakingProbabilities = $self->peakingProbabilities;
    my $targetUtilisation    = Dataset(
        name          => 'Target average capacity utilisation',
        defaultFormat => '%hard',
        rows          => $timebandSet,
        cols          => $self->{setup}->usageSet,
        number        => 1570,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        data          => [
            map {
                [ map { 1 } @{ $timebandSet->{list} } ]
            } @{ $self->{setup}->usageSet->{list} }
        ],
    );
    $self->{bandFactors} ||= [
        map {
            Arithmetic(
                name       => "$_ time band capacity contribution factors",
                rows       => Labelset( list => [$_] ),
                cols       => $self->{setup}->usageSet,
                arithmetic => '=A4*A6/A5/IF(A81,A82,1)',
                arguments  => {
                    A4  => $peakingProbabilities,
                    A5  => $hours->{arguments}{A1},
                    A6  => $hours->{arguments}{A2},
                    A81 => $targetUtilisation,
                    A82 => $targetUtilisation,
                },
            );
        } @{ $timebandSet->{list} }
    ];
}

sub finish {
    my ($self) = @_;
    push @{ $self->{model}{bandTables} }, $self->{hours} if $self->{hours};
    push @{ $self->{model}{bandTables} }, @{ $self->{bandFactors} }
      if $self->{bandFactors};
}

1;
