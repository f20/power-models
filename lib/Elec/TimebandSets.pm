package Elec::TimebandSets;

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
use base 'Elec::Timebands';
use SpreadsheetModel::Shortcuts ':all';

sub timebandConstants {
    my ($self) = @_;
    return @{ $self->{timebandConstants} } if $self->{timebandConstants};
    push @{ $self->{timebandConstants} },
      [
        $_->[0],
        Constant(
            name          => "$_->[0]: to deduct",
            rows          => $self->timebandSet,
            data          => [ $_->[1] ],
            defaultFormat => '0con',
        ),
        Constant(
            name          => "$_->[0]: to scale",
            rows          => $self->timebandSet,
            data          => [ $_->[2] ],
            defaultFormat => '0con',
        ),
      ]
      foreach @{ $self->{timebandSets} };
    @{ $self->{timebandConstants} };
}

sub timebandSet {
    my ($self) = @_;
    return $self->{timebandSet} if $self->{timebandSet};
    my @timebands;
    my @coreBand;
    my %id;
    foreach ( @{ $self->{model}{timebandSets} } ) {
        unless ( 'HASH' eq ref $_ ) {
            die "Duplicate time band $_" if exists $self->{timebandRules}{$_};
            push @timebands, $_;
            $coreBand[ $id{$_} = $#timebands ] = 1;
            next;
        }
        while ( my ( $set, $bands ) = each %$_ ) {
            my @sharedBand;
            my @ownBand;
            foreach (@$bands) {
                if ( my $id = $id{$_} ) {
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
    my $hours = Dataset(
        name          => 'Typical annual hours by time band',
        defaultFormat => '0.0hard',
        rows          => $self->timebandSet,
        number        => 1568,
        appendTo      => $self->{model}{inputTables},
        dataset       => $self->{model}{dataset},
        data          => [ map { 1000 } @{ $self->timebandSet->{list} } ],
    );
    my @calculationColumns;
    push @calculationColumns, $_->[1], $_->[2],
      $hours = Arithmetic(
        name          => 'Rescaled hours in ' . lcfirst $_->[0],
        defaultFormat => '0.0soft',
        arithmetic    => '=A1*IF(A3,'
          . '(24*A2-SUMPRODUCT(A21_A22,A11_A12))'
          . '/SUMPRODUCT(A31_A32,A13_A14),1)',
        arguments => {
            A1      => $hours,
            A11_A12 => $hours,
            A13_A14 => $hours,
            A2      => $self->{setup}->daysInYear,
            A21_A22 => $_->[1],
            A3      => $_->[2],
            A31_A32 => $_->[2],
        },
      ) foreach $self->timebandConstants;
    Columnset(
        name    => 'Normalise hours to fit charging year',
        columns => \@calculationColumns,
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
                [ map { 'spread' } @{ $self->timebandSet->{list} } ]
            } @{ $self->{setup}->usageSet->{list} }
        ],
    );

    foreach ( $self->timebandConstants ) {
        my $totalProbability = GroupBy(
            name          => "$_->[0]: total known peak probability",
            cols          => $self->{setup}->usageSet,
            defaultFormat => '%soft',
            source        => Arithmetic(
                name => "$_->[0]: relevant known peaking probabilities",
                defaultFormat => '%soft',
                arithmetic    => '=IF(AND(ISNUMBER(A1),A2+A3>0),A11,0)',
                arguments     => {
                    A1  => $peakingProbabilities,
                    A11 => $peakingProbabilities,
                    A2  => $_->[1],
                    A3  => $_->[2],
                },
            ),
        );
        my $relevantHours = Arithmetic(
            name          => "$_->[0]: hours for spreading",
            defaultFormat => '0.0soft',
            arithmetic    => '=IF(AND(NOT(ISNUMBER(A1)),A3>0),A2,0)',
            arguments     => {
                A1  => $peakingProbabilities,
                A11 => $peakingProbabilities,
                A2  => $self->hours,
                A3  => $_->[2],
            },
        );
        my $hoursToSpread = GroupBy(
            name          => "$_->[0]: total hours for spreading",
            cols          => $self->{setup}->usageSet,
            defaultFormat => '0.0soft',
            source        => $relevantHours,
        );
        $peakingProbabilities = Arithmetic(
            name          => "$_->[0]: filled in peaking probabilities",
            defaultFormat => '%soft',
            arithmetic    => '=IF(A41,IF(A31,(1-A5)*A32/A4,A1),'
              . 'IF(ABS(A51-1)<1e-7,A11,(1-A52)/A42))',
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
    }
    $self->{peakingProbabilities} = $peakingProbabilities;
}

1;
