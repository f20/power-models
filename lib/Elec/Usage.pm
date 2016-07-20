package Elec::Usage;

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
    my ( $class, $model, $setup, $customers ) = @_;
   $model->register(bless {
        model     => $model,
        setup     => $setup,
        customers => $customers,
        $model->{usageTypes} ? ( usageTypes => $model->{usageTypes} ) : (),
        $model->{noEnergy}   ? ( noEnergy   => $model->{noEnergy} )   : (),
    }, $class);
}

sub usageTypes {
    my ($self) = @_;
    $self->{usageTypes} ||= [
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

sub usageRates {
    my ($self) = @_;
    return $self->{usageRates} if $self->{usageRates};
    my ( $model, $setup, $customers ) = @{$self}{qw(model setup customers)};

    # All cells are available; otherwise it would be insufficiently flexible.
    my $allBlank = [
        map {
            [ map { '' } $customers->tariffSet->indices ]
        } $self->usageSet->indices
    ];

    push @{ $model->{usageTables} },
      my @usageRates = (
        Dataset(
            name => 'Network usage of 1kW of '
              . ( $self->{model}{timebands} ? '' : 'average ' )
              . 'consumption',
            rows     => $customers->tariffSet,
            cols     => $self->usageSet,
            number   => 1531,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            data     => $allBlank,
        ),
        Dataset(
            name     => 'Network usage of an exit point',
            rows     => $customers->tariffSet,
            cols     => $self->usageSet,
            number   => 1532,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            data     => $allBlank,
        ),
        Dataset(
            name     => 'Network usage of 1kVA of agreed capacity',
            rows     => $customers->tariffSet,
            cols     => $self->usageSet,
            number   => 1533,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            data     => $allBlank,
        ),
      );

    if ( $self->{model}{timebands} ) {

        my $timebandSet =
          Labelset( name => 'Timebands', list => $self->{model}{timebands} );

        my $hours = Dataset(
            name          => 'Typical annual hours by time band',
            defaultFormat => '0.0hard',
            rows          => $timebandSet,
            number        => 1568,
            appendTo      => $model->{inputTables},
            dataset       => $model->{dataset},
            data          => [ map { 1000 } @{ $timebandSet->{list} } ],
        );

        $hours = Arithmetic(
            name          => 'Hours by time band in the charging year',
            defaultFormat => '0.0soft',
            arithmetic    => '=A1/A2*A6*24',
            arguments     => {
                A1 => $hours,
                A2 => GroupBy(
                    name          => 'Total annual hours across all time bands',
                    defaultFormat => '0.0soft',
                    source        => $hours
                ),
                A6 => $self->{setup}->daysInYear,
            },
        );

        push @{ $model->{bandTables} }, $hours->{arguments}{A2};

        my $peakingProbabilities = Dataset(
            name          => 'Peaking probabilities',
            defaultFormat => '%hard',
            rows          => $timebandSet,
            cols          => $self->usageSet,
            number        => 1560,
            appendTo      => $model->{inputTables},
            dataset       => $model->{dataset},
            data          => [
                map {
                    [ map { .5 } @{ $timebandSet->{list} } ]
                } @{ $self->usageSet->{list} }
            ],
        );

        my $totalProb = GroupBy(
            name => Label(
                'Sum of peaking probabilities',
                'Sum of peaking probabilities (expected to be 100%)'
            ),
            cols          => $self->usageSet,
            defaultFormat => '%soft',
            source        => $peakingProbabilities
        );

        $peakingProbabilities = Arithmetic(
            name          => 'Normalised peaking probabilties',
            defaultFormat => '%soft',
            arithmetic    => '=IF(A2,A1/A3,A5/A6)',
            arguments     => {
                A1 => $peakingProbabilities,
                A2 => $totalProb,
                A3 => $totalProb,
                A5 => $hours->{arguments}{A1},
                A6 => $hours->{arguments}{A2},
            },
        );

        my $targetUtilisation = Dataset(
            name          => 'Target average capacity utilisation',
            defaultFormat => '%hard',
            rows          => $timebandSet,
            cols          => $self->usageSet,
            number        => 1565,
            appendTo      => $model->{inputTables},
            dataset       => $model->{dataset},
            data          => [
                map {
                    [ map { 1 } @{ $timebandSet->{list} } ]
                } @{ $self->usageSet->{list} }
            ],
        );

        my $routeingFactor = shift @usageRates;
        my @usageRatesUnitRates;

        foreach my $band ( @{ $timebandSet->{list} } ) {

            my $bandFactor = Arithmetic(
                name       => "$band time band capacity contribution factors",
                rows       => Labelset( list => [$band] ),
                cols       => $self->usageSet,
                arithmetic => '=A4*A6/A5/IF(A81,A82,1)',
                arguments  => {
                    A4  => $peakingProbabilities,
                    A5  => $hours->{arguments}{A1},
                    A6  => $hours->{arguments}{A2},
                    A81 => $targetUtilisation,
                    A82 => $targetUtilisation,
                },
            );

            push @{ $model->{bandTables} }, $bandFactor;

            push @usageRatesUnitRates, [ $routeingFactor, $bandFactor ];

        }

        unshift @usageRates, @usageRatesUnitRates;

    }

    $self->{usageRates} = \@usageRates;

}

sub boundaryUsageSet {
    my ($self) = @_;
    $self->{boundaryUsageSet} ||= Labelset(
        name => 'Boundary usage',
        list => [ $self->usageSet->{list}[0] ]
    );
}

sub energyUsageSet {
    my ($self) = @_;
    die join '|', map { defined $_ ? $_ : 'undef'; } caller
      if $self->{noEnergy};
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
        list => [ @{$listr}[ 1 .. ( $#$listr - $self->{noEnergy} ? 0 : 1 ) ] ]
    );
}

sub totalUsage {
    my ( $self, $volumes ) = @_;
    return $self->{totalUsage}{ 0 + $volumes }
      if $self->{totalUsage}{ 0 + $volumes };
    my $labelTail =
      $volumes->[0]{usetName} ? " for $volumes->[0]{usetName}" : '';
    my $usageRates = $self->usageRates;
    my $customerUsage;
    if ( $self->{model}{timebands} ) {
        $customerUsage = Arithmetic(
            name       => 'Network usage' . $labelTail,
            rows       => $volumes->[0]{rows},
            cols       => $self->usageSet,
            arithmetic => '=('
              . join( '+', map { "A1$_*A2$_*A3$_"; } 3 .. $#$volumes )
              . ')/24/A6+'
              . join( '+', map { "A1$_*A2$_"; } 1 .. 2 ),
            arguments => {
                A6 => $self->{setup}->daysInYear,
                (
                    map {
                        (
                            "A1$_" => $usageRates->[ $#$volumes - $_ ],
                            "A2$_" => $volumes->[ $#$volumes - $_ ]
                        );
                    } 1 .. 2
                ),
                (
                    map {
                        (
                            "A1$_" => $usageRates->[ $#$volumes - $_ ][0],
                            "A3$_" => $usageRates->[ $#$volumes - $_ ][1],
                            "A2$_" => $volumes->[ $#$volumes - $_ ]
                        );
                    } 3 .. $#$volumes
                ),
            },
            defaultFormat => '0soft',
            names         => $volumes->[0]{names},
        );
    }
    else {    # three columns, 0 is a unit rate, others are daily
        $customerUsage = Arithmetic(
            name       => 'Network usage' . $labelTail,
            rows       => $volumes->[0]{rows},
            cols       => $usageRates->[0]{cols},
            arithmetic => '=' . join(
                '+',
                map {
                    my $m = $_ + 1;
                    my $v = $_ + 100;
                    "A$m*A$v" . ( $_ ? '' : '/24/A666' );
                } 0 .. 2
            ),
            arguments => {
                A666 => $self->{setup}->daysInYear,
                map {
                    my $m = $_ + 1;
                    my $v = $_ + 100;
                    (
                        "A$m" => $usageRates->[$_],
                        "A$v" => $volumes->[$_]
                    );
                } 0 .. 2
            },
            defaultFormat => '0soft',
            names         => $volumes->[0]{names},
        );
    }
    $self->{totalUsage}{ 0 + $volumes } = GroupBy(
        defaultFormat => '0soft',
        name          => 'Total network usage' . $labelTail,
        rows          => 0,
        cols          => $customerUsage->{cols},
        source        => $customerUsage,
    );
}

sub finish { }

1;
