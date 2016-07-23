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
    my ( $class, $model, $setup, $customers, $timebands ) = @_;
    $model->register(
        bless {
            model     => $model,
            setup     => $setup,
            customers => $customers,
            timebands => $timebands,
        },
        $class
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
        } $setup->usageSet->indices
    ];

    my $unitsRouteingFactor = Dataset(
        name => 'Network usage of 1kW of '
          . ( $self->{model}{timebands} ? '' : 'average ' )
          . 'consumption',
        rows     => $customers->tariffSet,
        cols     => $setup->usageSet,
        number   => 1531,
        appendTo => $model->{inputTables},
        dataset  => $model->{dataset},
        data     => $allBlank,
    );

    my @usageRates = (
        $self->{timebands}
        ? ( map { [ $unitsRouteingFactor, $_ ] }
              @{ $self->{timebands}->bandFactors } )
        : $unitsRouteingFactor,
        Dataset(
            name     => 'Network usage of an exit point',
            rows     => $customers->tariffSet,
            cols     => $setup->usageSet,
            number   => 1532,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            data     => $allBlank,
        ),
        Dataset(
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
    );

    push @{ $model->{usageTables} }, @usageRates;

    $self->{usageRates} = \@usageRates;

}

sub detailedUsage {
    my ( $self, $volumes ) = @_;
    return $self->{detailedUsage}{ 0 + $volumes }
      if $self->{detailedUsage}{ 0 + $volumes };
    my $labelTail =
      $volumes->[0]{usetName} ? " for $volumes->[0]{usetName}" : '';
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
