package Elec::Tariffs;

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
use base 'Elec::TariffsBase';
use SpreadsheetModel::Shortcuts ':all';

sub tariffName {
    'use of system tariffs';
}

sub new {

    my ( $class, $model, $setup, $usage, $charging, $competitiveEnergy ) = @_;
    my $self = $model->register(
        bless {
            model => $model,
            setup => $setup,
        },
        $class
    );

    my @tariffContributions;
    my $usageRates       = $usage->usageRates;
    my $days             = $setup->daysInYear;
    my $tariffComponents = $setup->tariffComponents;
    my $digitsRounding   = $setup->digitsRounding;
    my @formatting       = map {
            !defined $digitsRounding->[$_] ? []
          : !$digitsRounding->[$_]     ? [ defaultFormat => '0soft' ]
          : $digitsRounding->[$_] == 2 ? [ defaultFormat => '0.00soft' ]
          :                              [];
    } 0 .. $#$tariffComponents;

    foreach my $charge ( $charging->charges ) {
        push @{ $model->{costTables} }, $charge;
        if ( $self->{model}{timebands} ) {
            push @tariffContributions, Columnset(
                name    => 'Contributions from ' . lcfirst( $charge->{name} ),
                columns => [
                    map {
                        my $usage   = $usageRates->[$_];
                        my $array   = ref $usage eq 'ARRAY';
                        my $contrib = Arithmetic(
                            name => 'Contributions from '
                              . lcfirst( $charge->{name} ) . ' to '
                              . lcfirst( $tariffComponents->[$_] ),
                            @{ $formatting[$_] },
                            arithmetic => '=A1*A2*100/A6'
                              . ( $array ? '/24*A3' : '' ),
                            rows => (
                                  $array ? $usageRates->[$_][0]
                                : $usageRates->[$_]
                              )->{rows},
                            arguments => {
                                $array
                                ? (
                                    A2 => $usageRates->[$_][0],
                                    A3 => $usageRates->[$_][1],
                                  )
                                : ( A2 => $usageRates->[$_] ),
                                A1 => $charge,
                                A6 => $days,
                            }
                        );
                        $contrib->lastCol
                          ? GroupBy(
                            name => 'Total contributions from '
                              . lcfirst( $charge->{name} ) . ' to '
                              . lcfirst( $tariffComponents->[$_] ),
                            @{ $formatting[$_] },
                            rows   => $contrib->{rows},
                            source => $contrib,
                          )
                          : $contrib;
                    } 0 .. $#$usageRates
                ],
            );
        }
        else {    # three columns, 0 is a unit rate, others are daily
            push @tariffContributions, Columnset(
                name    => 'Contributions from ' . lcfirst( $charge->{name} ),
                columns => [
                    map {
                        my $contrib = Arithmetic(
                            name => 'Contributions from '
                              . lcfirst( $charge->{name} ) . ' to '
                              . lcfirst( $tariffComponents->[$_] ),
                            @{ $formatting[$_] },
                            arithmetic => '=A1*A2*100/A666'
                              . ( $_ ? '' : '/24' ),
                            rows      => $usageRates->[$_]{rows},
                            arguments => {
                                A1   => $charge,
                                A2   => $usageRates->[$_],
                                A666 => $days,
                            }
                        );
                        $contrib->lastCol
                          ? GroupBy(
                            name => 'Total contributions from '
                              . lcfirst( $charge->{name} ) . ' to '
                              . lcfirst( $tariffComponents->[$_] ),
                            @{ $formatting[$_] },
                            rows   => $contrib->{rows},
                            source => $contrib,
                          )
                          : $contrib;
                    } 0 .. 2
                ],
            );
        }
    }
    push @{ $model->{buildupTables} }, @tariffContributions;
    $self->{tariffs} = [
        map {
            my $compno = $_;
            my @ingredients =
              grep { $_ } map { $_->{columns}[$compno] } @tariffContributions;
            my $formula = join '+', map { "A$_" } 1 .. @ingredients;
            Arithmetic(
                name       => $tariffComponents->[$_],
                arithmetic => defined $digitsRounding->[$_]
                ? "=ROUND($formula, $digitsRounding->[$_])"
                : "=$formula",
                arguments => {
                    map { ( "A$_" => $ingredients[ $_ - 1 ] ) }
                      1 .. @ingredients
                },
                @{ $formatting[$_] },
            );
        } 0 .. $#$tariffComponents
    ];

    $self;

}

sub finish {
    my ($self) = @_;
    $self->SUPER::finish;
    push @{ $self->{model}{tariffTables} },
      Columnset(
        name    => ucfirst( $self->tariffName ),
        columns => $self->{tariffs},
      );
}

1;
