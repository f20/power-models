package EUoS::Tariffs;

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
    my ( $class, $model, $setup, $usage, $charging ) = @_;
    my $self = bless { model => $model, setup => $setup }, $class;
    my @tariffContributions;
    my $usageRates       = $usage->usageRates;
    my $days             = $setup->daysInYear;
    my $tariffComponents = $setup->tariffComponents;
    my $digitsRounding   = $setup->digitsRounding;
    foreach my $charge ( $charging->charges ) {
        push @{ $model->{costTables} }, $charge;
        push @tariffContributions, Columnset(
            name    => "Contributions from $charge->{name}",
            columns => [
                map {
                    my $contrib = Arithmetic(
                        name => "Contributions from $charge->{name}"
                          . " to $tariffComponents->[$_]",
                        arithmetic => '=IV1*IV2*100/IV666'
                          . ( $_ ? '' : '/24' ),
                        rows      => $usageRates->[$_]{rows},
                        arguments => {
                            IV1   => $charge,
                            IV2   => $usageRates->[$_],
                            IV666 => $days,
                        }
                    );
                    $contrib->lastCol
                      ? GroupBy(
                        name => "Total contributions from $charge->{name}"
                          . " to $tariffComponents->[$_]",
                        rows   => $contrib->{rows},
                        source => $contrib,
                      )
                      : $contrib;
                } 0 .. 2
            ],
        );
    }
    push @{ $model->{buildupTables} }, @tariffContributions;
    $self->{tariffs} = [
        map {
            my $compno = $_;
            my @ingredients =
              grep { $_ } map { $_->{columns}[$compno] } @tariffContributions;
            my $formula = join '+', map { "IV$_" } 1 .. @ingredients;
            Arithmetic(
                name       => $tariffComponents->[$_],
                arithmetic => $digitsRounding->[$_]
                ? "=ROUND($formula, $digitsRounding->[$_])"
                : "=$formula",
                arguments => {
                    map { ( "IV$_" => $ingredients[ $_ - 1 ] ) }
                      1 .. @ingredients
                },
                !defined $digitsRounding->[$_] ? ()
                : !$digitsRounding->[$_]     ? ( defaultFormat => '0softnz' )
                : $digitsRounding->[$_] == 2 ? ( defaultFormat => '0.00softnz' )
                : (),
            );
        } 0 .. 2
    ];
    $self;
}

sub revenues {
    my ( $self, $volumes, $name, $omitGrandTotal ) = @_;
    my $labelTail =
      $volumes->[0]{usetName} ? " for $volumes->[0]{usetName}" : '';
    my $tariffs  = $self->{tariffs};
    my $revenues = Arithmetic(
        name => $name || ( 'Revenue £/year' . $labelTail ),
        arithmetic => '=(IV1*IV11+IV666*(IV2*IV12+IV3*IV13))/100',
        arguments  => {
            IV666 => $self->{setup}->daysInYear,
            map {
                (
                    "IV$_"  => $volumes->[ $_ - 1 ],
                    "IV1$_" => $tariffs->[ $_ - 1 ]
                  )
            } 1 .. 3,
        },
        defaultFormat => '0softnz',
    );
    push @{ $self->{revenueTables} }, $omitGrandTotal ? $revenues : GroupBy(
        name          => 'Total revenue £/year' . $labelTail,
        defaultFormat => '0softnz',
        source        => $revenues,
    );
}

sub finish {
    my ($self) = @_;
    push @{ $self->{model}{tariffTables} },
      Columnset( name => 'Tariffs', columns => $self->{tariffs}, );
    push @{ $self->{model}{revenueTables} }, @{ $self->{revenueTables} };
}

1;
