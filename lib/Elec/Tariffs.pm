package Elec::Tariffs;

=head Copyright licence and disclaimer

Copyright 2012-2014 Franck Latrémolière, Reckon LLP and others.

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

    my ( $class, $model, $setup, $usage, $charging, $competitiveEnergy ) = @_;
    my $self = bless {
        model => $model,
        setup => $setup,
    }, $class;

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

sub revenueCalculation {
    my ( $self, $volumes, $labelTail, $name ) = @_;
    $labelTail ||= '';
    if ( $self->{model}{timebands} ) {
        return Arithmetic(
            rows => $volumes->[0]->{rows},
            name => $name
              || (
                ucfirst( $self->tariffName ) . ' revenue £/year' . $labelTail ),
            arithmetic => '=('
              . join( '+', map { "A1$_*A2$_"; } 3 .. $#$volumes )
              . '+A6*(A11*A21+A12*A22))*0.01',
            arguments => {
                A6 => $self->{setup}->daysInYear,
                map {
                    (
                        "A1$_" => $volumes->[ $#$volumes - $_ ],
                        "A2$_" => $self->{tariffs}[ $#$volumes - $_ ]
                      )
                } 1 .. $#$volumes,
            },
            defaultFormat => '0soft',
        );
    }
    else {    # three columns, 0 is a unit rate, others are daily
        return Arithmetic(
            name => $name
              || (
                ucfirst( $self->tariffName ) . ' revenue £/year' . $labelTail ),
            arithmetic => '=(A1*A11+A666*(A2*A12+A3*A13))*0.01',
            arguments  => {
                A666 => $self->{setup}->daysInYear,
                map {
                    (
                        "A$_"  => $volumes->[ $_ - 1 ],
                        "A1$_" => $self->{tariffs}[ $_ - 1 ]
                      )
                } 1 .. 3,
            },
            defaultFormat => '0soft',
        );
    }
}

sub revenues {
    my ( $self, $volumes, $compareppu, $notGrandTotal, $name, @extras ) = @_;
    my $labelTail =
      $volumes->[0]{usetName} ? " for $volumes->[0]{usetName}" : '';
    my $revenues = $self->revenueCalculation( $volumes, $labelTail, $name );

    if ($compareppu) {
        my $ppu = Arithmetic(
            name       => 'Average p/kWh',
            arithmetic => '=IF(A3,A1/A2*100,"")',
            arguments  => {
                A1 => $revenues,
                $self->{model}{timebands}
                ? (
                    A2 => $volumes->[$#$volumes],
                    A3 => $volumes->[$#$volumes],
                  )
                : (
                    A2 => $volumes->[0],
                    A3 => $volumes->[0],
                )
            },
        );

        my $compare;
        if ( ref $compareppu ) {
            $compare = Arithmetic(
                defaultFormat => '0soft',
                name          => 'Comparison £/year',
                arguments     => {
                    A1 => $compareppu,
                    A2 => $compareppu,
                    A3 => $self->{model}{timebands}
                    ? $volumes->[$#$volumes]
                    : $volumes->[0],
                },
                arithmetic => '=IF(ISNUMBER(A1),A2*A3*0.01,0)'
            );
        }

        my $difference;
        $difference = Arithmetic(
            name          => 'Difference £/year',
            defaultFormat => '0softpm',
            arithmetic    => '=IF(A1,A2-A3,"")',
            arguments     => {
                A1 => $compare,
                A2 => $revenues,
                A3 => $compare,
            },
        ) if ref $compareppu;
        push @{ $self->{detailedTables} },
          Columnset(
            name    => 'Revenue (£/year) and average revenue (p/kWh)',
            columns => [
                $volumes->[0]{names}
                ? Stack( sources => [ $volumes->[0]{names} ] )
                : (),
                $revenues,
                @extras,
                ref $compareppu
                ? (
                    $compare,
                    $difference,
                    Arithmetic(
                        name          => 'Difference %',
                        defaultFormat => '%softpm',
                        arithmetic    => '=IF(A1,A2/A3-1,"")',
                        arguments     => {
                            A1 => $compare,
                            A2 => $revenues,
                            A3 => $compare,
                        },
                    ),
                    $ppu,
                    Stack( sources => [$compareppu] ),
                  )
                : $ppu,
            ]
          );

        if ( ref $compareppu && !$notGrandTotal ) {
            my @cols = (
                map {
                    my $n = 'Total ' . $_->{name}->shortName;
                    $n =~ s/Total (.)/ 'Total '.lc($1)/e;
                    GroupBy(
                        name          => $n,
                        defaultFormat => '0soft',
                        source        => $_,
                    );
                  } $revenues,
                @extras,
                $compare,
                $difference,
            );
            push @cols,
              Arithmetic(
                name          => 'Total difference %',
                defaultFormat => '%softpm',
                arithmetic    => '=IF(A1,A2/A3,"")',
                arguments     => {
                    A1 => $cols[0],
                    A2 => $cols[ $#cols - 1 ],
                    A3 => $cols[0],
                },
              );
            push @{ $self->{detailedTables} },
              Columnset(
                name    => 'Total £/year' . $labelTail,
                columns => \@cols,
              );
        }
    }

    elsif ( !$notGrandTotal ) {
        push @{ $self->{revenueTables} },
          GroupBy(
            name => 'Total '
              . $self->tariffName
              . ' revenue £/year'
              . $labelTail,
            singleRowName => 'Total',
            defaultFormat => '0soft',
            source        => $revenues,
          );
    }
}

sub tariffs {
    my ($self) = @_;
    $self->{tariffs};
}

sub tariffName {
    'use of system tariffs';
}

sub finish {
    my ($self) = @_;
    push @{ $self->{model}{tariffTables} },
      Columnset(
        name    => ucfirst( $self->tariffName ),
        columns => $self->{tariffs},
      );
    push @{ $self->{model}{$_} }, @{ $self->{$_} }
      foreach grep { $self->{$_} } qw(revenueTables detailedTables);
}

1;
