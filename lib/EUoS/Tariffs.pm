package EUoS::Tariffs;

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
                  } 0 .. 2    # undue hardcoding (only zero is a unit rate)
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
                arithmetic => defined $digitsRounding->[$_]
                ? "=ROUND($formula, $digitsRounding->[$_])"
                : "=$formula",
                arguments => {
                    map { ( "IV$_" => $ingredients[ $_ - 1 ] ) }
                      1 .. @ingredients
                },
                !defined $digitsRounding->[$_] ? ()
                : !$digitsRounding->[$_]     ? ( defaultFormat => '0softnz' )
                : $digitsRounding->[$_] == 2 ? ( defaultFormat => '0.00softnz' )
                :                              (),
            );
          } 0 .. 2    # undue hardcoding
    ];
    $self;
}

sub revenueCalculation {
    my ( $self, $volumes, $labelTail, $name ) = @_;
    $labelTail ||= '';
    Arithmetic(
        name => $name
          || ( ucfirst( $self->tariffName ) . ' revenue £/year' . $labelTail ),
        arithmetic => '=(IV1*IV11+IV666*(IV2*IV12+IV3*IV13))/100',  # hard coded
        arguments  => {
            IV666 => $self->{setup}->daysInYear,
            map {
                (
                    "IV$_"  => $volumes->[ $_ - 1 ],
                    "IV1$_" => $self->{tariffs}[ $_ - 1 ]
                  )
            } 1 .. 3,
        },
        defaultFormat => '0softnz',
    );
}

sub revenues {
    my ( $self, $volumes, $compareppu, $notGrandTotal, $name, @extras ) = @_;
    my $labelTail =
      $volumes->[0]{usetName} ? " for $volumes->[0]{usetName}" : '';
    my $revenues = $self->revenueCalculation( $volumes, $labelTail, $name );

    if ($compareppu) {
        my $ppu = Arithmetic(
            name       => 'Average p/kWh',
            arithmetic => '=IF(IV3,IV1/IV2*100,"")',
            arguments  => {
                IV1 => $revenues,
                IV2 => $volumes->[0],
                IV3 => $volumes->[0],
            },
        );
        my $compare;
        $compare = Arithmetic(
            defaultFormat => '0soft',
            name          => 'Comparison £/year',
            arguments =>
              { IV1 => $compareppu, IV2 => $compareppu, IV3 => $volumes->[0] },
            arithmetic => '=IF(ISNUMBER(IV1),IV2*IV3*0.01,0)'
        ) if ref $compareppu;
        my $difference;
        $difference = Arithmetic(
            name          => 'Difference £/year',
            defaultFormat => '0softpm',
            arithmetic    => '=IF(IV1,IV2-IV3,"")',
            arguments     => {
                IV1 => $compare,
                IV2 => $revenues,
                IV3 => $compare,
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
                        arithmetic    => '=IF(IV1,IV2/IV3-1,"")',
                        arguments     => {
                            IV1 => $compare,
                            IV2 => $revenues,
                            IV3 => $compare,
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
                        defaultFormat => '0softnz',
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
                arithmetic    => '=IF(IV1,IV2/IV3,"")',
                arguments     => {
                    IV1 => $cols[0],
                    IV2 => $cols[ $#cols - 1 ],
                    IV3 => $cols[0],
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
            defaultFormat => '0softnz',
            source        => $revenues,
          );
    }
}

sub tariffs {
    my ($self) = @_;
    $self->{tariffs};
}

sub tariffName { 'use of system tariffs'; }

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
