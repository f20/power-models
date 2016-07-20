package Elec::Comparison;

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
    my ( $class, $model, $omitTotal, $name, ) = @_;
    $model->register(
        bless {
            model     => $model,
            name      => $name,
            omitTotal => $omitTotal,
        },
        $class
    );
}

sub addComparisonPpu {
    my ( $self, $customers ) = @_;
    $customers->addColumns(
        $self->{comparisonppu} = Dataset(
            $self->{model}{table1653} ? () : ( number => 1599 ),
            appendTo => $self->{model}{inputTables},
            dataset  => $self->{model}{dataset},
            name     => 'Comparison p/kWh',
            rows     => $customers->userLabelset,
            data     => [ map { 10 } @{ $customers->userLabelset->{list} } ],
        )
    );
}

sub addComparisonTariff {
    my ( $self, $customers, $setup ) = @_;
    $self->{comparisonTariff} =
      Elec::TariffsBase->new( $self->{model}, $setup, $customers, 'Comparison',
        1599 );
}

sub revenueComparison {

    my ( $self, $tariff, $volumes, @extras ) = @_;
    my $labelTail =
      $volumes->[0]{usetName} ? " for $volumes->[0]{usetName}" : '';
    my $revenues =
      $tariff->revenueCalculation( $volumes, $labelTail, $self->{name} );

    my $totalUnits =
        $self->{model}{timebands}
      ? $volumes->[$#$volumes]
      : $volumes->[0];

    my $ppu = Arithmetic(
        name       => 'Average p/kWh',
        arithmetic => '=IF(A3,A1/A2*100,"")',
        arguments  => {
            A1 => $revenues,
            A2 => $totalUnits,
            A3 => $totalUnits,
        },
    );

    my ( $compare, $difference );
    if ( $self->{comparisonppu} ) {
        $compare = Arithmetic(
            defaultFormat => '0soft',
            name          => 'Comparison £/year',
            arguments     => {
                A1 => $self->{comparisonppu},
                A2 => $self->{comparisonppu},
                A3 => $totalUnits,
            },
            arithmetic => '=IF(ISNUMBER(A1),A2*A3*0.01,0)'
        );
    }
    elsif ( $self->{comparisonTariff} ) {
        $compare = $self->{comparisonTariff}->revenueCalculation($volumes);
    }

    $difference = Arithmetic(
        name          => 'Difference £/year',
        defaultFormat => '0softpm',
        arithmetic    => '=IF(A1,A2-A3,"")',
        arguments     => {
            A1 => $compare,
            A2 => $revenues,
            A3 => $compare,
        },
    ) if $compare;

    $self->{model}{detailedTablesNames} = $volumes->[0]{names};

    push @{ $self->{detailedTables} },
      Columnset(
        name    => 'Revenue (£/year) and average revenue (p/kWh)',
        columns => [
            $volumes->[0]{names}
            ? Stack( sources => [ $volumes->[0]{names} ] )
            : (),
            $revenues,
            @extras,
            $compare
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
                $self->{comparisonppu}
                ? Stack( sources => [ $self->{comparisonppu} ] )
                : Arithmetic(
                    name       => 'Comparison p/kWh',
                    arithmetic => '=IF(A3,A1/A2*100,"")',
                    arguments  => {
                        A1 => $compare,
                        A2 => $totalUnits,
                        A3 => $totalUnits,
                    }
                ),
              )
            : $ppu,
        ]
      );

    if ( !$self->{omitTotal} ) {
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
            $compare ? ( $compare, $difference, ) : ()
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
          ) if $compare;
        push @{ $self->{detailedTables} },
          Columnset(
            name    => 'Total £/year' . $labelTail,
            columns => \@cols,
          );
    }

}

sub finish {
    my ($self) = @_;
    push @{ $self->{model}{$_} }, @{ $self->{$_} }
      foreach grep { $self->{$_} } qw(revenueTables detailedTables);
}

1;
