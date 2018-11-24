package Elec::Comparison;

# Copyright 2012-2016 Franck Latrémolière, Reckon LLP and others.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice,
# this list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
# this list of conditions and the following disclaimer in the documentation
# and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
# EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
# THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Shortcuts ':all';

sub new {
    my ( $class, $model, $setup, $proportionUsed, $destinationTablesName, ) =
      @_;
    $model->register(
        bless {
            model                 => $model,
            setup                 => $setup,
            proportionUsed        => $proportionUsed,
            destinationTablesName => $destinationTablesName,
        },
        $class
    );
}

sub useAlternativeRowset {
    my ( $self, $rowset ) = @_;
    $self->{rows}           = $rowset;
    $self->{proportionUsed} = Stack(
        sources => [ $self->{proportionUsed} ],
        rows    => $rowset,
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

    my ( $self, $tariff, $volumes, $names, @extraColumns ) = @_;
    my $labelTail =
      $volumes->[0]{usetName} ? " for $volumes->[0]{usetName}" : '';
    my ( @srcCol, @columns );

    push @columns,
      Stack( $self->{rows} ? ( rows => $self->{rows} ) : (),
        sources => [$names] )
      if $names;
    $self->{model}{detailedTablesNames} = $names;

    my $totalUnits =
        $self->{setup}{timebands}
      ? $volumes->[$#$volumes]
      : $volumes->[0];
    push @columns,
      $totalUnits = Stack( rows => $self->{rows}, sources => [$totalUnits] )
      if $self->{rows};

    push @columns, @extraColumns;

    my $revenues = $tariff->revenueCalculation($volumes);
    if ( $self->{rows} ) {
        push @srcCol, $revenues;
        $revenues = Stack( rows => $self->{rows}, sources => [$revenues] );
    }
    push @columns, $revenues;

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

    if ($compare) {
        if ( $self->{rows} ) {
            push @srcCol, $compare;
            $compare = Stack( rows => $self->{rows}, sources => [$compare] );
        }
        push @columns, $compare, $difference = Arithmetic(
            name          => 'Difference £/year',
            defaultFormat => '0softpm',
            arithmetic    => '=A2-A1',              # '=IF(A1,A2-A3,"")',
            arguments     => {
                A1 => $compare,
                A2 => $revenues,
                A3 => $compare,
            },
          ),
          Arithmetic(
            name          => 'Difference %',
            defaultFormat => '%softpm',
            arithmetic    => '=IF(A1,A2/A3-1,"")',
            arguments     => {
                A1 => $compare,
                A2 => $revenues,
                A3 => $compare,
            },
          );
    }

    push @columns,
      my $ppu = Arithmetic(
        name       => 'Average p/kWh',
        arithmetic => '=IF(A3,A1/A2*100,"")',
        arguments  => {
            A1 => $revenues,
            A2 => $totalUnits,
            A3 => $totalUnits,
        },
      );

    push @columns,
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
      ) if $compare;

    push @columns, $self->{proportionUsed}
      if $self->{rows} && $self->{proportionUsed};

    push @{ $self->{revenueTables} },
      Columnset(
        name => 'Revenue'
          . ( $compare ? ' comparison' : '' )
          . ' (£/year)'
          . $labelTail,
        columns => \@srcCol,
      ) if @srcCol;

    push @{ $self->{detailedTables} },
      Columnset(
        name    => 'Revenue (£/year) and average revenue (p/kWh)' . $labelTail,
        columns => \@columns,
      );

    foreach my $groupedRows (
        $revenues->{rows}{groups}
        ? Labelset( list => $revenues->{rows}{groups} )
        : (),
        undef
      )
    {
        my $totalTerm = $groupedRows ? 'Subtotal' : 'Total';
        my @cols = (
            map {
                my $n =
                  $totalTerm . ' '
                  . lcfirst(
                    SpreadsheetModel::Object::_shortName( $_->{name} ) );
                $self->{proportionUsed}
                  ? SumProduct(
                    name          => $n,
                    rows          => $groupedRows,
                    defaultFormat => $_->{defaultFormat},
                    matrix        => $self->{proportionUsed},
                    vector        => $_,
                  )
                  : GroupBy(
                    name          => $n,
                    rows          => $groupedRows,
                    defaultFormat => $_->{defaultFormat},
                    source        => $_,
                  );
              } $totalUnits,
            $revenues,
            @extraColumns,
            $compare ? ( $compare, $difference, ) : ()
        );
        push @cols,
          Arithmetic(
            name          => "$totalTerm difference %",
            defaultFormat => '%softpm',
            arithmetic    => '=IF(A1,A2/A3,"")',
            arguments     => {
                A1 => $cols[ $#cols - 1 ],
                A2 => $cols[$#cols],
                A3 => $cols[ $#cols - 1 ],
            },
          ) if $compare;
        push @cols,
          Arithmetic(
            name       => 'Average p/kWh',
            arithmetic => '=IF(A3,A1/A2*100,"")',
            arguments  => {
                A1 => $cols[1],
                A2 => $cols[0],
                A3 => $cols[0],
            }
          );
        push @cols,
          Arithmetic(
            name       => 'Comparison p/kWh',
            arithmetic => '=IF(A3,A1/A2*100,"")',
            arguments  => {
                A1 => $cols[ $#cols - 3 ],
                A2 => $cols[0],
                A3 => $cols[0],
            }
          ) if $compare;
        push @{ $self->{detailedTables} },
          Columnset(
            name    => "$totalTerm £/year$labelTail",
            columns => \@cols,
          );
    }

}

sub finish {
    my ($self) = @_;
    push @{ $self->{model}{ $self->{destinationTablesName} || $_ } },
      @{ $self->{$_} }
      foreach grep { $self->{$_} } qw(revenueTables detailedTables);
}

1;
