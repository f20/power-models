package Elec::TariffsBase;

=head Copyright licence and disclaimer

Copyright 2012-2016 Franck Latremoliere, Reckon LLP and others.

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

sub tariffName {
    my ($self) = @_;
    $self->{name};
}

sub new {
    my ( $class, $model, $setup, $customers, $prefix, $number, ) = @_;
    my $rows = $customers->userLabelset;
    my @columns;
    push @columns, Dataset(
        name => ( $prefix ? "$prefix " . lcfirst($_) : $_ ) . ' p/kWh',
        rows => $rows,
        data       => [ map { 1 } @{ $rows->{list} } ],
        validation => {     # required to trigger lenient cell locking
            validate => 'any',
        },
    ) foreach $model->{timebands} ? @{ $model->{timebands} } : 'Units';
    push @columns, Dataset(
        name => ( $prefix ? "$prefix fixed" : 'Fixed' ) . ' p/day',
        rows => $rows,
        data       => [ map { 1 } @{ $rows->{list} } ],
        validation => {     # required to trigger lenient cell locking
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
        defaultFormat => '0.00hard',
    );
    push @columns, Dataset(
        name => ( $prefix ? "$prefix capacity" : 'Capacity' ) . ' p/kVA/day',
        rows => $rows,
        data       => [ map { 1 } @{ $rows->{list} } ],
        validation => {     # required to trigger lenient cell locking
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
        defaultFormat => '0.00hard',
    );
    push @columns, Dataset(
        name => ( $prefix ? "$prefix excess reactive" : 'Excess reactive' )
          . ' p/kVArh',
        rows => $rows,
        data => [ map { 1 } @{ $rows->{list} } ],
        validation => {    # required to trigger lenient cell locking
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    ) if $model->{reactive};
    $customers->addColumnset(
        name => $prefix ? "$prefix tariff" : 'Tariff',
        number   => $number,
        columns  => \@columns,
        appendTo => $model->{inputTables},
        dataset  => $model->{dataset},
    );
    $model->register(
        bless {
            model   => $model,
            setup   => $setup,
            name    => $prefix ? "$prefix tariffs" : 'tariffs',
            tariffs => \@columns,
        },
        $class
    );
}

sub revenueCalculation {
    my ( $self, $volumes, $labelTail ) = @_;
    $labelTail ||= '';
    return Arithmetic(
        rows => $volumes->[0]->{rows},
        name => ucfirst( $self->tariffName ) . ' revenue £/year' . $labelTail,
        arithmetic => '=('
          . join( '+',
            map    { "A1$_*A2$_"; }
              grep { $self->{tariffs}[$_]{name} !~ m#/day#i; }
              0 .. $#{ $self->{tariffs} } )
          . '+A6*('
          . join( '+',
            map    { "A1$_*A2$_"; }
              grep { $self->{tariffs}[$_]{name} =~ m#/day#i; }
              0 .. $#{ $self->{tariffs} } )
          . '))*0.01',
        arguments => {
            A6 => $self->{setup}->daysInYear,
            map { ( "A1$_" => $volumes->[$_], "A2$_" => $self->{tariffs}[$_] ) }
              0 .. $#{ $self->{tariffs} },
        },
        defaultFormat => '0soft',
    );
}

sub averageUnitRate {
    my ( $self, $volumes ) = @_;
    my $totalUnits =
        $self->{model}{timebands}
      ? $volumes->[$#$volumes]
      : $volumes->[0];
    return Arithmetic(
        rows       => $volumes->[0]->{rows},
        name       => 'Average unit rate p/kWh',
        arithmetic => '=IF(A51,('
          . join( '+',
            map    { "A1$_*A2$_"; }
              grep { $self->{tariffs}[$_]{name} !~ m#/day#i; }
              0 .. $#{ $self->{tariffs} } )
          . ')/A52,"")',
        arguments => {
            A6  => $self->{setup}->daysInYear,
            A51 => $totalUnits,
            A52 => $totalUnits,
            map { ( "A1$_" => $volumes->[$_], "A2$_" => $self->{tariffs}[$_] ) }
              0 .. $#{ $self->{tariffs} },
        },
    );
}

sub revenues {
    my ( $self, $volumes, $notGrandTotal, ) = @_;
    my $labelTail =
      $volumes->[0]{usetName} ? " for $volumes->[0]{usetName}" : '';
    my $revenues = $self->revenueCalculation( $volumes, $labelTail );
    unless ($notGrandTotal) {
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

sub finish {
    my ($self) = @_;
    push @{ $self->{model}{$_} }, @{ $self->{$_} }
      foreach grep { $self->{$_} } qw(revenueTables);
}

1;
