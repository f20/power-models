package Elec::Customers;

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
    my ( $class, $model, $setup ) = @_;
    $model->register(
        bless {
            model => $model,
            setup => $setup,
        },
        $class
    );
}

sub totalDemand {
    my ( $self, $usetName ) = @_;
    return $self->{totalDemand}{$usetName} if $self->{totalDemand}{$usetName};
    my $tariffSet = $self->tariffSet;
    return $self->{model}{interpolator}->totalDemand( $usetName, $tariffSet )
      if $self->{model}{interpolator};
    my $detailedVolumes = $self->detailedVolumes;
    push @{ $self->{scenarioProportions} }, my $prop = Dataset(
        name => 'Proportion '
          . (
            $usetName eq 'all users' ? 'taken into account' : "in $usetName" ),
        rows          => $self->userLabelsetForInput,
        defaultFormat => '%hard',
        data          => [ map { 1; } @{ $self->userLabelset->{list} } ],
        validation => {    # required to trigger lenient cell locking
            validate      => 'decimal',
            criteria      => 'between',
            minimum       => -1,
            maximum       => 1,
            input_message => 'Percentage',
        },
    );
    my @columns =
      map {
        SumProduct(
            name          => $_->{name},
            matrix        => $prop,
            vector        => $_,
            rows          => $tariffSet,
            usetName      => $usetName,
            defaultFormat => '0soft',
        );
      } @$detailedVolumes;
    if ( $self->{setup}{timebands} ) {
        push @columns,
          Arithmetic(
            name          => 'Total units kWh',
            defaultFormat => '0soft',
            arithmetic    => '='
              . join( '+', map { "A$_"; } 1 .. $self->{setup}->timebandNumber ),
            arguments => {
                map { ( "A$_" => $columns[ $_ - 1 ] ); }
                  1 .. $self->{setup}->timebandNumber
            },
          );
    }
    push @{ $self->{model}{volumeTables} },
      Columnset(
        name    => "Forecast volumes for $usetName",
        columns => \@columns,
      );
    $self->{totalDemand}{$usetName} = \@columns;
}

sub individualDemandUsed {
    my ( $self, $usetName ) = @_;
    return $self->{individualDemandUsed}{$usetName}
      if $self->{individualDemandUsed}{$usetName};
    my @columns =
      map {
        Arithmetic(
            name          => $_->{name},
            arguments     => { A1 => $_->{matrix}, A2 => $_->{vector}, },
            arithmetic    => '=A1*A2',
            defaultFormat => '0soft',
        );
      } grep { UNIVERSAL::isa( $_, 'SpreadsheetModel::SumProduct' ); }
      @{ $self->totalDemand($usetName) };
    if ( $self->{setup}{timebands} ) {
        push @columns,
          Arithmetic(
            name          => 'Total units kWh',
            defaultFormat => '0soft',
            arithmetic    => '='
              . join( '+', map { "A$_"; } 1 .. $self->{setup}->timebandNumber ),
            arguments => {
                map { ( "A$_" => $columns[ $_ - 1 ] ); }
                  1 .. $self->{setup}->timebandNumber
            },
          );
    }
    push @{ $self->{model}{volumeTables} },
      Columnset(
        name    => "Individual customer volumes in $usetName",
        columns => \@columns,
      );
    $self->{individualDemandUsed}{$usetName} = \@columns;
}

sub individualDemand {
    my ($self) = @_;
    return $self->{individualDemand} if $self->{individualDemand};
    my @columns = @{ $self->detailedVolumes };
    if ( $self->{setup}{timebands} ) {
        push @columns,
          my $total = Arithmetic(
            name          => 'Total units kWh',
            defaultFormat => '0soft',
            arithmetic    => '='
              . join( '+', map { "A$_"; } 1 .. $self->{setup}->timebandNumber ),
            arguments => {
                map { ( "A$_" => $columns[ $_ - 1 ] ); }
                  1 .. $self->{setup}->timebandNumber
            },
          );
        push @{ $self->{model}{volumeTables} }, $total;
    }
    $self->{individualDemand} = \@columns;
}

sub userLabelset {
    my ($self) = @_;
    return $self->{userLabelset} if $self->{userLabelset};
    return $self->{userLabelset} = Labelset(
        name   => 'Detailed list of customers',
        groups => [
            map { Labelset( name => keys %$_, list => values %$_ ); }
              @{ $self->{model}{ulist} }
        ]
    ) if $self->{model}{ulist};
    my $cat          = 0;
    my $userLabelset = Labelset(
        name   => 'Detailed list of customers',
        groups => [
            map {
                $cat += 10_000;
                my ( $name, $count ) = each %$_;
                Labelset(
                    name => $name,
                    list => [ map { "User " . ( $cat + $_ ); } 1 .. $count ]
                );
            } @{ $self->{model}{ucount} }
        ]
    );
    return $self->{userLabelset} = $userLabelset
      unless $self->{model}{table1653};
    $self->{userLabelsetForInput} = $userLabelset;
    $self->{namesInLabelset}      = Dataset(
        defaultFormat => 'texthard',
        data          => [ map { '' } @{ $userLabelset->{list} } ],
        name          => 'Name',
        rows          => $userLabelset,
        validation => {    # required to trigger lenient cell locking
            validate => 'any',
        },
    );
    $self->{userLabelset} = Labelset(
        name     => 'Editable labelset for customer names',
        editable => $self->{namesInLabelset}
    );
}

sub userLabelsetForInput {
    my ($self) = @_;
    my $lset = $self->userLabelset;
    $self->{userLabelsetForInput} || $lset;
}

sub volumeDataColumn {
    my ( $self, $component ) = @_;
    [ map { 0 } @{ $self->userLabelset->{list} } ];
}

sub detailedVolumes {
    my ($self) = @_;
    return $self->{detailedVolumes} if $self->{detailedVolumes};
    $self->{detailedVolumes} = [
        map {
            Dataset(
                rows          => $self->userLabelsetForInput,
                defaultFormat => '0hard',
                name          => $_,
                data          => $self->volumeDataColumn($_),
                validation => {   # required to trigger leniency in cell locking
                    validate => 'decimal',
                    criteria => '>=',
                    value    => 0,
                },
            );
        } @{ $self->{setup}->volumeComponents }
    ];
}

sub tariffSet {
    my ($self) = @_;
    $self->{tariffSet} ||= Labelset(
        name => 'Set of customer categories',
        list => $self->userLabelset->{groups},
    );
}

sub names {
    $_[0]{names};
}

sub addColumns {
    my ( $self, @columns ) = @_;
    push @{ $self->{extraColumns} }, @columns;
}

sub addColumnset {
    my ( $self, %columnsetContents ) = @_;
    if ( $self->{model}{table1653} ) {
        $self->addColumns( @{ $columnsetContents{columns} } );
    }
    else {
        Columnset(%columnsetContents);
    }
}

sub finish {
    my ( $self, $model ) = @_;
    if (   $model->{table1653}
        && $self->{detailedVolumes}
        && $self->{scenarioProportions} )
    {
        $model->{table1653Names} = $self->{names} || $self->{namesInLabelset};
        $model->{table1653} = Columnset(
            name     => 'Individual user data',
            number   => 1653,
            location => 'Customers',
            dataset  => $model->{dataset},
            columns  => [
                $model->{table1653Names} ? $model->{table1653Names} : (),
                @{ $self->{scenarioProportions} },
                @{ $self->{detailedVolumes} },
                $self->{extraColumns} ? @{ $self->{extraColumns} } : (),
            ],
            ignoreDatasheet => 1,
        );
    }
    elsif ($model->{table1513}
        && $self->{detailedVolumes}
        && $self->{scenarioProportions} )
    {
        Columnset(
            name     => 'Forecast volumes',
            number   => 1513,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            columns  => [
                @{ $self->{detailedVolumes} },
                @{ $self->{scenarioProportions} },
            ],
        );
    }
    else {
        Columnset(
            name     => 'Forecast volumes',
            number   => 1512,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            columns  => $self->{detailedVolumes},
        ) if $self->{detailedVolumes};
        Columnset(
            name     => 'Definition of user sets',
            number   => 1514,
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
            columns  => $self->{scenarioProportions},
        ) if $self->{scenarioProportions};
    }
}

1;
