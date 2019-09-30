package Elec::Interpolator;

# Copyright 2019 Franck Latrémolière, Reckon LLP and others.
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

sub firstDay {
    my ($self) = @_;
    $self->{firstDay} ||= Dataset(
        name          => 'First day of charging year',
        defaultFormat => 'datehard',
        data          => ['=DATE(2019,04,01)'],
    );
}

sub lastDay {
    my ($self) = @_;
    $self->{lastDay} ||= Dataset(
        name          => 'Last day of charging year',
        defaultFormat => 'datehard',
        data          => ['=DATE(2020,03,31)'],
    );
}

sub daysInYear {
    my ($self) = @_;
    $self->{daysInYear} ||= Arithmetic(
        name          => 'Number of days in the charging year',
        defaultFormat => '0soft',
        arithmetic    => '=A1-A2+1',
        arguments     => { A1 => $self->lastDay, A2 => $self->firstDay, },
    );
}

sub rowsDemand {
    my ($self) = @_;
    $self->{rowsDemand} ||= Labelset(
        list => [ 1 .. ( $self->{model}{numRowsDemand} || 24 ) ],
        defaultFormat => 'thitem'
    );
}

sub detailedVolumeData {

    my ($self) = @_;
    return $self->{detailedVolumeData} if $self->{detailedVolumeData};

    my $name = Dataset(
        name          => 'Name',
        rows          => $self->rowsDemand,
        defaultFormat => 'texthard',
        data          => [ map { ''; } @{ $self->rowsDemand->{list} } ],
    );
    my $tariff = Dataset(
        name          => 'Applicable tariff',
        rows          => $self->rowsDemand,
        defaultFormat => 'texthard',
        data          => [ map { ''; } @{ $self->rowsDemand->{list} } ],
    );
    my $startDate = Dataset(
        name          => 'Start date',
        rows          => $self->rowsDemand,
        defaultFormat => 'datehard',
        data => [ map { '=DATE(2019,4,1)'; } @{ $self->rowsDemand->{list} } ],
    );
    my $endDate = Dataset(
        name          => 'End date',
        rows          => $self->rowsDemand,
        defaultFormat => 'datehard',
        data          => [ map { ''; } @{ $self->rowsDemand->{list} } ],
    );
    my $growth = Dataset(
        name          => 'Annual growth rate',
        rows          => $self->rowsDemand,
        defaultFormat => '%hard',
        data          => [ map { 0; } @{ $self->rowsDemand->{list} } ],
    );
    my @volumes =
      map {
        Dataset(
            rows          => $self->rowsDemand,
            defaultFormat => '0hard',
            name          => $_,
            data          => [ map { ''; } @{ $self->rowsDemand->{list} } ],
            validation => {    # required to trigger leniency in cell locking
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
        );
      } @{ $self->{setup}->volumeComponents };

    $self->{detailedVolumeInputs} =
      [ $name, $tariff, $startDate, $endDate, $growth, @volumes ];

    my $first = Arithmetic(
        name          => 'Offset of first day',
        defaultFormat => '0soft',
        rows          => $self->rowsDemand,
        arithmetic    => '=IF(A201,MAX(0,A802-A202),0)',
        arguments     => {
            A201 => $startDate,
            A202 => $startDate,
            A802 => $self->firstDay,
        },
    );

    my $last = Arithmetic(
        name          => 'Offset of last day',
        defaultFormat => '0soft',
        rows          => $self->rowsDemand,
        arithmetic    => '=IF(A201,IF(A302,MIN(A901,A301),A902)-A202,-1)',
        arguments     => {
            A201 => $startDate,
            A202 => $startDate,
            A301 => $endDate,
            A302 => $endDate,
            A901 => $self->lastDay,
            A902 => $self->lastDay,
        },
    );

    my $factor = Arithmetic(
        name       => 'Scaling factor for the charging year',
        rows       => $self->rowsDemand,
        arithmetic => '=IF(OR(A103<0,A104<A203),0,IF(A701,'
          . '((1+A702)^((A102+1)/365.25)-(1+A703)^(A202/365.25))/LN(1+A704)'
          . ',(A101+1-A201)/365.25))',
        arguments => {
            A101 => $last,
            A102 => $last,
            A103 => $last,
            A104 => $last,
            A201 => $first,
            A202 => $first,
            A203 => $first,
            A701 => $growth,
            A702 => $growth,
            A703 => $growth,
            A704 => $growth,
        },
    );

    Columnset(
        name    => 'Volume forecasting calculations',
        columns => [ $first, $last, $factor, ]
    );

    $self->{detailedVolumeData} = [ $tariff, $factor, @volumes, ];

}

sub totalDemand {

    my ( $self, $usetName, $tariffSet ) = @_;
    return $self->{totalDemand}{$tariffSet}{$usetName}
      if $self->{totalDemand}{$tariffSet}{$usetName};

    my $lastRow = $#{ $self->rowsDemand->{list} };
    my $prop;
    push @{ $self->{scenarioProportions} }, $prop = Dataset(
        name          => "Proportion in $usetName",
        rows          => $self->rowsDemand,
        defaultFormat => '%hard',
        data          => [ map { 1; } @{ $self->rowsDemand->{list} } ],
        validation => {    # required to trigger leniency in cell locking
            validate      => 'decimal',
            criteria      => 'between',
            minimum       => -1,
            maximum       => 1,
            input_message => 'Percentage',
        },
    ) unless $usetName eq 'all users';
    my ( $tariff, $factor, @volumes ) = @{ $self->detailedVolumeData };

    my @columns =
      map {
        new SpreadsheetModel::Custom(
            name          => $_->objectShortName,
            defaultFormat => '0soft',
            rows          => $tariffSet,
            custom        => [
                    '=SUMPRODUCT((A1=A3)*A4*A6'
                  . ( $prop                   ? '*A8' : '' ) . ')'
                  . ( $_->{name} =~ /Wh|VArh/ ? ''    : '*365.25/A9' )
            ],
            arithmetic => '=SUMPRODUCT((A1=tariff)*A4*A6'
              . ( $prop                   ? '*A8' : '' ) . ')'
              . ( $_->{name} =~ /Wh|VArh/ ? ''    : '*365.25/A9' ),
            arguments => {
                A1 => $tariff,
                A4 => $factor,
                A6 => $_,
                $prop ? ( A8 => $prop ) : (),
                A9 => $self->{setup}->daysInYear,
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                my @map = (
                    qr/\bA9\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{A9}, $colh->{A9}, 1, 1
                      )
                );
                foreach (qw(A1 A4 A6 A8)) {
                    push @map,
                      qr/\b$_\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_}, $colh->{$_}, 1, 0 )
                      . ':'
                      . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + $lastRow,
                        $colh->{$_}, 1, 0 )
                      if exists $rowh->{$_};
                }
                sub {
                    my ( $x, $y ) = @_;
                    '', $format, $formula->[0], @map,
                      '\bA3\b' =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $self->{$wb}{row} + $y,
                        0, 0, 1 );
                };
            }
        );
      } @volumes;

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
    $self->{totalDemand}{$tariffSet}{$usetName} = \@columns;

}

sub finish {
    my ($self) = @_;
    return if $self->{firstLastDay};
    $self->{firstLastDay} = Columnset(
        name     => 'Period covered by this model',
        number   => 1501,
        appendTo => $self->{model}{inputTables},
        dataset  => $self->{model}{dataset},
        columns  => [ $self->firstDay, $self->lastDay, ],
    );
    Columnset(
        name     => 'Volume forecasting information',
        number   => 1515,
        appendTo => $self->{model}{inputTables},
        dataset  => $self->{model}{dataset},
        columns  => [
            @{ $self->{detailedVolumeInputs} },
            $self->{scenarioProportions}
            ? @{ $self->{scenarioProportions} }
            : (),
        ],
    ) if $self->{detailedVolumeInputs};
}

1;
