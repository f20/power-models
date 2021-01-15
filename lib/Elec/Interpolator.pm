package Elec::Interpolator;

# Copyright 2019-2021 Franck Latrémolière and others.
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
require Spreadsheet::WriteExcel::Utility;

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

sub chargingPeriodLabel {
    my ( $self, $wbook, $wsheet ) = @_;
    my ( $sh, $ro, $co ) = $self->firstDay->wsWrite( $wbook, $wsheet );
    $sh = $sh->get_name;
    my $fd = "'$sh'!"
      . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co );
    ( $sh, $ro, $co ) = $self->lastDay->wsWrite( $wbook, $wsheet );
    $sh = $sh->get_name;
    my $ld = "'$sh'!"
      . Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell( $ro, $co );
    qq%YEAR($fd)&"/"&YEAR($ld)%;
}

sub daysInYear {
    my ($self) = @_;
    return $self->{daysInYear} if $self->{daysInYear};
    push @{ $self->{model}{volumeTables} },
      $self->{daysInYear} = Arithmetic(
        name          => 'Number of days in the charging year',
        defaultFormat => '0soft',
        arithmetic    => '=A1-A2+1',
        arguments     => { A1 => $self->lastDay, A2 => $self->firstDay, },
      );
    $self->{daysInYear};
}

sub _forecastInputDataAndFactors {

    my ( $self, $rowset, $tableName ) = @_;

    my $startDate = Dataset(
        name          => 'Start date',
        rows          => $rowset,
        defaultFormat => 'datehard',
        data          => [ map { '=DATE(2019,4,1)'; } @{ $rowset->{list} } ],
    );
    my $endDate = Dataset(
        name          => 'End date',
        rows          => $rowset,
        defaultFormat => 'datehard',
        data          => [ map { ''; } @{ $rowset->{list} } ],
    );
    my $growth = Dataset(
        name          => 'Annual growth rate',
        rows          => $rowset,
        defaultFormat => '%hard',
        data          => [ map { 0; } @{ $rowset->{list} } ],
    );

    my $first = Arithmetic(
        name          => 'Offset of first day',
        defaultFormat => '0soft',
        rows          => $rowset,
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
        rows          => $rowset,
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
        name => 'Scaling factor for the charging year',
        rows => $rowset,
        arithmetic => '=IF(OR(A103<0,A104<A203),0,' .    # is there any overlap
          'IF(A701,'    # is there a non-zero growth rate?
          . '((1+A702)^((A102+1-A204)/A901)-1)*(1+A703)^(A202/365.25)/LN(1+A704),'
          . '(A101+1-A201)/A902))',
        arguments => {
            A101 => $last,
            A102 => $last,
            A103 => $last,
            A104 => $last,
            A201 => $first,
            A202 => $first,
            A203 => $first,
            A204 => $first,
            A701 => $growth,
            A702 => $growth,
            A703 => $growth,
            A704 => $growth,
            A901 => $self->daysInYear,
            A902 => $self->daysInYear,
        },
    );

    push @{ $self->{columnsets} },
      Columnset(
        name    => "$tableName interpolation and extrapolation calculations",
        columns => [ $first, $last, $factor, ]
      );

    $startDate, $endDate, $growth, $factor;

}

sub columnsets {
    my ($self) = @_;
    @{ $self->{columnsets} };
}

sub assetValuesLives {

    my ( $self, $rowset ) = @_;

    my $life = Dataset(
        name => 'Annualisation period (years)',
        rows => $rowset,
        data => [ map { ''; } @{ $rowset->{list} } ],
    );
    my $baseValue = Dataset(
        name          => 'Baseline value (£)',
        rows          => $rowset,
        defaultFormat => '0hard',
        data          => [ map { ''; } @{ $rowset->{list} } ],
    );
    my $startDate = Dataset(
        name          => 'Reference date',
        rows          => $rowset,
        defaultFormat => 'datehard',
        data          => [ map { ''; } @{ $rowset->{list} } ],
    );
    my $growth = Dataset(
        name          => 'Annual change in value',
        rows          => $rowset,
        defaultFormat => '%hard',
        data          => [ map { ''; } @{ $rowset->{list} } ],
    );

    Columnset(
        name     => 'Asset value forecasting information',
        number   => 1552,
        appendTo => $self->{model}{inputTables},
        dataset  => $self->{model}{dataset},
        columns  => [ $life, $baseValue, $startDate, $growth, ],
    );

    my $first = Arithmetic(
        name          => 'Offset of first day',
        defaultFormat => '0soft',
        rows          => $rowset,
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
        rows          => $rowset,
        arithmetic    => '=IF(A201,A901-A202,-1)',
        arguments     => {
            A201 => $startDate,
            A202 => $startDate,
            A901 => $self->lastDay,
        },
    );

    my $value = Arithmetic(
        name          => 'Asset valuation (£)',
        defaultFormat => '0soft',
        rows          => $rowset,
        arithmetic    => '=IF(OR(A103<0,A104<A203),0,IF(A701,'
          . '((1+A702)^((A102+1)/365.25)-(1+A703)^(A202/365.25))/LN(1+A704)'
          . ',(A101+1-A201)/365.25))*365.25/A9*A8',
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
            A8   => $baseValue,
            A9   => $self->daysInYear,
        },
    );

    Columnset(
        name    => 'Asset value forecasting calculations',
        columns => [ $first, $last, $value, ],
    );

    $value, $life;

}

sub _aggregateForecast {

    my ( $self, $wantedRowset, $wantedColumnset, $prop, $category, $factor,
        $column )
      = @_;

    my $name =
        $wantedColumnset && !$#{ $wantedColumnset->{list} }
      ? $wantedColumnset->{list}[0]
      : $column->objectShortName;

    new SpreadsheetModel::Custom(
        name          => $name,
        defaultFormat => '0soft',
        rows          => $wantedRowset,
        cols          => $wantedColumnset,
        custom =>
          [ '=SUMPRODUCT((A1=A3)*A4*A6' . ( $prop ? '*A8' : '' ) . ')' ],
        arithmetic => '=SUMPRODUCT((A1=label)*A4*A6'
          . ( $prop ? '*A8' : '' ) . ')',
        arguments => {
            A1 => $category,
            A4 => $factor,
            A6 => $column,
            $prop ? ( A8 => $prop ) : (),
            A9 => $self->{setup}->daysInYear,
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
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
                    $rowh->{$_} + $#{ $factor->{rows}{list} },
                    $colh->{$_}, 1, 0 )
                  if exists $rowh->{$_};
            }
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0], @map,
                  '\bA3\b' => defined $wantedRowset
                  ? Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $self->{$wb}{row} + $y,
                    0, 0, 1 )
                  : Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $self->{$wb}{row} - 1,
                    $self->{$wb}{col} + $x,
                    1, 0
                  );
            };
        }
    );

}

sub assetVolumes {
    my ( $self, $assetLabelset ) = @_;
    my $inputRowset = Labelset(
        list => [ 1 .. ( $self->{model}{numRowsAssetVolumes} || 15 ) ],
        defaultFormat => 'thitem'
    );
    my $name = Dataset(
        name          => 'Name',
        rows          => $inputRowset,
        defaultFormat => 'texthard',
        data          => [ map { ''; } @{ $inputRowset->{list} } ],
    );
    my $category = Dataset(
        name          => 'Asset category',
        rows          => $inputRowset,
        defaultFormat => 'texthard',
        data          => [ map { ''; } @{ $inputRowset->{list} } ],
    );
    my ( $startDate, $endDate, $growth, $factor ) =
      $self->_forecastInputDataAndFactors( $inputRowset, 'Asset volume' );
    my $quantity = Dataset(
        name => 'Asset count',
        rows => $inputRowset,
        data => [ map { ''; } @{ $inputRowset->{list} } ],
    );
    Columnset(
        name     => 'Asset volume forecasting information',
        number   => 1550,
        appendTo => $self->{model}{inputTables},
        dataset  => $self->{model}{dataset},
        columns =>
          [ $name, $category, $startDate, $endDate, $growth, $quantity, ],
    );
    $self->_aggregateForecast( $assetLabelset, undef, undef, $category,
        $factor, $quantity );
}

sub targetUsage {
    my ( $self, $usageSet ) = @_;
    my $inputRowset = Labelset(
        list => [ 1 .. ( $self->{model}{numRowsTargetUsage} || 5 ) ],
        defaultFormat => 'thitem'
    );
    my $name = Dataset(
        name          => 'Name',
        rows          => $inputRowset,
        defaultFormat => 'texthard',
        data          => [ map { ''; } @{ $inputRowset->{list} } ],
    );
    my $level = Dataset(
        name          => 'Usage level',
        rows          => $inputRowset,
        defaultFormat => 'texthard',
        data          => [ map { ''; } @{ $inputRowset->{list} } ],
    );
    my ( $startDate, $endDate, $growth, $factor ) =
      $self->_forecastInputDataAndFactors( $inputRowset, 'Target usage' );
    my $capacity = Dataset(
        name          => 'Network capacity (kVA)',
        rows          => $inputRowset,
        defaultFormat => '0hard',
        data          => [ map { ''; } @{ $inputRowset->{list} } ],
    );
    Columnset(
        name     => 'Target usage forecasting information',
        number   => 1538,
        appendTo => $self->{model}{inputTables},
        dataset  => $self->{model}{dataset},
        columns => [ $name, $level, $startDate, $endDate, $growth, $capacity, ],
    );
    $self->_aggregateForecast( undef, $usageSet, undef, $level,
        $factor, $capacity );
}

sub runningCostData {
    my ($self) = @_;
    return $self->{runningCostData} if $self->{runningCostData};
    my $inputRowset = Labelset(
        list => [ 1 .. ( $self->{model}{numRowsRunningCosts} || 10 ) ],
        defaultFormat => 'thitem'
    );
    my $name = Dataset(
        name          => 'Name',
        rows          => $inputRowset,
        defaultFormat => 'texthard',
        data          => [ map { ''; } @{ $inputRowset->{list} } ],
    );
    my $category = Dataset(
        name          => 'Cost category',
        rows          => $inputRowset,
        defaultFormat => 'texthard',
        data          => [ map { ''; } @{ $inputRowset->{list} } ],
    );
    my ( $startDate, $endDate, $growth, $factor ) =
      $self->_forecastInputDataAndFactors( $inputRowset, 'Running cost' );
    my $amount = Dataset(
        name          => 'Annual cost (£/year)',
        rows          => $inputRowset,
        defaultFormat => '0hard',
        data          => [ map { ''; } @{ $inputRowset->{list} } ],
    );
    Columnset(
        name     => 'Running cost forecasting information',
        number   => 1558,
        appendTo => $self->{model}{inputTables},
        dataset  => $self->{model}{dataset},
        columns =>
          [ $name, $category, $startDate, $endDate, $growth, $amount, ],
    );
    $self->{runningCostData} = [ $category, $factor, $amount ];
}

sub runningCosts {
    my ( $self, $usageSet ) = @_;
    $self->_aggregateForecast( undef, $usageSet, undef,
        @{ $self->runningCostData } );
}

sub demandInputAndFactor {
    my ($self) = @_;
    return $self->{demandInputAndFactor} if $self->{demandInputAndFactor};
    my $inputRowset = Labelset(
        list => [ 1 .. ( $self->{model}{numRowsDemand} || 24 ) ],
        defaultFormat => 'thitem'
    );
    my $name = Dataset(
        name          => 'Name',
        rows          => $inputRowset,
        defaultFormat => 'texthard',
        data          => [ map { ''; } @{ $inputRowset->{list} } ],
    );
    my $tariff = Dataset(
        name          => 'Applicable tariff',
        rows          => $inputRowset,
        defaultFormat => 'texthard',
        data          => [ map { ''; } @{ $inputRowset->{list} } ],
    );
    my ( $startDate, $endDate, $growth, $factor ) =
      $self->_forecastInputDataAndFactors( $inputRowset, 'Demand' );
    my @demands = map {
        Dataset(
            rows          => $inputRowset,
            defaultFormat => '0hard',
            name          => $_,
            data          => [ map { ''; } @{ $inputRowset->{list} } ],
        );
    } @{ $self->{setup}->volumeComponents };
    $self->{demandInputColumns} =
      [ $name, $tariff, $startDate, $endDate, $growth, @demands, ];
    $self->{demandInputAndFactor} = [ $tariff, $factor, @demands, ];
}

sub totalDemand {
    my ( $self, $usetName, $wantedRowset ) = @_;
    return $self->{totalDemand}{$wantedRowset}{$usetName}
      if $self->{totalDemand}{$wantedRowset}{$usetName};
    my ( $tariff, $factor, @demands ) = @{ $self->demandInputAndFactor };
    my $prop;
    push @{ $self->{scenarioProportions} },
      $prop = Dataset(
        name          => "Proportion in $usetName",
        rows          => $factor->{rows},
        defaultFormat => '%hard',
        data          => [ map { 1; } @{ $factor->{rows}{list} } ],
      ) unless $usetName eq 'all users';
    my @columns = map {
        $self->_aggregateForecast( $wantedRowset, undef, $prop, $tariff,
            $factor, $_ );
    } @demands;
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
        name    => "Forecast demands for $usetName",
        columns => \@columns,
      );
    $self->{totalDemand}{$wantedRowset}{$usetName} = \@columns;
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
            @{ $self->{demandInputColumns} },
            $self->{scenarioProportions}
            ? @{ $self->{scenarioProportions} }
            : (),
        ],
    ) if $self->{demandInputColumns};
}

1;
