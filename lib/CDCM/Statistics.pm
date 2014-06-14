package CDCM;

=head Copyright licence and disclaimer

Copyright 2014 Franck Latrémolière, Reckon LLP and others.

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

sub makeStatisticsAssumptions {

    my ($model) = @_;

    return @{ $model->{sharedData}{statisticsAssumptionsStructure} }
      if $model->{sharedData}
      && $model->{sharedData}{statisticsAssumptionsStructure};

    my $colspec;
    $colspec = $model->{statistics} if ref $model->{statistics} eq 'ARRAY';
    $colspec = $model->{dataset}{1202}
      if !$colspec
      && $model->{summary} =~ /1202/
      && $model->{dataset}
      && $model->{dataset}[1202];
    unless ($colspec) {
        require YAML;
        ($colspec) = YAML::Load(<<'EOY');
---
1202:
  - Customer 11 Low use: 1900
    Customer 12 Medium use: 3800
    Customer 15 High use: 7600
    _column: Total consumption (kWh/year)
  - Customer 31 Small continuous: 69
    Customer 33 Small off-peak: 69
    Customer 33 Small peak-time: 69
    Customer 51 Continuous: 500
    Customer 81 Large continuous: 5000
    Customer 83 Large off-peak: 5000
    Customer 85 Large peak-time: 5000
    _column: Capacity (kVA)
  - Customer 31 Small continuous: 65
    Customer 33 Small off-peak: 65
    Customer 33 Small peak-time: 65
    Customer 51 Continuous: 450
    Customer 81 Large continuous: 4500
    Customer 83 Large off-peak: 4500
    Customer 85 Large peak-time: 4500
    _column: Consumption (kW)
  - Customer 31 Small continuous: 168
    Customer 33 Small off-peak: 77
    Customer 33 Small peak-time: 0
    Customer 51 Continuous: 168
    Customer 81 Large continuous: 168
    Customer 83 Large off-peak: 77
    Customer 85 Large peak-time: 0
    _column: Off-peak hours/week
  - Customer 31 Small continuous: 0
    Customer 33 Small off-peak: 0
    Customer 33 Small peak-time: 45
    Customer 51 Continuous: 0
    Customer 81 Large continuous: 0
    Customer 83 Large off-peak: 0
    Customer 85 Large peak-time: 45
    _column: Peak-time hours/week
  - ~
  - Customer 11 Low use: '/(?:^|: )Domestic Unrestricted/'
    Customer 12 Medium use: '/(?:^|: )Domestic Unrestricted/'
    Customer 15 High use: /Unrestricted/
    Customer 31 Small continuous: /^Small Non Domestic Unrestricted|^LV HH|^LV Network Non/
    Customer 33 Small off-peak: /^Small Non Domestic Unrestricted|^LV HH|^LV Network Non/
    Customer 33 Small peak-time: /^Small Non Domestic Unrestricted|^LV HH|^LV Network Non/
    Customer 51 Continuous: /(?:LV|LV Sub|HV) HH/
    Customer 81 Large continuous: '/^HV HH|LDNO HV: (?:LV|LV Sub) HH/'
    Customer 83 Large off-peak: '/^HV HH|LDNO HV: (?:LV|LV Sub) HH/'
    Customer 85 Large peak-time: '/^HV HH|LDNO HV: (?:LV|LV Sub) HH/'
EOY

=head Omitted for now

  - Customer 15 High use: 1900
    Customer 11 Low use: 475
    Customer 12 Medium use: 950
    _column: Rate 2 consumption (kWh/year)

=cut

        $colspec = $colspec->{1202};
    }

    my %capabilities;
    foreach (@$colspec) {
        next unless ref $_ eq 'HASH';
        if ( my $colName = $_->{_column} ) {
            push @{ $capabilities{$_} }, $colName
              foreach grep { $_ ne '_column' } keys %$_;
        }
        else {
            while ( my ( $k, $v ) = each $_ ) {
                push @{ $capabilities{$k} }, $v;
            }
        }
    }

    my @rows = sort keys %capabilities;

    my $rowset = Labelset( list => \@rows );

    my @columns = map {
        my $col = $_;
        Dataset(
            name          => $col->{_column},
            defaultFormat => '0hardnz',
            rows          => $rowset,
            data          => [ map { $col->{$_} } @rows ],
          )
    } grep { $_->{_column} } @$colspec;

    Columnset(
        name     => 'Assumed usage for illustrative customers',
        number   => 1202,
        appendTo => $model->{sharedData}
        ? $model->{sharedData}{statsAssumptions}
        : $model->{inputTables},
        dataset => $model->{dataset},
        columns => \@columns,
    );

    $model->{sharedData}{statisticsAssumptionsStructure} =
      [ \@rows, \%capabilities, @columns ]
      if $model->{sharedData};

    \@rows, \%capabilities, @columns;

}

sub makeStatisticsTables {

    my ( $model, $tariffTable, $daysInYear, $nonExcludedComponents,
        $componentMap, $unitsInYear )
      = @_;

    my $allTariffs ||= ( values %$tariffTable )[0]{rows};

    my ( $users, $capabilities, @columns ) = $model->makeStatisticsAssumptions;

    my ( %map, @groups );

    for ( my $uid = 0 ; $uid < @$users ; ++$uid ) {
        my $short = my $user = $users->[$uid];
        $short =~ s/^Customer *[0-9]+ *//;
        my $caps = $capabilities->{$user};
        my @list;
      TARIFF: for ( my $tid = 0 ; $tid < @{ $allTariffs->{list} } ; ++$tid ) {
            next
              if $allTariffs->{groupid}
              && !defined $allTariffs->{groupid}[$tid];
            my $tariff = $allTariffs->{list}[$tid];
            foreach (@$caps) {
                if (m#^/(.+)/$#s) {
                    next TARIFF unless $tariff =~ /$1/m;
                    next;
                }
                next TARIFF
                  if m#kWh/year#
                  && $componentMap->{$tariff}{'Unit rates p/kWh'};
            }
            $tariff =~ s/^.*\n//s;
            my $row = "$short ($tariff)";
            push @list, $row;
            $map{$row} = [ $uid, $tid ];
        }
        push @groups, Labelset( name => $user, list => \@list );
    }

    my ($annualUnits) =
      grep { $_->{name} =~ m#kWh/year# && $_->{name} =~ /total/i } @columns;
    my ($kW) = grep { $_->{name} =~ m#kW# && $_->{name} !~ /kWh/i } @columns;
    my ($offhours)  = grep { $_->{name} =~ /off[- ]peak hours/i } @columns;
    my ($peakhours) = grep { $_->{name} =~ /peak[- ]time hours/i } @columns;
    my ($capacity) =
      grep { $_->{name} =~ m#kVA# && $_->{name} !~ /kVArh/i } @columns;

    my $units = Arithmetic(
        name => Label(
            'MWh/year',
            'Annual consumption of illustrative customers (MWh/year)'
        ),
        rows       => Labelset( list => \@groups ),
        arithmetic => '=0.001*(IV1+IV2*(IV3+IV4)/7*IV5)',
        arguments  => {
            IV1 => $annualUnits,
            IV2 => $kW,
            IV3 => $offhours,
            IV4 => $peakhours,
            IV5 => $daysInYear,
        },
    );

    my $fullRowset = Labelset( groups => \@groups );
    my @map = @map{ @{ $fullRowset->{list} } };

    my $charge = SpreadsheetModel::Custom->new(
        name => Label(
            'Charge £/year',
            'Annual charges for illustrative customers (£/year)'
        ),
        defaultFormat => '0softnz',
        rows          => $fullRowset,
        custom        => [
            '=10*IV1*IV91+0.01*IV71*IV94',
            '=0.01*(IV3*('
              . 'MIN(IV61,IV72/7*IV51+MAX(0,IV77/7*IV43-IV631-IV623))*IV91'
              . '+MIN(IV62,MAX(0,IV73/7*IV52-IV611)+MAX(0,IV75/7*IV42-IV632))*IV92'
              . '+MIN(IV63,MAX(0,IV74/7*IV53-IV612-IV622)+IV76/7*IV41)*IV93'
              . ')+IV71*(IV94+IV2*IV95))'
        ],
        arguments => {
            IV1   => $units,
            IV2   => $capacity,
            IV3   => $kW,
            IV41  => $offhours,
            IV42  => $offhours,
            IV43  => $offhours,
            IV51  => $peakhours,
            IV52  => $peakhours,
            IV53  => $peakhours,
            IV61  => $model->{hoursByRedAmberGreen},
            IV611 => $model->{hoursByRedAmberGreen},
            IV612 => $model->{hoursByRedAmberGreen},
            IV62  => $model->{hoursByRedAmberGreen},
            IV622 => $model->{hoursByRedAmberGreen},
            IV623 => $model->{hoursByRedAmberGreen},
            IV63  => $model->{hoursByRedAmberGreen},
            IV631 => $model->{hoursByRedAmberGreen},
            IV632 => $model->{hoursByRedAmberGreen},
            IV71  => $daysInYear,
            IV72  => $daysInYear,
            IV73  => $daysInYear,
            IV74  => $daysInYear,
            IV75  => $daysInYear,
            IV76  => $daysInYear,
            IV77  => $daysInYear,
            IV91  => $tariffTable->{'Unit rate 1 p/kWh'},
            IV92  => $tariffTable->{'Unit rate 2 p/kWh'},
            IV93  => $tariffTable->{'Unit rate 3 p/kWh'},
            IV94  => $tariffTable->{'Fixed charge p/MPAN/day'},
            IV95  => $tariffTable->{'Capacity charge p/kVA/day'},
        },
        rowFormats => [ map { $_ ? undef : 'unavailable'; } @map ],
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                my $cellFormat =
                    $self->{rowFormats}[$y]
                  ? $wb->getFormat( $self->{rowFormats}[$y] )
                  : $format;
                return '', $cellFormat unless $map[$y];
                my ( $uid, $tid ) = @{ $map[$y] };
                my $tariff = $allTariffs->{list}[$tid];
                '', $cellFormat,
                  $formula->[ $tariff !~ /gener/i
                  && !$componentMap->{$tariff}{'Unit rates p/kWh'} ? 0 : 1 ],
                  map {
                    qr/\b$_\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + (
                              /^IV9/         ? $tid
                            : /^IV(?:[2-5])/ ? $uid
                            : /^IV1/         ? $fullRowset->{groupid}[$y]
                            : 0
                        ),
                        $colh->{$_} + ( /^IV62/ ? 1 : /^IV63/ ? 2 : 0 ),
                        1, 1,
                      )
                  } @$pha;
            };
        },
    );

    my $stats = Arithmetic(
        name          => '£/MWh for illustrative customers',
        defaultFormat => '0.00soft',
        arithmetic    => '=IV1/IV2',
        arguments     => { IV1 => $charge, IV2 => $units, }
    );

    push @{ $model->{statisticsTables} }, $stats;

    if ( $model->{sharedData} ) {
        $model->{sharedData}
          ->addStats( 'Annual charges for illustrative customers (£/year)',
            $model, $charge );
        $model->{sharedData}
          ->addStats( 'Distribution costs for illustrative customers (£/MWh)',
            $model, $stats );
    }

}

1;
