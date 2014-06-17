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
    $colspec = $model->{dataset}{$1}
      if !$colspec
      && $model->{summary} =~ /([0-9]{3,4})/
      && $model->{dataset}
      && $model->{dataset}{$1};
    unless ($colspec) {
        require YAML;
        ($colspec) = YAML::Load(<<'EOY');
---
1234:
  - Customer 11 Low use: 1900
    Customer 12 Medium use: 3800
    Customer 15 High use: 7600
    _column: Total consumption (kWh/year)
  - Customer 15 High use: 5000
    Customer 11 Low use: 475
    Customer 12 Medium use: 950
    _column: Rate 2 consumption (kWh/year)
  - Customer 31 Small continuous: 69
    Customer 33 Small off-peak: 69
    Customer 33 Small peak-time: 69
    Customer 51 Continuous: 500
    Customer 53 Off-peak: 500
    Customer 55 Peak-time: 500
    Customer 81 Large continuous: 5000
    Customer 83 Large off-peak: 5000
    Customer 85 Large peak-time: 5000
    _column: Capacity (kVA)
  - Customer 31 Small continuous: 65
    Customer 33 Small off-peak: 65
    Customer 33 Small peak-time: 65
    Customer 51 Continuous: 450
    Customer 53 Off-peak: 450
    Customer 55 Peak-time: 450
    Customer 81 Large continuous: 4500
    Customer 83 Large off-peak: 4500
    Customer 85 Large peak-time: 4500
    _column: Consumption (kW)
  - Customer 31 Small continuous: 168
    Customer 33 Small off-peak: 77
    Customer 33 Small peak-time: 0
    Customer 51 Continuous: 168
    Customer 53 Off-peak: 77
    Customer 55 Peak-time: 0
    Customer 81 Large continuous: 168
    Customer 83 Large off-peak: 77
    Customer 85 Large peak-time: 0
    _column: Off-peak hours/week
  - Customer 31 Small continuous: 0
    Customer 33 Small off-peak: 0
    Customer 33 Small peak-time: 48
    Customer 51 Continuous: 0
    Customer 53 Off-peak: 0
    Customer 55 Peak-time: 48
    Customer 81 Large continuous: 0
    Customer 83 Large off-peak: 0
    Customer 85 Large peak-time: 48
    _column: Peak-time hours/week
  - ~
  - Customer 11 Low use: '/^(?:|LDNO.*: )Domestic Unrestricted/'
    Customer 12 Medium use: '/^(?:|LDNO.*: )Domestic Unrestricted/'
    Customer 15 High use: '/^(?:(Small Non )?Domestic (?:Unrestricted|Two)|LV.*Medium)/'
    Customer 31 Small continuous: /^Small Non Domestic Unrestricted|^LV HH|^LV Network Non/
    Customer 33 Small off-peak: /^Small Non Domestic Unrestricted|^LV HH|^LV Network Non/
    Customer 33 Small peak-time: /^Small Non Domestic Unrestricted|^LV HH|^LV Network Non/
    Customer 51 Continuous: '/^(?:LV|LV Sub|HV|LDNO HV: (?:LV|LV Sub)) HH/'
    Customer 53 Off-peak: '/^(?:LV|LV Sub|HV|LDNO HV: (?:LV|LV Sub)) HH/'
    Customer 55 Peak-time: '/^(?:LV|LV Sub|HV|LDNO HV: (?:LV|LV Sub)) HH/'
    Customer 81 Large continuous: /^HV HH/
    Customer 83 Large off-peak: /^HV HH/
    Customer 85 Large peak-time: /^HV HH/
EOY

        ($colspec) = values %$colspec;

    }

    my %capabilities;
    foreach (@$colspec) {
        next unless ref $_ eq 'HASH';
        if ( my $colName = $_->{_column} ) {
            push @{ $capabilities{$_} }, $colName
              foreach grep { $_ ne '_column' } keys %$_;
        }
        else {
            while ( my ( $k, $v ) = each %$_ ) {
                push @{ $capabilities{$k} }, $v;
            }
        }
    }

    my @rows = sort keys %capabilities;

    my $rowset = Labelset( list => \@rows );

    my @columns = map {
        my $col = $_;
        Constant( # Editable constant: auto-populates irrespective of data from dataset
            name          => $col->{_column},
            defaultFormat => '0hardnz',
            rows          => $rowset,
            data          => [ map { $col->{$_} } @rows ],
          )
    } grep { $_->{_column} } @$colspec;

    Columnset(
        name => 'Assumed usage for illustrative customers',
        $model->{sharedData}
        ? ( appendTo => $model->{sharedData}{statsAssumptions} )
        : (),
        columns => \@columns,
    );

    $model->{sharedData}{statisticsAssumptionsStructure} =
      [ \@rows, \%capabilities, @columns ]
      if $model->{sharedData};

    \@rows, \%capabilities, @columns;

}

sub makeStatisticsTables {

    my ( $model, $tariffTable, $daysInYear, $nonExcludedComponents,
        $componentMap, )
      = @_;

    my ($allTariffs) = values %$tariffTable;
    $allTariffs = $allTariffs->{rows};

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
            $map{$row} = [ $uid, $tid, $#list ];
            if ( $tariff =~ /^LDNO ([^:]+): (.+)/ ) {
                my $boundary = $1;
                my $atw      = $2;
                if ( my $atwmapped = $map{"$short ($atw)"} ) {
                    my $marginrow = "$short (Margin $boundary: $atw)";
                    push @list, $marginrow;
                    $map{$marginrow} =
                      [ undef, $atwmapped->[2] - $#list, $#list ];
                }
            }
        }
        push @groups, Labelset( name => $user, list => \@list );
    }

    my ($annualUnits) =
      grep { $_->{name} =~ m#kWh/year#i && $_->{name} =~ /total/i } @columns;
    my ($rate2Units) =
      grep { $_->{name} =~ m#kWh/year#i && $_->{name} =~ /rate 2/i } @columns;
    my ($kW) = grep { $_->{name} =~ /kW/i && $_->{name} !~ /kWh/i } @columns;
    my ($offhours)  = grep { $_->{name} =~ /off[- ]peak hours/i } @columns;
    my ($peakhours) = grep { $_->{name} =~ /peak[- ]time hours/i } @columns;
    my ($capacity) =
      grep { $_->{name} =~ /kVA/i && $_->{name} !~ /kVArh/i } @columns;

    my $fullRowset = Labelset( groups => \@groups );
    my @map = @map{ @{ $fullRowset->{list} } };

    my $charge = SpreadsheetModel::Custom->new(
        name => Label(
            '£/year', 'Annual charges for illustrative customers (£/year)',
        ),
        defaultFormat => '0softnz',
        rows          => $fullRowset,
        custom        => [
            '=0.01*((IV11+IV12*(IV13+IV14)/7*IV78)*IV91+IV71*IV94)',
            '=0.01*('
              . '(IV11+IV12*(IV13+IV14)/7*IV78-IV17)*IV91'
              . '+IV18*IV92+IV71*IV94)',
            '=0.01*(IV3*('
              . 'MIN(IV61,IV72/7*IV51+MAX(0,IV77/7*IV43-IV631-IV623))*IV91'
              . '+MIN(IV62,MAX(0,IV73/7*IV52-IV611)+MAX(0,IV75/7*IV42-IV632))*IV92'
              . '+MIN(IV63,MAX(0,IV74/7*IV53-IV612-IV622)+IV76/7*IV41)*IV93'
              . ')+IV71*(IV94+IV2*IV95))',
            '=IV81-IV82',
        ],
        arguments => {
            IV11  => $annualUnits,
            IV12  => $kW,
            IV13  => $offhours,
            IV14  => $peakhours,
            IV17  => $rate2Units,
            IV18  => $rate2Units,
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
            IV78  => $daysInYear,
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
                unless ( defined $uid ) {
                    return '', $cellFormat, $formula->[3],
                      qr/\bIV81\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $self->{$wb}{row} + $y + $tid,
                        $self->{$wb}{col}
                      ),
                      qr/\bIV82\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $self->{$wb}{row} + $y - 1,
                        $self->{$wb}{col} );
                }
                my $tariff = $allTariffs->{list}[$tid];
                '', $cellFormat,
                  $formula->[
                    $componentMap->{$tariff}{'Unit rates p/kWh'}  ? 2
                  : $componentMap->{$tariff}{'Unit rate 2 p/kWh'} ? 1
                  : 0
                  ],
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
        name => Label(
            '£/MWh', 'Average charges for illustrative customers (£/MWh)'
        ),
        defaultFormat => '0.00soft',
        arithmetic    => '=IV1/(IV11+IV12*(IV13+IV14)/7*IV78)*1000',
        arguments     => {
            IV1  => $charge,
            IV11 => $annualUnits,
            IV12 => $kW,
            IV13 => $offhours,
            IV14 => $peakhours,
            IV78 => $daysInYear,
        }
    );

    if ( $model->{sharedData} ) {
        $model->{sharedData}
          ->addStats( 'Annual charges for illustrative customers (£/year)',
            $model, $charge );
        $model->{sharedData}
          ->addStats( 'Distribution costs for illustrative customers (£/MWh)',
            $model, $stats );
    }

    push @{ $model->{statisticsTables} },
      Columnset(
        name    => 'Statistics for illustrative customers',
        columns => [ $charge, $stats ],
      );

}

1;
