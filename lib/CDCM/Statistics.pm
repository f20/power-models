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

use YAML;
local undef $/;
binmode DATA, ':utf8';
my ($colspecDefault) = Load <DATA>;
($colspecDefault) = values %$colspecDefault;

sub makeStatisticsAssumptions {

    my ($model) = @_;

    return $model->{sharedData}{statisticsAssumptionsStructure}
      if $model->{sharedData}
      && $model->{sharedData}{statisticsAssumptionsStructure};

    my $colspec;
    $colspec = $model->{statistics} if ref $model->{statistics} eq 'ARRAY';
    $colspec = $model->{dataset}{$1}
      if !$colspec
      && $model->{summary} =~ /([0-9]{3,4})/
      && $model->{dataset}
      && $model->{dataset}{$1};
    $colspec ||= $colspecDefault;

    my @rows = sort {
        my @n =
          map {
            my $n = 0;
            $n = $1        if /([0-9]+)\s*kVA/;
            $n = $1 * 1000 if /([0-9]+)\s*MVA/;
            $n;
          } $a, $b;
        $n[0] <=> $n[1] || $a cmp $b;
    } grep { $_ ne '_column'; } keys %{ $colspec->[1] };

    my $rowset = Labelset( list => \@rows );

  # Editable constant columns so that they auto-populate irrespective of dataset
    my @columns = map {
        my $col = $_;
        Constant(
            name          => $col->{_column},
            defaultFormat => $col->{_column} =~ /hours\/week/ ? '0.0hard'
            : $col->{_column} =~ /kVA/ ? '0hard'
            : '0.000hard',
            rows => $rowset,
            data => [ @$_{@rows} ],
          )
    } @$colspec[ 2 .. 7 ];

    my $result = Columnset(
        name => 'Consumption assumptions for illustrative customers',
        $model->{sharedData}
        ? (
            number   => 4001,
            appendTo => $model->{sharedData}{statsAssumptions},
          )
        : (),
        columns => \@columns,
        regex   => $colspec->[1],
    );

    $model->{sharedData}{statisticsAssumptionsStructure} = $result
      if $model->{sharedData};

    $result;

}

sub makeStatisticsTables {

    my ( $model, $tariffTable, $daysInYear, $nonExcludedComponents,
        $componentMap, )
      = @_;

    my ($allTariffs) = values %$tariffTable;
    $allTariffs = $allTariffs->{rows};

    my $assumptions = $model->makeStatisticsAssumptions;

    my ( %override, @columns2 );

    unless ( $model->{sharedData} ) {

        my $blank = [ map { '' } @{ $assumptions->{rows}{list} } ];
        push @columns2,
          $override{$_} = Constant(
            name          => "Override\t$_ kWh/year",
            defaultFormat => '0hard',
            rows          => $assumptions->{rows},
            data          => $blank,
          ) foreach qw(red amber green);

        push @columns2,
          $override{total} = Arithmetic(
            name          => "Total override kWh/year",
            defaultFormat => '0soft',
            arithmetic    => '=IV1+IV2+IV3',
            arguments     => {
                IV1 => $override{red},
                IV2 => $override{amber},
                IV3 => $override{green},
            },
          );

    }

    push @columns2,
      my $totalUnits = Arithmetic(
        name          => 'Total kWh/year',
        defaultFormat => '0soft',
        arithmetic    => %override
        ? '=IF(IV8,IV9,(IV1*IV3+IV2*IV4+(168-IV11-IV21)*IV5)*IV7/7)'
        : '=(IV1*IV3+IV2*IV4+(168-IV11-IV21)*IV5)*IV7/7',
        arguments => {
            IV7 => $daysInYear,
            %override ? ( IV8 => $override{total}, IV9 => $override{total} )
            : (),
            IV1  => $assumptions->{columns}[0],
            IV11 => $assumptions->{columns}[0],
            IV2  => $assumptions->{columns}[1],
            IV21 => $assumptions->{columns}[1],
            IV3  => $assumptions->{columns}[2],
            IV4  => $assumptions->{columns}[3],
            IV5  => $assumptions->{columns}[4],
        },
      );

    push @columns2,
      my $rate2 = Arithmetic(
        name          => 'Rate 2 kWh/year',
        defaultFormat => '0soft',
        arithmetic    => '=IV1*IV3*IV2/7',
        arguments     => {
            IV1 => $assumptions->{columns}[1],
            IV3 => $assumptions->{columns}[3],
            IV2 => $daysInYear,
        },
      );

    push @columns2, my $red = SpreadsheetModel::Custom->new(
        name          => 'Red kWh/year',
        defaultFormat => '0soft',
        rows          => $assumptions->{rows},
        custom        => [
            %override
            ? '=IF(IV10,IV11,IV321*IV613+(IV311-IV322)*MIN(IV61,IV72/7*IV51)+'
              . '(IV301-IV323)*MAX(0,IV77/7*IV43-IV631-IV623))'
            : '=IV321*IV613+(IV311-IV322)*MIN(IV61,IV72/7*IV51)+'
              . '(IV301-IV323)*MAX(0,IV77/7*IV43-IV631-IV623)'
        ],
        arguments => {
            %override
            ? (
                IV10 => $override{total},
                IV11 => $override{red}
              )
            : (),
            IV301 => $assumptions->{columns}[3],
            IV311 => $assumptions->{columns}[2],
            IV321 => $assumptions->{columns}[4],
            IV322 => $assumptions->{columns}[4],
            IV323 => $assumptions->{columns}[4],
            IV61  => $model->{hoursByRedAmberGreen},
            IV613 => $model->{hoursByRedAmberGreen},
            IV623 => $model->{hoursByRedAmberGreen},
            IV631 => $model->{hoursByRedAmberGreen},
            IV43  => $assumptions->{columns}[1],
            IV51  => $assumptions->{columns}[0],
            IV72  => $daysInYear,
            IV77  => $daysInYear,
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0], map {
                    qr/\b$_\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + (
                            /^IV[1-5]/
                            ? $y
                            : 0
                        ),
                        $colh->{$_} + ( /^IV62/ ? 1 : /^IV63/ ? 2 : 0 ),
                        /^IV[1-5]/ ? 0 : 1,
                        1,
                      )
                } @$pha;
            };
        },
    );

    push @columns2, my $amber = SpreadsheetModel::Custom->new(
        name          => 'Amber kWh/year',
        defaultFormat => '0soft',
        rows          => $assumptions->{rows},
        custom        => [
            %override
            ? '=IF(IV10,IV12,'
              . 'IV324*IV624+(IV312-IV325)*MIN(IV620,MAX(0,IV73/7*IV52-IV611))+'
              . '(IV302-IV326)*MIN(IV621,MAX(0,IV75/7*IV42-IV632)))'
            : '=IV324*IV624+(IV312-IV325)*MIN(IV620,MAX(0,IV73/7*IV52-IV611))+'
              . '(IV302-IV326)*MIN(IV621,MAX(0,IV75/7*IV42-IV632))'
        ],
        arguments => {
            %override
            ? (
                IV10 => $override{total},
                IV12 => $override{amber}
              )
            : (),
            IV302 => $assumptions->{columns}[3],
            IV312 => $assumptions->{columns}[2],
            IV324 => $assumptions->{columns}[4],
            IV325 => $assumptions->{columns}[4],
            IV326 => $assumptions->{columns}[4],
            IV42  => $assumptions->{columns}[1],
            IV52  => $assumptions->{columns}[0],
            IV611 => $model->{hoursByRedAmberGreen},
            IV620 => $model->{hoursByRedAmberGreen},
            IV621 => $model->{hoursByRedAmberGreen},
            IV624 => $model->{hoursByRedAmberGreen},
            IV632 => $model->{hoursByRedAmberGreen},
            IV73  => $daysInYear,
            IV75  => $daysInYear,
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0], map {
                    qr/\b$_\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + ( /^IV[1-5]/ ? $y : 0 ),
                        $colh->{$_} + ( /^IV62/ ? 1 : /^IV63/ ? 2 : 0 ),
                        /^IV[1-5]/ ? 0 : 1,
                        1,
                      )
                } @$pha;
            };
        },
    );

    push @columns2, my $green = SpreadsheetModel::Custom->new(
        name          => 'Green kWh/year',
        defaultFormat => '0soft',
        rows          => $assumptions->{rows},
        custom        => [
            %override
            ? '=IF(IV10,IV13,'
              . 'IV327*IV633+(IV313-IV328)*MAX(0,IV74/7*IV53-IV612-IV622)+'
              . '(IV303-IV329)*MIN(IV63,IV76/7*IV41))'
            : '=IV327*IV633+(IV313-IV328)*MAX(0,IV74/7*IV53-IV612-IV622)+'
              . '(IV303-IV329)*MIN(IV63,IV76/7*IV41)'
        ],
        arguments => {
            %override
            ? (
                IV10 => $override{total},
                IV13 => $override{green}
              )
            : (),
            IV303 => $assumptions->{columns}[3],
            IV313 => $assumptions->{columns}[2],
            IV327 => $assumptions->{columns}[4],
            IV328 => $assumptions->{columns}[4],
            IV329 => $assumptions->{columns}[4],
            IV41  => $assumptions->{columns}[1],
            IV53  => $assumptions->{columns}[0],
            IV612 => $model->{hoursByRedAmberGreen},
            IV622 => $model->{hoursByRedAmberGreen},
            IV63  => $model->{hoursByRedAmberGreen},
            IV633 => $model->{hoursByRedAmberGreen},
            IV74  => $daysInYear,
            IV76  => $daysInYear,
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0], map {
                    qr/\b$_\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + (
                            /^IV[1-5]/
                            ? $y
                            : 0
                        ),
                        $colh->{$_} + ( /^IV62/ ? 1 : /^IV63/ ? 2 : 0 ),
                        /^IV[1-5]/ ? 0 : 1,
                        1,
                      )
                } @$pha;
            };
        },
    );

    Columnset(
        name => (
            %override
            ? 'Additional consumption assumptions'
            : 'Consumption calculations'
          )
          . ' for illustrative customers',
        columns => \@columns2,
    );

    my $users = $assumptions->{rows}{list};

    my ( %map, @groups );
    for ( my $uid = 0 ; $uid < @$users ; ++$uid ) {
        my ( @list, @listmargin );
        my $short = my $user = $users->[$uid];
        $short =~ s/^Customer *[0-9]+ *//;
        my $regex = qr/$assumptions->{regex}{$user}/m;
        for ( my $tid = 0 ; $tid < @{ $allTariffs->{list} } ; ++$tid ) {
            next
              if $allTariffs->{groupid}
              && !defined $allTariffs->{groupid}[$tid];
            my $tariff = $allTariffs->{list}[$tid];
            next unless $tariff =~ /$regex/m;
            $tariff =~ s/^.*\n//s;
            my $row = "$short ($tariff)";
            push @list, $row;
            $map{$row} = [ $uid, $tid, $#list ];
            if ( $tariff =~ /^LDNO ([^:]+): (.+)/ ) {
                my $boundary = $1;
                my $atw      = $2;
                my $margin   = "Margin $boundary: $atw";
                next unless $margin =~ /$regex/m;
                if ( my $atwmapped = $map{"$short ($atw)"} ) {
                    my $marginrow = "$short ($margin)";
                    push @listmargin, $marginrow;
                    $map{$marginrow} = [ $atwmapped->[2], $#list ];
                }
            }
        }
        foreach (@listmargin) {
            push @list, $_;
            $map{$_} = [ undef, $map{$_}[0] - $#list, $map{$_}[1] - $#list, ];
        }
        push @groups, Labelset( name => $user, list => \@list );
    }

    my $fullRowset = Labelset( groups => \@groups );
    my @map = @map{ @{ $fullRowset->{list} } };

    my $charge = SpreadsheetModel::Custom->new(
        name => Label(
            '£/year', 'Annual charges for illustrative customers (£/year)',
        ),
        defaultFormat => '0softnz',
        rows          => $fullRowset,
        custom        => [
            '=0.01*(IV11*IV91+IV71*IV94)',
            '=0.01*(IV11*IV91+IV12*IV13/7*IV78*(IV92-IV911)+IV71*IV94)',
            '=0.01*(IV31*IV91+IV32*IV92+IV33*IV93+IV71*(IV94+IV2*IV95))',
            '=IV81-IV82',
        ],
        arguments => {
            IV11  => $totalUnits,
            IV12  => $assumptions->{columns}[3],
            IV13  => $assumptions->{columns}[1],
            IV2   => $assumptions->{columns}[5],
            IV31  => $red,
            IV32  => $amber,
            IV33  => $green,
            IV71  => $daysInYear,
            IV78  => $daysInYear,
            IV91  => $tariffTable->{'Unit rate 1 p/kWh'},
            IV911 => $tariffTable->{'Unit rate 1 p/kWh'},
            IV92  => $tariffTable->{'Unit rate 2 p/kWh'},
            IV93  => $tariffTable->{'Unit rate 3 p/kWh'},
            IV94  => $tariffTable->{'Fixed charge p/MPAN/day'},
            IV95  => $tariffTable->{'Capacity charge p/kVA/day'},
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                my $cellFormat =
                    $self->{rowFormats}[$y]
                  ? $wb->getFormat( $self->{rowFormats}[$y] )
                  : $format;
                return '', $cellFormat unless $map[$y];
                my ( $uid, $tid, $eid ) = @{ $map[$y] };
                unless ( defined $uid ) {
                    return '', $cellFormat, $formula->[3],
                      qr/\bIV81\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $self->{$wb}{row} + $y + $tid,
                        $self->{$wb}{col}
                      ),
                      qr/\bIV82\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $self->{$wb}{row} + $y + $eid,
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
        defaultFormat => '0.0soft',
        arithmetic    => '=IV1/IV2*1000',
        arguments     => {
            IV1 => $charge,
            IV2 => $totalUnits,
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

__DATA__
---
4001:
  - _table: 4001. Consumption assumptions for illustrative customers
  - 500kVA business: '^(?:LV|LV Sub|HV|LDNO .*) HH Metered$'
    500kVA continuous: '^(?:LV|LV Sub|HV|LDNO .*) HH Metered$'
    500kVA off-peak: '^(?:LV|LV Sub|HV|LDNO .*) HH Metered$'
    500kVA random: '^(?:LV|LV Sub|HV|LDNO .*) HH Metered$'
    5MVA business: '^HV HH Metered$'
    5MVA continuous: '^HV HH Metered$'
    5MVA off-peak: '^HV HH Metered$'
    5MVA random: '^HV HH Metered$'
    68kVA business: '^(?:|LDNO .*: |Margin.*: )(?:Small Non Domestic (?:Unrestricted|Two)|LV.*(?:HH Metered$|Medium)|LV Network)'
    68kVA continuous: '^(?:Small Non Domestic (?:Unrestricted|Two)|LV.*(?:HH Metered$|Medium)|LV Network)'
    68kVA off-peak: '^(?:Small Non Domestic (?:Unrestricted|Two)|LV.*(?:HH Metered$|Medium)|LV Network)'
    68kVA random: '^(?:Small Non Domestic (?:Unrestricted|Two)|LV.*(?:HH Metered$|Medium)|LV Network)'
    Average home: '^(?:|LDNO .*: |Margin.*: )(?:Domestic Unrestricted|LV Network Dom)'
    Average home x250: '^(?:LV|LV Sub|HV|LDNO .*) HH Metered$'
    Average home x2500: '^HV HH Metered$'
    Electric heating home: '^(?:|LDNO .*: |Margin.*: )(?:Domestic (?:Unrestricted|Two)|LV Network Dom)'
    Electric heating home x100: '^(?:LV|LV Sub|HV|LDNO .*) HH Metered$'
    Electric heating home x1000: '^HV HH Metered$'
    Low use home: '^(?:|LDNO .*: |Margin.*: )(?:Domestic Unrestricted|LV Network Dom)'
    _column: Tariff selection
  - 500kVA business: 66
    500kVA continuous: ''
    500kVA off-peak: ''
    500kVA random: ''
    5MVA business: 66
    5MVA continuous: ''
    5MVA off-peak: ''
    5MVA random: ''
    68kVA business: 66
    68kVA continuous: ''
    68kVA off-peak: ''
    68kVA random: ''
    Average home: 35
    Average home x250: 35
    Average home x2500: 35
    Electric heating home: 35
    Electric heating home x100: 35
    Electric heating home x1000: 35
    Low use home: 35
    _column: Peak-time hours/week
  - 500kVA business: 77
    500kVA continuous: ''
    500kVA off-peak: 74.6666666666667
    500kVA random: ''
    5MVA business: 77
    5MVA continuous: ''
    5MVA off-peak: 74.6666666666667
    5MVA random: ''
    68kVA business: 77
    68kVA continuous: ''
    68kVA off-peak: 74.6666666666667
    68kVA random: ''
    Average home: 49
    Average home x250: 49
    Average home x2500: 49
    Electric heating home: 49
    Electric heating home x100: 49
    Electric heating home x1000: 49
    Low use home: 49
    _column: Off-peak hours/week
  - 500kVA business: 355
    500kVA continuous: ''
    500kVA off-peak: ''
    500kVA random: ''
    5MVA business: 3550
    5MVA continuous: ''
    5MVA off-peak: ''
    5MVA random: ''
    68kVA business: 48.28
    68kVA continuous: ''
    68kVA off-peak: ''
    68kVA random: ''
    Average home: 1
    Average home x250: 250
    Average home x2500: 2500
    Electric heating home: 1.25
    Electric heating home x100: 125
    Electric heating home x1000: 1250
    Low use home: 0.5
    _column: Peak-time load (kW)
  - 500kVA business: 100
    500kVA continuous: ''
    500kVA off-peak: 450
    500kVA random: ''
    5MVA business: 1000
    5MVA continuous: ''
    5MVA off-peak: 4500
    5MVA random: ''
    68kVA business: 13.6
    68kVA continuous: ''
    68kVA off-peak: 61.2
    68kVA random: ''
    Average home: 0.1
    Average home x250: 25
    Average home x2500: 250
    Electric heating home: 1.5
    Electric heating home x100: 150
    Electric heating home x1000: 1500
    Low use home: 0.05
    _column: Off-peak load (kW)
  - 500kVA business: 100
    500kVA continuous: 450
    500kVA off-peak: ''
    500kVA random: 200
    5MVA business: 1000
    5MVA continuous: 4500
    5MVA off-peak: ''
    5MVA random: 2000
    68kVA business: 13.6
    68kVA continuous: 61.2
    68kVA off-peak: ''
    68kVA random: 27.2
    Average home: 0.4
    Average home x250: 100
    Average home x2500: 1000
    Electric heating home: 0.5
    Electric heating home x100: 50
    Electric heating home x1000: 500
    Low use home: 0.2
    _column: Load at other times (kW)
  - 500kVA business: 500
    500kVA continuous: 500
    500kVA off-peak: 500
    500kVA random: 500
    5MVA business: 5000
    5MVA continuous: 5000
    5MVA off-peak: 5000
    5MVA random: 5000
    68kVA business: 68
    68kVA continuous: 68
    68kVA off-peak: 68
    68kVA random: 68
    Average home: 6
    Average home x250: 500
    Average home x2500: 5000
    Electric heating home: 18
    Electric heating home x100: 500
    Electric heating home x1000: 5000
    Low use home: 3
    _column: Capacity (kVA)
  - _column: Total kWh/year
  - _column: Rate 2 kWh/year
  - _column: Load factor (kW/kVA)
