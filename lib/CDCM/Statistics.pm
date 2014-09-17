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

    my @rows = sort grep { $_ ne '_column'; } keys %{ $colspec->[1] };

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
        ? ( appendTo => $model->{sharedData}{statsAssumptions} )
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
        rows          => $assumptions->{rows},
        arithmetic    => %override
        ? '=IF(IV5,IV6,(C8*E8+D8*F8+(168-C81-D81)*G8)*IV2/7)'
        : '=(C8*E8+D8*F8+(168-C81-D81)*G8)*IV2/7',
        arguments => {
            IV2 => $daysInYear,
            %override ? ( IV5 => $override{total}, IV6 => $override{total} )
            : (),
            C8  => $assumptions->{columns}[0],
            C81 => $assumptions->{columns}[0],
            D8  => $assumptions->{columns}[1],
            D81 => $assumptions->{columns}[1],
            E8  => $assumptions->{columns}[2],
            F8  => $assumptions->{columns}[3],
            G8  => $assumptions->{columns}[4],
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
1202:
  - _table: 1202. Consumption assumptions for illustrative customers
  - Customer A: '^(?:|LDNO .*: |Margin.*: )(?:Domestic Unrestricted|LV Network Dom)'
    Customer B: '^(?:|LDNO .*: |Margin.*: )(?:Domestic Unrestricted|LV Network Dom)'
    Customer C: '^(?:(Small Non )?Domestic (?:Unrestricted|Two)|LV.*Medium|LV Network)'
    Customer D: '^(?:|LDNO .*: |Margin.*: )(?:Small Non Domestic (?:Unrestricted|Two)|LV.*(?:HH Metered$|Medium)|LV Network)'
    Customer E: '^(?:Small Non Domestic (?:Unrestricted|Two)|LV.*(?:HH Metered$|Medium)|LV Network)'
    Customer F: '^(?:|LDNO .*: |Margin.*: )LV.*HH Metered$'
    Customer G: '^LV.*HH Metered$'
    Customer H: '^(?:LV|LV Sub|HV|LDNO .*) HH Metered$'
    Customer I: '^(?:LV|LV Sub|HV|LDNO .*) HH Metered$'
    Customer J: '^HV HH Metered$'
    Customer K: '^HV HH Metered$'
    Customer L: '^HV HH Metered$'
    _column: Tariff selection
  - Customer A: 35
    Customer B: 35
    Customer C: 35
    Customer D: 35
    Customer E: ''
    Customer F: 55
    Customer G: ''
    Customer H: 35
    Customer I: ''
    Customer J: ''
    Customer K: ''
    Customer L: 55
    _column: Peak-time hours/week
  - Customer A: ''
    Customer B: ''
    Customer C: 48
    Customer D: 100
    Customer E: ''
    Customer F: ''
    Customer G: 77
    Customer H: 100
    Customer I: ''
    Customer J: 77
    Customer K: ''
    Customer L: ''
    _column: Off-peak hours/week
  - Customer A: 0.4165
    Customer B: 0.8325
    Customer C: 0.75
    Customer D: 60
    Customer E: ''
    Customer F: 200
    Customer G: ''
    Customer H: 400
    Customer I: ''
    Customer J: ''
    Customer K: ''
    Customer L: 4500
    _column: Peak-time load (kW)
  - Customer A: ''
    Customer B: ''
    Customer C: 0.9995
    Customer D: 1
    Customer E: ''
    Customer F: ''
    Customer G: 200
    Customer H: 100
    Customer I: ''
    Customer J: 4500
    Customer K: ''
    Customer L: ''
    _column: Off-peak load (kW)
  - Customer A: 0.15
    Customer B: 0.3
    Customer C: 0.3
    Customer D: 50
    Customer E: 65
    Customer F: ''
    Customer G: ''
    Customer H: 300
    Customer I: 450
    Customer J: ''
    Customer K: 4500
    Customer L: ''
    _column: Load at other times (kW)
  - Customer A: ''
    Customer B: ''
    Customer C: 6
    Customer D: 68
    Customer E: 68
    Customer F: 250
    Customer G: 250
    Customer H: 500
    Customer I: 500
    Customer J: 5000
    Customer K: 5000
    Customer L: 5000
    _column: Capacity (kVA)
