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

    my @rows =
      sort { $colspec->[1]{$a} <=> $colspec->[1]{$b} }
      grep { !/^_/; } keys %{ $colspec->[1] };
    @rows = grep { !( $colspec->[1]{$_} % 10 ) } @rows
      if $model->{summary} =~ /brief/;

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
    } @$colspec[ 3 .. 8 ];

    my $result = Columnset(
        name => 'Consumption assumptions for illustrative customers',
        $model->{sharedData}
        ? (
            number   => 4001,
            appendTo => $model->{sharedData}{statsAssumptions},
          )
        : (),
        columns => \@columns,
        regex   => $colspec->[2],
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

    my $ppy = SpreadsheetModel::Custom->new(
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

    my $ppu = Arithmetic(
        name => Label(
            '£/MWh', 'Average charges for illustrative customers (£/MWh)'
        ),
        defaultFormat => '0.0soft',
        arithmetic    => '=IV1/IV2*1000',
        arguments     => {
            IV1 => $ppy,
            IV2 => $totalUnits,
        }
    );

    if ( $model->{sharedData} ) {
        $model->{sharedData}
          ->addStats( 'Illustrative charges (£/year)', $model, $ppy );
        $model->{sharedData}
          ->addStats( 'Illustrative charges (£/MWh)', $model, $ppu );
    }

    push @{ $model->{statisticsTables} },
      Columnset(
        name    => 'Statistics for illustrative customers',
        columns => [ $ppy, $ppu ],
      );

}

1;

__DATA__
---
4001:
  - _table: 4001. Consumption assumptions for illustrative customers
  - Domestic electric heat: 740
    Domestic low use: 700
    Domestic standard: 720
    Generator: 805
    Large business: 310
    Large continuous: 320
    Large housing electric: 355
    Large housing standard: 365
    Large intermittent: 345
    Large off-peak: 330
    Medium business: 210
    Medium continuous: 220
    Medium housing electric: 255
    Medium housing standard: 265
    Medium intermittent: 245
    Medium off-peak: 230
    Small business: 110
    Small continuous: 120
    Small intermittent: 145
    Small off-peak: 130
    XL business: 410
    XL continuous: 420
    XL housing electric: 455
    XL housing standard: 465
    XL intermittent: 445
    XL off-peak: 430
    _column: Order
  - Domestic electric heat: '(?:^|: )(?:LV Network Domestic|Domestic [UT])'
    Domestic low use: '(?:^|: )(?:LV Network Domestic|Domestic [UT])'
    Domestic standard: '(?:^|: )(?:LV Network Domestic|Domestic [UT])'
    Generator: '^HV.*Gener'
    Large business: '^(?:LV|LV Sub|HV|LDNO .*) HH Metered$'
    Large continuous: '^(?:LV|LV Sub|HV|LDNO .*) HH Metered$'
    Large housing electric: '^(?:LV|LV Sub|HV|LDNO .*) HH Metered$'
    Large housing standard: '^(?:LV|LV Sub|HV|LDNO .*) HH Metered$'
    Large intermittent: '^(?:LV|LV Sub|HV|LDNO .*) HH Metered$'
    Large off-peak: '^(?:LV|LV Sub|HV|LDNO .*) HH Metered$'
    Medium business: '^(?:Small|LV).*(?:Non[- ]Domestic(?: [UT]|$)|HH Metered$)'
    Medium continuous: '^(?:Small|LV).*(?:Non[- ]Domestic(?: [UT]|$)|HH Metered$)'
    Medium housing electric: '^(?:Small|LV).*(?:Non[- ]Domestic(?: [UT]|$)|HH Metered$)'
    Medium housing standard: '^(?:Small|LV).*(?:Non[- ]Domestic(?: [UT]|$)|HH Metered$)'
    Medium intermittent: '^(?:Small|LV).*(?:Non[- ]Domestic(?: [UT]|$)|HH Metered$)'
    Medium off-peak: '^(?:Small|LV).*(?:Non[- ]Domestic(?: [UT]|$)|HH Metered$)'
    Small business: '^(?:Small|LV).*Non[- ]Domestic(?: [UT]|$)'
    Small continuous: '^(?:Small|LV).*Non[- ]Domestic(?: [UT]|$)'
    Small intermittent: '^(?:Small|LV).*Non[- ]Domestic(?: [UT]|$)'
    Small off-peak: '^(?:Small|LV).*Non[- ]Domestic(?: [UT]|$)'
    XL business: '^(?:|LDNO .*)HV HH Metered$'
    XL continuous: '^(?:|LDNO .*)HV HH Metered$'
    XL housing electric: '^(?:|LDNO .*)HV HH Metered$'
    XL housing standard: '^(?:|LDNO .*)HV HH Metered$'
    XL intermittent: '^(?:|LDNO .*)HV HH Metered$'
    XL off-peak: '^(?:|LDNO .*)HV HH Metered$'
    _column: Tariff selection
  - Domestic electric heat: 35
    Domestic low use: 35
    Domestic standard: 35
    Generator: 0
    Large business: 66
    Large continuous: 0
    Large housing electric: 35
    Large housing standard: 35
    Large intermittent: 0
    Large off-peak: 0
    Medium business: 66
    Medium continuous: 0
    Medium housing electric: 35
    Medium housing standard: 35
    Medium intermittent: 0
    Medium off-peak: 0
    Small business: 66
    Small continuous: 0
    Small intermittent: 0
    Small off-peak: 0
    XL business: 66
    XL continuous: 0
    XL housing electric: 35
    XL housing standard: 35
    XL intermittent: 0
    XL off-peak: 0
    _column: Peak-time hours/week
  - Domestic electric heat: 49
    Domestic low use: 49
    Domestic standard: 49
    Generator: 0
    Large business: 0
    Large continuous: 0
    Large housing electric: 49
    Large housing standard: 49
    Large intermittent: 0
    Large off-peak: 74.6666666666667
    Medium business: 0
    Medium continuous: 0
    Medium housing electric: 49
    Medium housing standard: 49
    Medium intermittent: 0
    Medium off-peak: 74.6666666666667
    Small business: 0
    Small continuous: 0
    Small intermittent: 0
    Small off-peak: 74.6666666666667
    XL business: 0
    XL continuous: 0
    XL housing electric: 49
    XL housing standard: 49
    XL intermittent: 0
    XL off-peak: 74.6666666666667
    _column: Off-peak hours/week
  - Domestic electric heat: 1.1
    Domestic low use: 0.4
    Domestic standard: 0.8
    Generator: 0
    Large business: 350
    Large continuous: 0
    Large housing electric: 110
    Large housing standard: 200
    Large intermittent: 0
    Large off-peak: 0
    Medium business: 48.3
    Medium continuous: 0
    Medium housing electric: 11
    Medium housing standard: 20
    Medium intermittent: 0
    Medium off-peak: 0
    Small business: 16.1
    Small continuous: 0
    Small intermittent: 0
    Small off-peak: 0
    XL business: 3500
    XL continuous: 0
    XL housing electric: 1100
    XL housing standard: 2000
    XL intermittent: 0
    XL off-peak: 0
    _column: Peak-time load (kW)
  - Domestic electric heat: 1.6
    Domestic low use: 0.125
    Domestic standard: 0.25
    Generator: 0
    Large business: 0
    Large continuous: 0
    Large housing electric: 160
    Large housing standard: 62.5
    Large intermittent: 0
    Large off-peak: 450
    Medium business: 0
    Medium continuous: 0
    Medium housing electric: 16
    Medium housing standard: 6.25
    Medium intermittent: 0
    Medium off-peak: 62.1
    Small business: 0
    Small continuous: 0
    Small intermittent: 0
    Small off-peak: 20.7
    XL business: 0
    XL continuous: 0
    XL housing electric: 1600
    XL housing standard: 625
    XL intermittent: 0
    XL off-peak: 4500
    _column: Off-peak load (kW)
  - Domestic electric heat: 0.434817351598174
    Domestic low use: 0.194206621004566
    Domestic standard: 0.388413242009132
    Generator: 600
    Large business: 102.941176470588
    Large continuous: 450
    Large housing electric: 43.4817351598174
    Large housing standard: 97.1033105022831
    Large intermittent: 200
    Large off-peak: 0
    Medium business: 14.2058823529412
    Medium continuous: 62.1
    Medium housing electric: 4.34817351598174
    Medium housing standard: 9.71033105022831
    Medium intermittent: 27.6
    Medium off-peak: 0
    Small business: 4.73529411764706
    Small continuous: 20.7
    Small intermittent: 9.2
    Small off-peak: 0
    XL business: 1029.41176470588
    XL continuous: 4500
    XL housing electric: 434.817351598174
    XL housing standard: 971.033105022831
    XL intermittent: 2000
    XL off-peak: 0
    _column: Load at other times (kW)
  - Domestic electric heat: 18
    Domestic low use: 6
    Domestic standard: 9
    Generator: 1500
    Large business: 500
    Large continuous: 500
    Large housing electric: 500
    Large housing standard: 500
    Large intermittent: 500
    Large off-peak: 500
    Medium business: 69
    Medium continuous: 69
    Medium housing electric: 69
    Medium housing standard: 69
    Medium intermittent: 69
    Medium off-peak: 69
    Small business: 23
    Small continuous: 23
    Small intermittent: 23
    Small off-peak: 23
    XL business: 5000
    XL continuous: 5000
    XL housing electric: 5000
    XL housing standard: 5000
    XL intermittent: 5000
    XL off-peak: 5000
    _column: Capacity (kVA)
