package CDCM;

=head Copyright licence and disclaimer

Copyright 2014-2015 Franck Latrémolière, Reckon LLP and others.

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

    return $model->{sharedData}{statisticsAssumptions}
      if $model->{sharedData} && $model->{sharedData}{statisticsAssumptions};

    my $colspec;
    $colspec = $model->{statistics} if ref $model->{statistics} eq 'ARRAY';
    $colspec = $model->{dataset}{$1}
      if !$colspec
      && $model->{summary} =~ /([0-9]{3,4})/
      && $model->{dataset}
      && $model->{dataset}{$1};
    $colspec = $model->table1202->{
        $model->{statistics} && $model->{statistics} =~ /simple/
        ? '1202simple'
        : '1202'
      }
      if !$colspec && require CDCM::StatisticsDefaults;

    my @rows =
      sort { $colspec->[1]{$a} <=> $colspec->[1]{$b} }
      grep { !/^_/; } keys %{ $colspec->[1] };
    @rows = grep { !( $colspec->[1]{$_} % 10 ) } @rows
      if $model->{summary} =~ /brief/;

    my $rowset = Labelset( list => \@rows );

    my @columns = map {
        my $col = $_;
        Dataset(
            name          => $col->{_column},
            defaultFormat => $col->{_column} =~ /hours\/week/ ? '0.0hard'
            : $col->{_column} =~ /kVA/ ? '0hard'
            : '0.000hard',
            rows               => $rowset,
            data               => [ @$_{@rows} ],
            usePlaceholderData => 1,
            dataset            => $model->{dataset},
          )
    } @$colspec[ 3 .. 8 ];

    if (1) {

        my $blank = [ map { '' } @{ $rowset->{list} } ];

        push @columns,
          Dataset(
            name               => "Override\t$_ kWh/year",
            defaultFormat      => '0hard',
            rows               => $rowset,
            data               => $blank,
            usePlaceholderData => 1,
            dataset            => $model->{dataset},
          ) foreach qw(red amber green);

    }

    my $assumptions = Columnset(
        name   => 'Consumption assumptions for illustrative customers',
        number => 1202,
        $model->{sharedData}
        ? (
            appendTo        => $model->{sharedData}{statsAssumptions},
            ignoreDatasheet => 1,
          )
        : (
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
        ),
        columns => \@columns,
        regex   => $colspec->[2],
    );

    $model->{sharedData}{statisticsAssumptions} = $assumptions
      if $model->{sharedData};

    $assumptions;

}

sub makeStatisticsTables {

    my ( $model, $tariffTable, $daysInYear, $nonExcludedComponents,
        $componentMap, )
      = @_;

    my ($allTariffs) = values %$tariffTable;
    $allTariffs = $allTariffs->{rows};

    my $assumptions = $model->makeStatisticsAssumptions;

    my (
        $peakHours,   $offPeakHours,  $peakLoad,
        $offPeakLoad, $otherLoad,     $capacity,
        $overrideRed, $overrideAmber, $overrideGreen,
    ) = @{ $assumptions->{columns} };

    my @columns;
    my $overrideTotal;

    push @columns,
      $overrideTotal = Arithmetic(
        name          => "Total override kWh/year",
        defaultFormat => '0soft',
        arithmetic    => '=IV1+IV2+IV3',
        arguments     => {
            IV1 => $overrideRed,
            IV2 => $overrideAmber,
            IV3 => $overrideGreen,
        },
      ) if $overrideRed;

    push @columns,
      my $totalUnits = Arithmetic(
        name          => 'Total kWh/year',
        defaultFormat => '0soft',
        arithmetic    => $overrideTotal
        ? '=IF(IV8,IV9,(IV1*IV3+IV2*IV4+(168-IV11-IV21)*IV5)*IV7/7)'
        : '=(IV1*IV3+IV2*IV4+(168-IV11-IV21)*IV5)*IV7/7',
        arguments => {
            IV7 => $daysInYear,
            $overrideTotal
            ? (
                IV8 => $overrideTotal,
                IV9 => $overrideTotal,
              )
            : (),
            IV1  => $peakHours,
            IV11 => $peakHours,
            IV2  => $offPeakHours,
            IV21 => $offPeakHours,
            IV3  => $peakLoad,
            IV4  => $offPeakLoad,
            IV5  => $otherLoad,
        },
      );

    push @columns,
      my $rate2 = Arithmetic(
        name          => 'Rate 2 kWh/year',
        defaultFormat => '0soft',
        arithmetic    => '=IV1*IV3*IV2/7',
        arguments     => {
            IV1 => $offPeakHours,
            IV3 => $offPeakLoad,
            IV2 => $daysInYear,
        },
      );

    push @columns, my $red = SpreadsheetModel::Custom->new(
        name          => 'Red kWh/year',
        defaultFormat => '0soft',
        rows          => $assumptions->{rows},
        custom        => [
            $overrideTotal
            ? '=IF(IV10,IV11,IV321*IV613+(IV311-IV322)*MIN(IV61,IV72/7*IV51)+'
              . '(IV301-IV323)*MAX(0,IV77/7*IV43-IV631-IV623))'
            : '=IV321*IV613+(IV311-IV322)*MIN(IV61,IV72/7*IV51)+'
              . '(IV301-IV323)*MAX(0,IV77/7*IV43-IV631-IV623)'
        ],
        arithmetic => 'Special calculation',
        arguments  => {
            $overrideTotal
            ? (
                IV10 => $overrideTotal,
                IV11 => $overrideRed
              )
            : (),
            IV301 => $offPeakLoad,
            IV311 => $peakLoad,
            IV321 => $otherLoad,
            IV322 => $otherLoad,
            IV323 => $otherLoad,
            IV61  => $model->{hoursByRedAmberGreen},
            IV613 => $model->{hoursByRedAmberGreen},
            IV623 => $model->{hoursByRedAmberGreen},
            IV631 => $model->{hoursByRedAmberGreen},
            IV43  => $offPeakHours,
            IV51  => $peakHours,
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

    push @columns, my $amber = SpreadsheetModel::Custom->new(
        name          => 'Amber kWh/year',
        defaultFormat => '0soft',
        rows          => $assumptions->{rows},
        custom        => [
            $overrideTotal
            ? '=IF(IV10,IV12,'
              . 'IV324*IV624+(IV312-IV325)*MIN(IV620,MAX(0,IV73/7*IV52-IV611))+'
              . '(IV302-IV326)*MIN(IV621,MAX(0,IV75/7*IV42-IV632)))'
            : '=IV324*IV624+(IV312-IV325)*MIN(IV620,MAX(0,IV73/7*IV52-IV611))+'
              . '(IV302-IV326)*MIN(IV621,MAX(0,IV75/7*IV42-IV632))'
        ],
        arithmetic => 'Special calculation',
        arguments  => {
            $overrideTotal
            ? (
                IV10 => $overrideTotal,
                IV12 => $overrideAmber
              )
            : (),
            IV302 => $offPeakLoad,
            IV312 => $peakLoad,
            IV324 => $otherLoad,
            IV325 => $otherLoad,
            IV326 => $otherLoad,
            IV42  => $offPeakHours,
            IV52  => $peakHours,
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

    push @columns, my $green = SpreadsheetModel::Custom->new(
        name          => 'Green kWh/year',
        defaultFormat => '0soft',
        rows          => $assumptions->{rows},
        custom        => [
            $overrideTotal
            ? '=IF(IV10,IV13,'
              . 'IV327*IV633+(IV313-IV328)*MAX(0,IV74/7*IV53-IV612-IV622)+'
              . '(IV303-IV329)*MIN(IV63,IV76/7*IV41))'
            : '=IV327*IV633+(IV313-IV328)*MAX(0,IV74/7*IV53-IV612-IV622)+'
              . '(IV303-IV329)*MIN(IV63,IV76/7*IV41)'
        ],
        arithmetic => 'Special calculation',
        arguments  => {
            $overrideTotal
            ? (
                IV10 => $overrideTotal,
                IV13 => $overrideGreen
              )
            : (),
            IV303 => $offPeakLoad,
            IV313 => $peakLoad,
            IV327 => $otherLoad,
            IV328 => $otherLoad,
            IV329 => $otherLoad,
            IV41  => $offPeakHours,
            IV53  => $peakHours,
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
        name    => 'Consumption calculations for illustrative customers',
        columns => \@columns,
    );

    my $users = $assumptions->{rows}{list};

    my ( %map, @groups );
    for ( my $uid = 0 ; $uid < @$users ; ++$uid ) {
        my ( @list, @listmargin );
        my $short = my $user = $users->[$uid];
        $short =~ s/^Customer *[0-9]+ *//;
        my $filter;
        {
            my $regex = $assumptions->{regex}{$user};
            if ( $regex eq 'All-the-way demand' ) {
                $filter = sub {
                    $_[0] !~ /^LDNO /i && $_[0] !~ /\bunmeter|\bums\b|\bgener/i;
                };
            }
            elsif ( $regex eq 'All-the-way generation' ) {
                $filter = sub { $_[0] !~ /^LDNO /i && $_[0] =~ /\bgener/i; };
            }
            else {
                $regex = qr/$regex/m;
                $filter = sub { $_[0] =~ /$regex/m; };
            }
        }
        for ( my $tid = 0 ; $tid < @{ $allTariffs->{list} } ; ++$tid ) {
            next
              if $allTariffs->{groupid}
              && !defined $allTariffs->{groupid}[$tid];
            my $tariff = $allTariffs->{list}[$tid];
            next unless $filter->($tariff);
            $tariff =~ s/^.*\n//s;
            my $row = "$short ($tariff)";
            push @list, $row;
            $map{$row} = [ $uid, $tid, $#list ];
            if ( $tariff =~ /^LDNO ([^:]+): (.+)/ ) {
                my $boundary = $1;
                my $atw      = $2;
                my $margin   = "Margin $boundary: $atw";
                next unless $filter->($margin);
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
        arithmetic => 'Special calculation',
        arguments  => {
            IV11  => $totalUnits,
            IV12  => $offPeakLoad,
            IV13  => $offPeakHours,
            IV2   => $capacity,
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
        name    => 'Charges for illustrative customers',
        columns => [ $ppy, $ppu, ],
      );

}

1;
