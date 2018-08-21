package CDCM;

=head Copyright licence and disclaimer

Copyright 2014-2018 Franck Latrémolière, Reckon LLP and others.

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

    my $table1202ByColumn;
    $table1202ByColumn = $model->{statistics}
      if ref $model->{statistics} eq 'ARRAY';
    $table1202ByColumn = $model->{dataset}{$1}
      if !$table1202ByColumn
      && $model->{summary} =~ /(120[2-9])/
      && $model->{dataset}
      && $model->{dataset}{$1};
    $table1202ByColumn = $model->table1202
      if !$table1202ByColumn && require CDCM::StatisticsDefaults;

    my @rows =
      sort { $table1202ByColumn->[1]{$a} <=> $table1202ByColumn->[1]{$b} }
      grep { !/^_/; } keys %{ $table1202ByColumn->[1] };
    @rows = grep { !( $table1202ByColumn->[1]{$_} % 10 ) } @rows
      unless $model->{summary} =~ /long/;

    my $rowset = Labelset( list => \@rows );

    my @columns = map {
        my $dataColumn = $_;
        Dataset(
            name          => $dataColumn->{_column},
            defaultFormat => $dataColumn->{_column} =~ /hours\/week/ ? '0.0hard'
            : $dataColumn->{_column} =~ /kVA/ ? '0hard'
            : '0.000hard',
            rows               => $rowset,
            data               => [ @$_{@rows} ],
            usePlaceholderData => 1,
            dataset            => $model->{dataset},
          )
    } @$table1202ByColumn[ 3 .. 8 ];

    if ( $model->{summary} =~ /override/i ) {

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
        name   => 'Illustrative customer usage assumptions',
        number => 1202,
        $model->{sharedData}
        ? (
            appendTo        => $model->{sharedData}{statsAssumptions},
            dataset         => $model->{sharedData}{dataset},
            ignoreDatasheet => 1,
          )
        : (
            appendTo => $model->{inputTables},
            dataset  => $model->{dataset},
        ),
        columns => \@columns,
        regex   => $table1202ByColumn->[2],
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

    my ( @columns, $overrideTotal, $doNotUseDaysInYear );

    if ($overrideRed) {
        push @columns,
          $overrideTotal = Arithmetic(
            name          => "Total override kWh/year",
            defaultFormat => '0soft',
            arithmetic    => '=A1+A2+A3',
            arguments     => {
                A1 => $overrideRed,
                A2 => $overrideAmber,
                A3 => $overrideGreen,
            },
          );
    }
    else {
        $doNotUseDaysInYear = $model->{summary} =~ /365.?25/;
    }

    push @columns,
      my $totalUnits = Arithmetic(
        name          => 'Total kWh/year',
        defaultFormat => '0soft',
        arithmetic    => $overrideTotal
        ? '=IF(A8,A9,(A1*A3+A2*A4+(168-A11-A21)*A5)*A7/7)'
        : $doNotUseDaysInYear ? '=(A1*A3+A2*A4+(168-A11-A21)*A5)*365.25/7'
        : '=(A1*A3+A2*A4+(168-A11-A21)*A5)*A7/7',
        arguments => {
            A7 => $daysInYear,
            $overrideTotal
            ? (
                A8 => $overrideTotal,
                A9 => $overrideTotal,
              )
            : (),
            A1  => $peakHours,
            A11 => $peakHours,
            A2  => $offPeakHours,
            A21 => $offPeakHours,
            A3  => $peakLoad,
            A4  => $offPeakLoad,
            A5  => $otherLoad,
        },
      );

    push @columns,
      my $rate2 = Arithmetic(
        name          => 'Rate 2 kWh/year',
        defaultFormat => '0soft',
        arithmetic => $doNotUseDaysInYear ? '=A1*A3*365.25/7' : '=A1*A3*A2/7',
        arguments  => {
            A1 => $offPeakHours,
            A3 => $offPeakLoad,
            A2 => $daysInYear,
        },
      );

    push @columns, my $red = SpreadsheetModel::Custom->new(
        name          => 'Red kWh/year',
        defaultFormat => '0soft',
        rows          => $assumptions->{rows},
        custom        => [
            $overrideTotal
            ? '=IF(A10,A11,A321*A613+(A311-A322)*MIN(A61,A72/7*A51)+'
              . '(A301-A323)*MAX(0,A77/7*A43-A631-A623))'
            : $doNotUseDaysInYear
            ? '=A321*A613+(A311-A322)*MIN(A61,365.25/7*A51)+'
              . '(A301-A323)*MAX(0,365.25/7*A43-A631-A623)'
            : '=A321*A613+(A311-A322)*MIN(A61,A72/7*A51)+'
              . '(A301-A323)*MAX(0,A77/7*A43-A631-A623)'
        ],
        arithmetic => 'Special calculation',
        arguments  => {
            $overrideTotal
            ? (
                A10 => $overrideTotal,
                A11 => $overrideRed
              )
            : (),
            A301 => $offPeakLoad,
            A311 => $peakLoad,
            A321 => $otherLoad,
            A322 => $otherLoad,
            A323 => $otherLoad,
            A61  => $model->{hoursByRedAmberGreen},
            A613 => $model->{hoursByRedAmberGreen},
            A623 => $model->{hoursByRedAmberGreen},
            A631 => $model->{hoursByRedAmberGreen},
            A43  => $offPeakHours,
            A51  => $peakHours,
            A72  => $daysInYear,
            A77  => $daysInYear,
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0], map {
                    qr/\b$_\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + (
                            /^A[1-5]/
                            ? $y
                            : 0
                        ),
                        $colh->{$_} + ( /^A62/ ? 1 : /^A63/ ? 2 : 0 ),
                        /^A[1-5]/ ? 0 : 1,
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
            ? '=IF(A10,A12,'
              . 'A324*A624+(A312-A325)*MIN(A620,MAX(0,A73/7*A52-A611))+'
              . '(A302-A326)*MIN(A621,MAX(0,A75/7*A42-A632)))'
            : $doNotUseDaysInYear
            ? '=A324*A624+(A312-A325)*MIN(A620,MAX(0,365.25/7*A52-A611))+'
              . '(A302-A326)*MIN(A621,MAX(0,365.25/7*A42-A632))'
            : '=A324*A624+(A312-A325)*MIN(A620,MAX(0,A73/7*A52-A611))+'
              . '(A302-A326)*MIN(A621,MAX(0,A75/7*A42-A632))'
        ],
        arithmetic => 'Special calculation',
        arguments  => {
            $overrideTotal
            ? (
                A10 => $overrideTotal,
                A12 => $overrideAmber
              )
            : (),
            A302 => $offPeakLoad,
            A312 => $peakLoad,
            A324 => $otherLoad,
            A325 => $otherLoad,
            A326 => $otherLoad,
            A42  => $offPeakHours,
            A52  => $peakHours,
            A611 => $model->{hoursByRedAmberGreen},
            A620 => $model->{hoursByRedAmberGreen},
            A621 => $model->{hoursByRedAmberGreen},
            A624 => $model->{hoursByRedAmberGreen},
            A632 => $model->{hoursByRedAmberGreen},
            A73  => $daysInYear,
            A75  => $daysInYear,
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0], map {
                    qr/\b$_\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + ( /^A[1-5]/ ? $y : 0 ),
                        $colh->{$_} + ( /^A62/ ? 1 : /^A63/ ? 2 : 0 ),
                        /^A[1-5]/ ? 0 : 1, 1,
                      )
                } @$pha;
            };
        },
    );

    push @columns,
      my $green = SpreadsheetModel::Custom
      ->new(    # Hack to avoid splitting $model->{hoursByRedAmberGreen}
        name          => 'Green kWh/year',
        defaultFormat => '0soft',
        rows          => $assumptions->{rows},
        custom        => [
            $overrideTotal
            ? '=IF(A10,A13,'
              . 'A327*A633+(A313-A328)*MAX(0,A74/7*A53-A612-A622)+'
              . '(A303-A329)*MIN(A63,A76/7*A41))'
            : $doNotUseDaysInYear
            ? '=A327*A633+(A313-A328)*MAX(0,365.25/7*A53-A612-A622)+'
              . '(A303-A329)*MIN(A63,365.25/7*A41)'
            : '=A327*A633+(A313-A328)*MAX(0,A74/7*A53-A612-A622)+'
              . '(A303-A329)*MIN(A63,A76/7*A41)'
        ],
        arithmetic => 'Special calculation',
        arguments  => {
            $overrideTotal
            ? (
                A10 => $overrideTotal,
                A13 => $overrideGreen
              )
            : (),
            A303 => $offPeakLoad,
            A313 => $peakLoad,
            A327 => $otherLoad,
            A328 => $otherLoad,
            A329 => $otherLoad,
            A41  => $offPeakHours,
            A53  => $peakHours,
            A612 => $model->{hoursByRedAmberGreen},
            A622 => $model->{hoursByRedAmberGreen},
            A63  => $model->{hoursByRedAmberGreen},
            A633 => $model->{hoursByRedAmberGreen},
            A74  => $daysInYear,
            A76  => $daysInYear,
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                '', $format, $formula->[0], map {
                    qr/\b$_\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + (
                            /^A[1-5]/
                            ? $y
                            : 0
                        ),
                        $colh->{$_} + ( /^A62/ ? 1 : /^A63/ ? 2 : 0 ),
                        /^A[1-5]/ ? 0 : 1,
                        1,
                      )
                } @$pha;
            };
        },
      );

    Columnset(
        name    => 'Illustrative customer usage calculations',
        columns => \@columns,
    );

    my $users = $assumptions->{rows}{list};

    my ( @groupList, %mapping, %margins );
    for ( my $uid = 0 ; $uid < @$users ; ++$uid ) {
        my (@tariffList);
        my $short = my $user = $users->[$uid];
        $short =~ s/^Customer *[0-9]+ *//;
        my $filter;
        {
            my $regex = $assumptions->{regex}{$user};
            if ( $regex eq 'All-the-way metered demand' ) {
                $filter = sub {
                    $_[0] !~ /^(?:LD|Q)NO /i
                      && $_[0] !~ /\bunmeter|\bums\b|\bgener|off.?peak/i;
                };
            }
            elsif ( $regex eq 'All-the-way generation' ) {
                $filter =
                  sub { $_[0] !~ /^(?:LD|Q)NO /i && $_[0] =~ /\bgener/i; };
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
            push @tariffList, $row;
            $mapping{$row} = [ $uid, $tid, $#tariffList ];
            if ( $tariff =~ /^(?:LD|Q)NO ([^:]+): (.+)/ ) {
                $margins{$1}{"$short ($2)"} = $row;
            }
        }
        push @groupList, Labelset( name => $user, list => \@tariffList );
    }

    my $fullRowset = Labelset( groups => \@groupList );
    my %rowNumberByRowLabel =
      map { ( $fullRowset->{list}[$_] => $_ ); } 0 .. $#{ $fullRowset->{list} };
    my @mapping = @mapping{ @{ $fullRowset->{list} } };

    my $annualChargeUnits = SpreadsheetModel::Custom->new(
        name => Label(
            'Units £/year',
            'Illustrative customer annual usage charge (£/year)',
        ),
        defaultFormat => '0softnz',
        rows          => $fullRowset,
        custom        => [
            '=0.01*A11*A91',
            '=0.01*(A11*A91+A12*A13/7*A78*(A92-A911))',
            '=0.01*(A31*A91+A32*A92+A33*A93)',
            '=A81-A82',
        ],
        arithmetic => 'Special calculation',
        arguments  => {
            A11  => $totalUnits,
            A12  => $offPeakLoad,
            A13  => $offPeakHours,
            A31  => $red,
            A32  => $amber,
            A33  => $green,
            A78  => $daysInYear,
            A91  => $tariffTable->{'Unit rate 1 p/kWh'},
            A911 => $tariffTable->{'Unit rate 1 p/kWh'},
            A92  => $tariffTable->{'Unit rate 2 p/kWh'},
            A93  => $tariffTable->{'Unit rate 3 p/kWh'},
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                my $cellFormat =
                    $self->{rowFormats}[$y]
                  ? $wb->getFormat( $self->{rowFormats}[$y] )
                  : $format;
                return '', $wb->getFormat('unavailable') unless $mapping[$y];
                my ( $uid, $tid, $eid ) = @{ $mapping[$y] };
                unless ( defined $uid ) {
                    return '', $cellFormat, $formula->[3],
                      qr/\bA81\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $self->{$wb}{row} + $y + $tid,
                        $self->{$wb}{col}
                      ),
                      qr/\bA82\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $self->{$wb}{row} + $y + $eid,
                        $self->{$wb}{col} );
                }
                my $tariff = $allTariffs->{list}[$tid];
                '', $cellFormat,
                  $formula->[
                    $componentMap->{$tariff}{'Capacity charge p/kVA/day'} ? 2
                  : $componentMap->{$tariff}{'Unit rate 2 p/kWh'}         ? 1
                  : 0
                  ],
                  map {
                    qr/\b$_\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + (
                              /^A9/         ? $tid
                            : /^A(?:[2-5])/ ? $uid
                            : /^A1/         ? $fullRowset->{groupid}[$y]
                            : 0
                        ),
                        $colh->{$_} + ( /^A62/ ? 1 : /^A63/ ? 2 : 0 ),
                        1, 1,
                      )
                  } @$pha;
            };
        },
    );

    my $annualChargeFixed = SpreadsheetModel::Custom->new(
        name => Label(
            'Fixed £/year',
            'Illustrative customer annual fixed charge (£/year)',
        ),
        defaultFormat => '0softnz',
        rows          => $fullRowset,
        custom        => [ '=0.01*A71*A94', '=A81-A82', ],
        arithmetic    => 'Special calculation',
        arguments     => {
            A11 => $totalUnits,
            A71 => $daysInYear,
            A94 => $tariffTable->{'Fixed charge p/MPAN/day'},
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                my $cellFormat =
                    $self->{rowFormats}[$y]
                  ? $wb->getFormat( $self->{rowFormats}[$y] )
                  : $format;
                return '', $wb->getFormat('unavailable') unless $mapping[$y];
                my ( $uid, $tid, $eid ) = @{ $mapping[$y] };
                unless ( defined $uid ) {
                    return '', $cellFormat, $formula->[1],
                      qr/\bA81\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $self->{$wb}{row} + $y + $tid,
                        $self->{$wb}{col}
                      ),
                      qr/\bA82\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $self->{$wb}{row} + $y + $eid,
                        $self->{$wb}{col} );
                }
                my $tariff = $allTariffs->{list}[$tid];
                '', $cellFormat, $formula->[0], map {
                    qr/\b$_\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + (
                              /^A9/         ? $tid
                            : /^A(?:[2-5])/ ? $uid
                            : /^A1/         ? $fullRowset->{groupid}[$y]
                            : 0
                        ),
                        $colh->{$_} + ( /^A62/ ? 1 : /^A63/ ? 2 : 0 ),
                        1, 1,
                      )
                } @$pha;
            };
        },
    );

    my $annualChargeCapacity = SpreadsheetModel::Custom->new(
        name => Label(
            'Capacity £/year',
            'Illustrative customer annual capacity charge (£/year)',
        ),
        defaultFormat => '0softnz',
        rows          => $fullRowset,
        custom        => [ '=0.01*A2*A95', '=A81-A82', ],
        arithmetic    => 'Special calculation',
        arguments     => {
            A2  => $capacity,
            A71 => $daysInYear,
            A95 => $tariffTable->{'Capacity charge p/kVA/day'},
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                my $cellFormat =
                    $self->{rowFormats}[$y]
                  ? $wb->getFormat( $self->{rowFormats}[$y] )
                  : $format;
                return '', $wb->getFormat('unavailable') unless $mapping[$y];
                my ( $uid, $tid, $eid ) = @{ $mapping[$y] };
                unless ( defined $uid ) {
                    return '', $cellFormat, $formula->[3],
                      qr/\bA81\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $self->{$wb}{row} + $y + $tid,
                        $self->{$wb}{col}
                      ),
                      qr/\bA82\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $self->{$wb}{row} + $y + $eid,
                        $self->{$wb}{col} );
                }
                my $tariff = $allTariffs->{list}[$tid];
                return '', $wb->getFormat('unavailable')
                  unless $componentMap->{$tariff}{'Capacity charge p/kVA/day'};
                '', $cellFormat, $formula->[0], map {
                    qr/\b$_\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + (
                              /^A9/         ? $tid
                            : /^A(?:[2-5])/ ? $uid
                            : /^A1/         ? $fullRowset->{groupid}[$y]
                            : 0
                        ),
                        $colh->{$_} + ( /^A62/ ? 1 : /^A63/ ? 2 : 0 ),
                        1, 1,
                      )
                } @$pha;
            };
        },
    );

    my $annualCharge = SpreadsheetModel::Custom->new(
        name => Label(
            'Total £/year', 'Illustrative customer annual charge (£/year)',
        ),
        defaultFormat => '0softnz',
        rows          => $fullRowset,
        custom        => [
            '=0.01*(A11*A91+A71*A94)',
            '=0.01*(A11*A91+A12*A13/7*A78*(A92-A911)+A71*A94)',
            '=0.01*(A31*A91+A32*A92+A33*A93+A71*(A94+A2*A95))',
            '=A81-A82',
        ],
        arithmetic => 'Special calculation',
        arguments  => {
            A11  => $totalUnits,
            A12  => $offPeakLoad,
            A13  => $offPeakHours,
            A2   => $capacity,
            A31  => $red,
            A32  => $amber,
            A33  => $green,
            A71  => $daysInYear,
            A78  => $daysInYear,
            A91  => $tariffTable->{'Unit rate 1 p/kWh'},
            A911 => $tariffTable->{'Unit rate 1 p/kWh'},
            A92  => $tariffTable->{'Unit rate 2 p/kWh'},
            A93  => $tariffTable->{'Unit rate 3 p/kWh'},
            A94  => $tariffTable->{'Fixed charge p/MPAN/day'},
            A95  => $tariffTable->{'Capacity charge p/kVA/day'},
        },
        wsPrepare => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                my $cellFormat =
                    $self->{rowFormats}[$y]
                  ? $wb->getFormat( $self->{rowFormats}[$y] )
                  : $format;
                return '', $wb->getFormat('unavailable') unless $mapping[$y];
                my ( $uid, $tid, $eid ) = @{ $mapping[$y] };
                unless ( defined $uid ) {
                    return '', $cellFormat, $formula->[3],
                      qr/\bA81\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $self->{$wb}{row} + $y + $tid,
                        $self->{$wb}{col}
                      ),
                      qr/\bA82\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $self->{$wb}{row} + $y + $eid,
                        $self->{$wb}{col} );
                }
                my $tariff = $allTariffs->{list}[$tid];
                '', $cellFormat,
                  $formula->[
                    $componentMap->{$tariff}{'Capacity charge p/kVA/day'} ? 2
                  : $componentMap->{$tariff}{'Unit rate 2 p/kWh'}         ? 1
                  : 0
                  ],
                  map {
                    qr/\b$_\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{$_} + (
                              /^A9/         ? $tid
                            : /^A(?:[2-5])/ ? $uid
                            : /^A1/         ? $fullRowset->{groupid}[$y]
                            : 0
                        ),
                        $colh->{$_} + ( /^A62/ ? 1 : /^A63/ ? 2 : 0 ),
                        1, 1,
                      )
                  } @$pha;
            };
        },
    );

    my $ppu = Arithmetic(
        name => Label(
            'Average £/MWh',
            'Illustrative customer average charge (£/MWh)'
        ),
        defaultFormat => '0.0soft',
        arithmetic    => '=A1/A2*1000',
        arguments     => {
            A1 => $annualCharge,
            A2 => $totalUnits,
        }
    );

    my $averageUnitRate = Arithmetic(
        name => Label(
            'Average £/MWh',
            'Illustrative customer average unit rate (p/kWh)'
        ),
        arithmetic => '=A1/A2*100',
        arguments  => {
            A1 => $annualChargeUnits,
            A2 => $totalUnits,
        }
    );

    if ( $model->{sharedData} ) {
        $model->{sharedData}
          ->addStats( 'Illustrative customer annual charge (£/year)',
            $model, $annualCharge );
        $model->{sharedData}
          ->addStats( 'Illustrative customer average charge (£/MWh)',
            $model, $ppu );
        $model->{sharedData}
          ->addStats( 'Illustrative customer usage (kWh/year)',
            $model, $totalUnits );
        $model->{sharedData}->addStats( 'Illustrative usage charges (£/year)',
            $model, $annualChargeUnits );
        $model->{sharedData}->addStats( 'Illustrative fixed charges (£/year)',
            $model, $annualChargeFixed );
        $model->{sharedData}
          ->addStats( 'Illustrative capacity charges (£/year)',
            $model, $annualChargeCapacity );
        $model->{sharedData}->addStats( 'Illustrative average unit rate p/kWh)',
            $model, $averageUnitRate );
    }

    push @{ $model->{statisticsTables} },
      Columnset(
        name    => 'Charges for illustrative customers',
        columns => [
            $annualCharge,      $ppu,                  $annualChargeUnits,
            $annualChargeFixed, $annualChargeCapacity, $averageUnitRate,
        ],
      );

    if ( my @boundaries = sort keys %margins ) {
        my $atwRowset = Labelset(
            groups => [
                map {
                    my @list = grep {
                        my $a = $_;
                        grep { $margins{$_}{$a} } @boundaries;
                    } @{ $_->{list} };
                    @list ? Labelset( name => $_->{name}, list => \@list ) : ();
                } @groupList
            ]
        );
        my $atwTable = SpreadsheetModel::Custom->new(
            name => Label( 'All the way', 'All-the-way charge (£/year)' ),
            defaultFormat => '0copy',
            rows          => $atwRowset,
            custom        => [ '=A1', ],
            arithmetic    => '=A1',
            arguments     => {
                A1 => $annualCharge,
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    my $row = $rowNumberByRowLabel{ $atwRowset->{list}[$y] };
                    return '', $wb->getFormat('unavailable')
                      unless defined $row;
                    my $cellFormat =
                        $self->{rowFormats}[$y]
                      ? $wb->getFormat( $self->{rowFormats}[$y] )
                      : $format;
                    '', $cellFormat, $formula->[0],
                      qr/\bA1\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{A1} + $row,
                        $colh->{A1} );
                };
            },
        );
        my $marginTable = SpreadsheetModel::Custom->new(
            name          => "Apparent $model->{ldnoWord} margin (£/year)",
            defaultFormat => '0soft',
            rows          => $atwRowset,
            cols   => Labelset( list => [ map { "$_ margin"; } @boundaries ] ),
            custom => [ '=A2-A1', ],
            arithmetic => '=A2-A1',
            arguments  => {
                A1 => $annualCharge,
                A2 => $atwTable,
            },
            cellFormats => [
                map {
                    my $map = $_;
                    [
                        map {
                            $map && defined $map->{$_} ? undef : 'unavailable';
                        } @{ $atwRowset->{list} }
                    ]
                } @margins{@boundaries}
            ],
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                my $unavailableFormat = $wb->getFormat('unavailable');
                sub {
                    my ( $x, $y ) = @_;
                    my $row =
                      $margins{ $boundaries[$x] }{ $atwRowset->{list}[$y] };
                    $row = $rowNumberByRowLabel{$row} if $row;
                    return '', $unavailableFormat unless defined $row;
                    my $cellFormat =
                        $self->{rowFormats}[$y]
                      ? $wb->getFormat( $self->{rowFormats}[$y] )
                      : $format;
                    '', $cellFormat, $formula->[0],
                      qr/\bA1\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{A1} + $row,
                        $colh->{A1}
                      ),
                      qr/\bA2\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{A2} + $y,
                        $colh->{A2} );
                };
            },
        );
        push @{ $model->{statisticsTables} },
          Columnset(
            name =>
              "$model->{ldnoWord} margins for illustrative customers (£/year)",
            columns => [ $atwTable, $marginTable, ],
          );
        $model->{sharedData}
          ->addStats( "$model->{ldnoWord} margins for illustrative customers",
            $model, $marginTable )
          if $model->{sharedData};

    }

}

1;
