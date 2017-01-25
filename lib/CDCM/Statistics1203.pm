package CDCM;

=head Copyright licence and disclaimer

Copyright 2014-2017 Franck Latrémolière, Reckon LLP and others.

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

sub table1203 {
    my ($model) = @_;
    my @rows = map { "Illustrative customer $_"; } 1 .. 7;
    @rows = split /\n/, <<EOL;
Business long day
Business off peak
Business short day
Continuous high load factor
Continuous low load factor
Domestic two rate average
Domestic unrestricted average
Medium solar generator
Medium wind generator
Small business day
Small solar generator
Small wind generator
Unmetered continuous
Unmetered dawn to dusk
Unmetered dusk to dawn
Unmetered part night
EOL
    my $rowset = Labelset( list => \@rows );
    my $blank = [ map { '' } @{ $rowset->{list} } ];
    my @columns = map {
        Dataset(
            name          => $_,
            defaultFormat => '0hard',
            rows          => $rowset,
            data          => $blank,
            dataset       => $model->{dataset},
          )
      } ( map { "HH rate $_ kWh/year"; } 1 .. 3 ),
      'Capacity kVA', 'NHH rate 2 kWh/year';
    my $assumptions = Columnset(
        name     => 'Consumption assumptions for illustrative customers',
        number   => 1203,
        appendTo => $model->{inputTables},
        dataset  => $model->{dataset},
        columns  => \@columns,
    );
    $assumptions;
}

sub makeStatisticsTables1203 {

    my ( $model, $tariffTable, $daysInYear, $nonExcludedComponents,
        $componentMap, )
      = @_;

    my ($allTariffs) = values %$tariffTable;
    $allTariffs = $allTariffs->{rows};

    my $assumptions = $model->table1203;

    my ( $units1, $units2, $units3, $capacity, $rate2, ) =
      @{ $assumptions->{columns} };

    my ( @columns, $overrideTotal, $doNotUseDaysInYear );

    push @columns,
      my $totalUnits = Arithmetic(
        name          => 'Total kWh/year',
        defaultFormat => '0soft',
        arithmetic    => '=A1+A2+A3',
        arguments     => {
            A1 => $units1,
            A2 => $units2,
            A3 => $units3,
        },
      );

    push @columns,
      my $rate1 = Arithmetic(
        name          => 'NHH rate 1 kWh/year',
        defaultFormat => '0soft',
        arithmetic    => '=A1-A2',
        arguments     => {
            A1 => $totalUnits,
            A2 => $rate2,
        },
      );

    push @columns,
      my $kW = Arithmetic(
        name       => 'Average consumption (kW)',
        arithmetic => '=A1/A2/24',
        arguments  => { A1 => $totalUnits, A2 => $daysInYear, },
      );

    push @columns,
      Arithmetic(
        name          => 'Average capacity utilisation',
        defaultFormat => '%soft',
        arithmetic    => '=IF(A1,A2/A3,"")',
        arguments     => { A1 => $kW, A2 => $kW, A3 => $capacity, },
      );

    Columnset(
        name    => 'Consumption calculations for illustrative customers',
        columns => \@columns,
    );

    my $users = $assumptions->{rows}{list};

    my ( @groupList, %mapping, %margins );
    for ( my $uid = 0 ; $uid < @$users ; ++$uid ) {
        my (@tariffList);
        my $short = my $user = $users->[$uid];
        $short =~ s/^Customer *[0-9]+ *//;
        my $filter = sub {
            $_[0] !~ /\bunmeter|\bums\b|\bgener/i;
        };
        $filter = sub { 1; };
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
    my %ppyrow =
      map { ( $fullRowset->{list}[$_] => $_ ); } 0 .. $#{ $fullRowset->{list} };
    my @mapping = @mapping{ @{ $fullRowset->{list} } };

    my $ppy = SpreadsheetModel::Custom->new(
        name => Label(
            '£/year', 'Annual charges for illustrative customers (£/year)',
        ),
        defaultFormat => '0softnz',
        rows          => $fullRowset,
        custom        => [
            '=0.01*(A11*A91+A71*A94)',
            '=0.01*(A21*A91+A22*A92+A71*A94)',
            '=0.01*(A31*A91+A32*A92+A33*A93+A71*(A94+A42*A95))',
            '=A81-A82',
        ],
        arithmetic => 'Special calculation',
        arguments  => {
            A11 => $totalUnits,
            A21 => $rate1,
            A22 => $rate2,
            A31 => $units1,
            A32 => $units2,
            A33 => $units3,
            A42 => $capacity,
            A71 => $daysInYear,
            A78 => $daysInYear,
            A91 => $tariffTable->{'Unit rate 1 p/kWh'},
            A92 => $tariffTable->{'Unit rate 2 p/kWh'},
            A93 => $tariffTable->{'Unit rate 3 p/kWh'},
            A94 => $tariffTable->{'Fixed charge p/MPAN/day'},
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
            '£/MWh', 'Average charges for illustrative customers (£/MWh)'
        ),
        defaultFormat => '0.0soft',
        arithmetic    => '=A1/A2*1000',
        arguments     => {
            A1 => $ppy,
            A2 => $totalUnits,
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
                A1 => $ppy,
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    my $ppyrow = $ppyrow{ $atwRowset->{list}[$y] };
                    return '', $wb->getFormat('unavailable')
                      unless defined $ppyrow;
                    my $cellFormat =
                        $self->{rowFormats}[$y]
                      ? $wb->getFormat( $self->{rowFormats}[$y] )
                      : $format;
                    '', $cellFormat, $formula->[0],
                      qr/\bA1\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{A1} + $ppyrow,
                        $colh->{A1} );
                };
            },
        );
        my $ldnoWord =
          $model->{portfolio} && $model->{portfolio} =~ /qno/i ? 'QNO' : 'LDNO';
        my $marginTable = SpreadsheetModel::Custom->new(
            name          => "Apparent $ldnoWord margin (£/year)",
            defaultFormat => '0soft',
            rows          => $atwRowset,
            cols   => Labelset( list => [ map { "$_ margin"; } @boundaries ] ),
            custom => [ '=A2-A1', ],
            arithmetic => '=A2-A1',
            arguments  => {
                A1 => $ppy,
                A2 => $atwTable,
            },
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    my $ppyrow =
                      $margins{ $boundaries[$x] }{ $atwRowset->{list}[$y] };
                    $ppyrow = $ppyrow{$ppyrow} if $ppyrow;
                    return ' ', $wb->getFormat('unavailable')
                      unless defined $ppyrow;
                    my $cellFormat =
                        $self->{rowFormats}[$y]
                      ? $wb->getFormat( $self->{rowFormats}[$y] )
                      : $format;
                    '', $cellFormat, $formula->[0],
                      qr/\bA1\b/ =>
                      Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                        $rowh->{A1} + $ppyrow,
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
            name    => "$ldnoWord margins for illustrative customers (£/year)",
            columns => [ $atwTable, $marginTable, ],
          );
        $model->{sharedData}
          ->addStats( "$ldnoWord margins for illustrative customers",
            $model, $marginTable )
          if $model->{sharedData};

    }

}

1;
