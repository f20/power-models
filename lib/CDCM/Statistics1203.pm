﻿package CDCM;

# Copyright 2014-2025 Franck Latrémolière and others.
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

sub configuration1203 {
    my ( $model, $allTariffs, $tariffComponents, $componentMap ) = @_;
    my $configurationMethod =
      $model->{summary} =~ /allmetereddemand/i
      ? 'configuration_allMeteredDemand'
      : 'configuration_dcp268';
    $model->$configurationMethod( $allTariffs, $tariffComponents,
        $componentMap );
}

sub configuration_dcp268 {

    my ( $model, $allTariffs, $tariffComponents, $componentMap ) = @_;

    my @users = split /\n/, <<EOL;
Domestic Unrestricted 1
Domestic Two Rate 1
Domestic Two Rate 2
Small Non Domestic Unrestricted 1
Small Non Domestic Unrestricted 2
Small Non Domestic Two Rate 1
Small Non Domestic Two Rate 2
LV Medium Non-Domestic 1
LV Medium Non-Domestic 2
LV Medium Non-Domestic 3
LV Medium Non-Domestic 4
LV Medium Non-Domestic 5
LV Medium Non-Domestic 6
LV Medium Non-Domestic 7
LV Medium Non-Domestic 8
LV Network Non-Domestic Non-CT 1
LV Site Specific 1
LV Sub Site Specific 1
HV Site Specific 1
NHH UMS category A 1
NHH UMS category B 1
NHH UMS category C 1
LV UMS (Pseudo HH Metered) 1
EOL

    my ( %mapping, %margins );
    foreach my $uid ( 0 .. $#users ) {
        my $user  = $users[$uid];
        my $user2 = $user;
        $user2 =~ s/^Domestic Unrestricted/Domestic Aggregated/;
        $user2 =~ s/^Domestic Two Rate/Domestic Aggregated/;
        $user2 =~ s/^Small Non Domestic Unrestricted/Non-Domestic Aggregated/;
        $user2 =~ s/^Small Non Domestic Two Rate/Non-Domestic Aggregated/;
        $user2 =~ s/^LV Medium Non-Domestic/Non-Domestic Aggregated/;
        $user2 =~ s/^LV Network Domestic/Domestic Aggregated/;
        $user2 =~ s/^LV Network Non-Domestic Non-CT/Non-Domestic Aggregated/;
        $user2 =~ s/Site Specific/HH Metered/;
        $user2 =~ s/^NHH UMS category [ABCD]/Unmetered Supplies/;
        $user2 =~ s/^LV UMS \(Pseudo HH Metered\)/Unmetered Supplies/;

        for ( my $tid = 0 ; $tid < @{ $allTariffs->{list} } ; ++$tid ) {
            next
              if $allTariffs->{groupid}
              && !defined $allTariffs->{groupid}[$tid];
            my $tariff = $allTariffs->{list}[$tid];
            $tariff =~ s/^.*\n//s;
            if (   index( $user, $tariff ) == 0
                || index( $user2, $tariff ) == 0 )
            {
                $mapping{$user} = [ $uid, $tid ];
                last;
            }
        }
    }

    \@users, Labelset( list => \@users ), \%mapping, \%margins;

}

sub configuration_allMeteredDemand {
    my ( $model, $allTariffs, $tariffComponents, $componentMap ) = @_;
    my @users =
      $model->{table1203}
      ? @{ $model->{table1203} }
      : map { "Illustrative user $_"; } 1 .. 3;
    my $tariffFilter = sub { $_[0] !~ /\bunmeter|\bums\b|\bgener/i; };
    my @groupList;
    my ( %mapping, %margins );
    foreach my $uid ( 0 .. $#users ) {
        my $user = $users[$uid];
        my @tariffList;
        for ( my $tid = 0 ; $tid < @{ $allTariffs->{list} } ; ++$tid ) {
            next
              if $allTariffs->{groupid}
              && !defined $allTariffs->{groupid}[$tid];
            my $tariff = $allTariffs->{list}[$tid];
            next unless $tariffFilter->($tariff);
            $tariff =~ s/^.*\n//s;
            my $row = "$user ($tariff)";
            push @tariffList, $row;
            $mapping{$row} = [ $uid, $tid ];
            if ( $tariff =~ /^(?:LD|Q)NO ([^:]+): (.+)/ ) {
                $margins{$1}{"$user ($2)"} = $row;
            }
        }
        push @groupList, Labelset( name => $user, list => \@tariffList );
    }
    \@users, Labelset( groups => \@groupList ), \%mapping, \%margins;
}

sub table1203 {
    my ( $model, $userList ) = @_;
    my $rowset = Labelset( list => $userList );
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
    Columnset(
        name     => 'Consumption assumptions for illustrative customers',
        number   => 1203,
        appendTo => $model->{inputTables},
        dataset  => $model->{dataset},
        columns  => \@columns,
    );
}

sub makeStatisticsTables1203 {

    my ( $model, $tariffTable, $daysInYear, $tariffComponents, $componentMap, )
      = @_;
    my ($allTariffs) = values %$tariffTable;
    $allTariffs = $allTariffs->{rows};

    my ( $userList, $fullRowset, $mappingHash, $marginsHash ) =
      $model->configuration1203( $allTariffs, $tariffComponents,
        $componentMap );

    my ( $units1, $units2, $units3, $capacity, $rate2, ) =
      @{ $model->table1203($userList)->{columns} };
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
    my $kW = Arithmetic(
        name       => 'Average consumption (kW)',
        arithmetic => '=A1/A2/24',
        arguments  => { A1 => $totalUnits, A2 => $daysInYear, },
    );
    my $utilisation = Arithmetic(
        name          => 'Average capacity utilisation',
        defaultFormat => '%soft',
        arithmetic    => '=IF(A1,A2/A3,"")',
        arguments     => { A1 => $kW, A2 => $kW, A3 => $capacity, },
    );
    Columnset(
        name    => 'Consumption calculations for illustrative customers',
        columns => [ $totalUnits, $kW, $utilisation, ],
    );

    my $ppy = SpreadsheetModel::Custom->new(
        name => Label(
            '£/year', 'Annual charges for illustrative customers (£/year)',
        ),
        defaultFormat => '0softnz',
        rows          => $fullRowset,
        custom        => [
            '=0.01*(A11*A91+A71*A94)',
            '=0.01*((A21-A23)*A91+A22*A92+A71*A94)',
            '=0.01*(A31*A91+A32*A92+A33*A93+A71*(A94+A42*A95))',
        ],
        arithmetic => 'Special calculation',
        arguments  => {
            A11 => $totalUnits,
            A21 => $totalUnits,
            A22 => $rate2,
            A23 => $rate2,
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
                my $mapping = $mappingHash->{ $fullRowset->{list}[$y] }
                  or return '', $wb->getFormat('unavailable');
                my ( $uid, $tid ) = @$mapping;
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
                            : /^A1/         ? (
                                  $fullRowset->{groupid}
                                ? $fullRowset->{groupid}[$y]
                                : $y
                              )
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
        name => Label( 'p/kWh', 'Average p/kWh for illustrative customers' ),
        arithmetic => '=A1/A2*100',
        arguments  => {
            A1 => $ppy,
            A2 => $totalUnits,
        }
    );

    if ( $model->{sharedData} ) {
        $model->{sharedData}
          ->addStats( 'Illustrative charges (£/year)', $model, $ppy );
        $model->{sharedData}
          ->addStats( 'Illustrative charges (p/kWh)', $model, $ppu );
    }

    push @{ $model->{statisticsTables} },
      Columnset(
        name    => 'Charges for illustrative customers',
        columns => [ $ppy, $ppu, ],
      );

    if ( my @boundaries = sort keys %$marginsHash ) {
        my $atwRowset = Labelset(
            $fullRowset->{groups}
            ? (
                groups => [
                    map {
                        my @list = grep {
                            my $a = $_;
                            grep { $marginsHash->{$_}{$a} } @boundaries;
                        } @{ $_->{list} };
                        @list
                          ? Labelset( name => $_->{name}, list => \@list )
                          : ();
                    } @{ $fullRowset->{groups} }
                ]
              )
            : (
                list => [
                    grep {
                        my $a = $_;
                        grep { $marginsHash->{$_}{$a} } @boundaries;
                    } @{ $fullRowset->{list} }
                ]
            )
        );
        my %ppyrow =
          map { ( $fullRowset->{list}[$_] => $_ ); }
          0 .. $#{ $fullRowset->{list} };
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
        my $marginTable = SpreadsheetModel::Custom->new(
            name          => "Apparent $model->{ldnoWord} margin (£/year)",
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
                      $marginsHash->{ $boundaries[$x] }
                      { $atwRowset->{list}[$y] };
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
