﻿package CDCM;

# Copyright 2012-2016 Franck Latrémolière, Reckon LLP and others.
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

sub degroupTariffs {

    my ( $model, $allComponents, $tariffTable ) = @_;

    # $unitsInYear is calculated for the summary table
    my $unitsInYear = Arithmetic(
        noCopy     => 1,
        name       => 'All units (MWh)',
        arithmetic => '='
          . join( '+', map { "A$_" } 1 .. $model->{maxUnitRates} ),
        arguments => {
            map {
                ( "A$_" =>
                      $model->{ungrouped}{volumeData}{"Unit rate $_ p/kWh"} )
            } 1 .. $model->{maxUnitRates}
        },
        defaultFormat => '0soft',
    );

    my %applicability = map {
        $_ => Constant(
            name          => $_,
            defaultFormat => '0con',
            rows          => $model->{ungrouped}{allTariffsByGroup},
            data          => [
                m#/day#i
                ? [
                    map { /related mpan/i ? 0 : 1; }
                      @{ $model->{ungrouped}{allTariffsByGroup}{list} }
                  ]
                : [
                    map { 1; } @{ $model->{ungrouped}{allTariffsByGroup}{list} }
                ]
            ],
        );
    } @$allComponents;
    Columnset(
        name    => 'Component applicability matrix',
        columns => [ @applicability{@$allComponents} ]
    );
    my ( @columns1, @columns2, $tariffTableNew );
    foreach (@$allComponents) {
        my $defaultFormat = $tariffTable->{$_}{defaultFormat} || '0.000soft';
        $defaultFormat =~ s/soft/copy/;
        push @columns1,
          my $flattened = Arithmetic(
            name       => $_,
            arithmetic => '=A1',
            rows       => $model->{ungrouped}{allTariffGroups},
            cols       => $tariffTable->{$_}{cols},
            arguments  => { A1 => $tariffTable->{$_} },
          );
        $defaultFormat =~ s/copy/soft/;
        push @columns2,
          my $exploded = Arithmetic(
            name          => $_,
            defaultFormat => $defaultFormat,
            arithmetic    => '=A1*A2',
            rows          => $model->{ungrouped}{allTariffsByGroup},
            cols          => $tariffTable->{$_}{cols},
            arguments     => { A1 => $flattened, A2 => $applicability{$_}, },
          );
        $defaultFormat =~ s/soft/copy/;
        $tariffTableNew->{$_} = Arithmetic(
            name          => $_,
            defaultFormat => $defaultFormat,
            arithmetic    => '=A1',
            rows          => $model->{ungrouped}{allTariffsByEndUser},
            cols          => $tariffTable->{$_}{cols},
            arguments     => { A1 => $exploded },
        );
    }

    Columnset(
        name    => 'Tariff table before flattening',
        columns => [ @{$tariffTable}{@$allComponents} ],
    ) unless grep { $_->{location}; } @{$tariffTable}{@$allComponents};

    Columnset(
        name    => 'Tariff table before tariff degrouping',
        columns => \@columns1,
    );

    Columnset(
        name    => 'Tariff table after tariff degrouping',
        columns => \@columns2,
    );

    $model->{ungrouped}{allTariffs},
      $model->{ungrouped}{allTariffsByEndUser},
      $model->{ungrouped}{volumeData}, $unitsInYear,
      $tariffTableNew;

}

sub groupVolumes {
    my ( $model, $volumeDataUngrouped, $allTariffsByEndUser,
        $nonExcludedComponents, $componentVolumeNameMap, )
      = @_;
    $model->{ungrouped}{volumeData} = $volumeDataUngrouped;
    my (
        @columnsRegrouped, @columnsAggregated,
        @columnsReordered, %volumeDataGrouped
    );
    foreach (@$nonExcludedComponents) {
        my @patches;
        if ( $_ eq 'Fixed charge p/MPAN/day' ) {
            my $relatedMPANrows = Labelset(
                name => 'Related MPANs',
                list => [
                    grep { /related MPAN/i; }
                      @{ $model->{ungrouped}{allTariffs}{list} }
                ]
            );
            push @patches,
              Constant(
                name          => 'Related MPANs not counted',
                defaultFormat => '0con',
                rows          => $relatedMPANrows,
                data => [ [ map { 0; } @{ $relatedMPANrows->{list} } ] ],
              );
        }
        push @columnsRegrouped,
          my $regrouped = Stack(
            name          => $componentVolumeNameMap->{$_},
            rows          => $model->{ungrouped}{allTariffsByGroup},
            defaultFormat => '0copy',
            sources       => [ @patches, $volumeDataUngrouped->{$_} ],
          );
        push @columnsAggregated,
          my $aggregated = GroupBy(
            name          => $componentVolumeNameMap->{$_},
            rows          => $model->{ungrouped}{allTariffGroups},
            source        => $regrouped,
            defaultFormat => '0soft',
          );
        push @columnsReordered,
          $volumeDataGrouped{$_} = Stack(
            name          => $componentVolumeNameMap->{$_},
            defaultFormat => '0copy',
            sources       => [$aggregated],
            rows          => $allTariffsByEndUser,
          );
    }
    Columnset(
        name    => 'Volumes to be aggregated by tariff group',
        columns => \@columnsRegrouped
    );
    Columnset(
        name    => 'Volumes aggregated by tariff group',
        columns => \@columnsAggregated
    );
    Columnset(
        name    => 'Aggregated volumes reordered by end user',
        columns => \@columnsReordered
    );
    \%volumeDataGrouped;
}

sub setUpGrouping {
    my ( $model, $componentMap,
        $allEndUsers, $allTariffsByEndUser, $allTariffs ) = @_;
    my %map;
    if ( $model->{tariffGrouping} =~ /total/i ) {
        my ( $targetlv, $targetlvsub, $targethv );
        foreach ( @{ $allEndUsers->{list} } ) {
            next
              if /gener/i
              || !$componentMap->{$_}{'Unit rates p/kWh'}
              || !$componentMap->{$_}{'Capacity charge p/kVA/day'};
            if    (/^(> )?HV/i)     { $targethv    ||= $_; }
            elsif (/^(> )?LV Sub/i) { $targetlvsub ||= $_; }
            elsif (/^(> )?LV/i)     { $targetlv    ||= $_; }
        }
        foreach ( @{ $allEndUsers->{list} } ) {
            next if /gener/i;
            my $target;
            if    (/^(> )?HV/i)     { $target = $targethv; }
            elsif (/^(> )?LV Sub/i) { $target = $targetlvsub; }
            elsif (/^(> )?LV/i)     { $target = $targetlv; }
            if ( $target && $target ne $_ ) {
                $map{$_} = $target;
                if ( ref $_ && ref $target ) {
                    my @targetLines = map { s/^> //; $_; } split /\n/, $target;
                    foreach ( @{ $_->{list} } ) {
                        my ($prefix) = /^(.+?: )/s;
                        $prefix ||= '';
                        $map{$_} = join "\n",
                          map { $prefix . $_; } @targetLines;
                    }
                }
            }
        }
    }
    elsif ( $model->{tariffGrouping} =~ /supercustomer/i ) {
        my ( $targetdomestic, $targetbusiness, $targetunmetered );
        foreach ( @{ $allEndUsers->{list} } ) {
            next
              if /gener/i
              || /^(> )?HV/i
              || /^(> )?LV Sub/i
              || $componentMap->{$_}{'Capacity charge p/kVA/day'}
              || !$componentMap->{$_}{'Unit rates p/kWh'};
            if    (/unmeter|ums/i) { $targetunmetered ||= $_; }
            elsif (/non.dom/i)     { $targetbusiness  ||= $_; }
            else                   { $targetdomestic  ||= $_; }
        }
        foreach ( @{ $allEndUsers->{list} } ) {
            next
              if /gener/i
              || /^(> )?HV/i
              || /^(> )?LV Sub/i
              || $componentMap->{$_}{'Capacity charge p/kVA/day'};
            my $target;
            if    (/unmeter|ums/i) { $target = $targetunmetered; }
            elsif (/non.dom/i)     { $target = $targetbusiness; }
            else                   { $target = $targetdomestic; }
            if ( $target && $target ne $_ ) {
                $map{$_} = "$target";
                if ( ref $_ && ref $target ) {
                    my @targetLines = map { s/^> //; $_; } split /\n/, $target;
                    foreach ( @{ $_->{list} } ) {
                        my ($prefix) = /^(.+?: )/s;
                        $prefix ||= '';
                        $map{$_} = join "\n",
                          map { $prefix . $_; } @targetLines;
                    }
                }
            }
        }
    }
    my @groups = map {
        my $g = $_;
        Labelset(
            name => $g,
            list => [
                grep { $g eq $_ || $map{$_} && $g eq $map{$_}; }
                  @{ $allTariffs->{list} }
            ]
        );
      } grep { !$map{$_}; }
      $allTariffsByEndUser->{groups}
      ? map { @{ $_->{list} }; } @{ $allTariffsByEndUser->{groups} }
      : @{ $allTariffsByEndUser->{list} };
    my $allTariffGroups =
      Labelset( name => 'All tariff groups', list => \@groups );
    my $allTariffsByGroup =
      Labelset( name => 'All tariffs by tariff group', groups => \@groups );
    $model->{ungrouped} = {
        allEndUsers         => $allEndUsers,
        allTariffsByEndUser => $allTariffsByEndUser,
        allTariffs          => $allTariffs,
        allTariffGroups     => $allTariffGroups,
        allTariffsByGroup   => $allTariffsByGroup,
    };
    my $retainedEndUsers = Labelset(
        name => 'End user groups',
        list => [ grep { !$map{$_}; } @{ $allEndUsers->{list} } ]
    );
    my $groupedTariffsByEndUser = Labelset(
        name   => 'All tariff groups by end user',
        groups => $retainedEndUsers->{list}
    );
    $retainedEndUsers, $groupedTariffsByEndUser, $allTariffGroups;
}

1;
