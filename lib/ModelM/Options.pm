﻿package ModelM;

# Copyright 2011 The Competitive Networks Association and others.
# Copyright 2012-2022 Franck Latrémolière and others.
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

sub allocationRules {

    my ($model) = @_;

    my $key = 'allocationRules?' . join '&',
      map { defined $model->{$_} ? "$_=$model->{$_}" : (); }
      qw(dcp094 dcp096 dcp097 dcp097A dcp117 dcp306 dcp395 fixedIndirectPercentage);

    return $model->{objects}{$key}{columns} if $model->{objects}{$key};

    my $expenditureSet = $model->{objects}{
        join '',
        'expenditureSet',
        $model->{dcp306} ? 306 : (),
        $model->{dcp395} ? 395 : ()
      }
      ||= Labelset(
        list => [
            split( /\n/, <<END_OF_LIST ),
Load related new connections & reinforcement (net of contributions)
Non-load new & replacement assets (net of contributions)
Non-operational capex
Faults
Inspections, & Maintenance
Tree Cutting
Network Policy
Network Design & Engineering
Project Management
Engineering Mgt & Clerical Support
Control Centre
System Mapping - Cartographical
Customer Call Centre
Stores
Vehicles & Transport
IT & Telecoms
Property Mgt
HR & Non-operational Training
Health & Safety & Operational Training
Finance & Regulation
CEO etc
Atypical cash costs
Pension deficit payments
Metering
Excluded services & de minimis
Relevant distributed generation (less contributions)
IFI
Disallowed Related Party Margins
Statutory Depreciation
Network Rates
Transmission Exit Charges
Pension deficit repair payments by related parties (note 2)
Non activity costs and reconciling amounts (note 3)
END_OF_LIST
            $model->{dcp306}
            ? 'Non activity costs and reconciling amounts - Ofgem licence fees'
            : (),
            $model->{dcp395}
            ? 'Pass-through Smart Meter Communication Licence Costs'
            : (),
        ]
      );

    my @rules = (
        $model->{dcp094}                ? 'Kill'            : $model->{dcp117}
          && $model->{dcp117} =~ /2014/ ? 'Do not allocate' : 'MEAV',
        ( map { 'MEAV' } 1 .. 11 ),
        $model->{dcp097A}
          || $model->{dcp097} ? ( 'LV services', 'MEAV', 'MEAV' )
        : ( map { 'MEAV' } 1 .. 3 ),
        $model->{dcp097A} ? ( map { 'LV services' } 1 .. 2 )
        : $model->{dcp097}
          || $model->{includeTelecoms} ? ( map { 'MEAV' } 1 .. 2 )
        : ( map { 'Do not allocate' } 1 .. 2 ),
        $model->{dcp097A} || $model->{dcp097}
        ? ( 'LV services', 'MEAV', 'LV services', 'LV services' )
        : ( map { 'MEAV' } 1 .. 4 ),
        ( map { 'Do not allocate' } 1 .. 9 ),
        $model->{dcp096} ? 'Deduct from revenue' : 'EHV only',
        ( map { 'Do not allocate' } 1 .. 2 ),
        $model->{dcp306} ? 'LV services' : (),
        $model->{dcp395} ? 'LV services' : (),
    );

    @rules[ 6 .. $#rules ] =
      map { $_ eq 'MEAV' ? '60%MEAV' : $_; } @rules[ 6 .. $#rules ]
      if $model->{fixedIndirectPercentage};

    my @c = (
        Constant(
            name  => 'Allocation key',
            lines => 'In a legacy Method M workbook, these data are on'
              . ' sheet Calc-Opex, possibly starting at cell K7.',
            data          => \@rules,
            defaultFormat => $model->{multiModelSharing}
            ? 'texthard'
            : 'textcon',
            rows => $expenditureSet,
        ),
        Constant(
            name  => 'Percentage capitalised',
            lines => 'In a legacy Method M workbook, these data are on'
              . ' sheet Calc-Opex, possibly starting at cell AL7.',
            data => [
                qw(1 1
                  .235 .235 .235 .235
                  .5257 .5257 .5257 .5257 .5257 .5257 .5257 .5257 .5257 .5257 .5257 .5257 .5257 .5257 .5257
                  0
                  .577
                  0 0 0 0 0 0 0 0 0 0), $model->{dcp306} ? 0 : (),
                $model->{dcp395} ? 0 : (),
            ],
            defaultFormat => $model->{multiModelSharing} ? '%hardnz'
            : '%connz',
            rows       => $expenditureSet,
            validation => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
        ),
        Constant(
            rows          => $expenditureSet,
            name          => 'Direct cost indicator',
            defaultFormat => $model->{multiModelSharing} ? '0hardnz'
            : '0connz',
            data => [
                [
                    ( map { 1 } 1 .. 6 ),    # direct cost categories
                    ( map { 0 } 1 .. 15 ),   # indirect cost categories
                    ( map { 1 } 1 .. 9 ),    # other costs weirdly marked direct
                    1,                       # transmission exit
                    ( map { 1 } 1 .. 2 ),    # other costs weirdly marked direct
                    $model->{dcp306} ? 0 : (),    # Ofgem licence fees
                    $model->{dcp395} ? 0 : (),    # Metering communications
                ]
            ]
        )
    );

    $model->{objects}{$key} = Columnset(
        name => 'Allocation rules',
        $model->{multiModelSharing}
          && !$model->{waterfalls} ? ( number => 1399 ) : (),
        columns  => \@c,
        location => 'Options'
    );

    \@c;

}

1;
