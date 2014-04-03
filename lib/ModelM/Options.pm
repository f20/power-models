package ModelM;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others.

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

sub allocationRules {

    my ($model) = @_;

    my $expenditureSet = Labelset( list => [ split /\n/, <<END_OF_LIST] );
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

    my @rules = (
        $model->{dcp094} ? 'Kill' : 'MEAV',
        ( map { 'MEAV' } 1 .. 11 ),
        $model->{dcp097A} || $model->{dcp097} ? ( 'LV only', 'MEAV', 'MEAV' )
        : ( map { 'MEAV' } 1 .. 3 ),
        $model->{dcp097A} ? ( map { 'LV only' } 1 .. 2 )
        : $model->{dcp097}
          || $model->{includeTelecoms} ? ( map { 'MEAV' } 1 .. 2 )
        : ( map { 'Do not allocate' } 1 .. 2 ),
        $model->{dcp097A}
          || $model->{dcp097} ? ( 'LV only', 'MEAV', 'LV only', 'LV only' )
        : ( map { 'MEAV' } 1 .. 4 ),
        ( map   { 'Do not allocate' } 1 .. 9 ),
        $model->{dcp096} ? 'Deduct from revenue' : 'EHV only',
        ( map { 'Do not allocate' } 1 .. 2 ),
    );

    @rules[ 6 .. $#rules ] =
      map { $_ eq 'MEAV' ? '60%MEAV' : $_; } @rules[ 6 .. $#rules ]
      if $model->{fixedIndirectPercentage};

    my @c = (
        Constant(
            name          => 'Allocation key',
            lines         => 'From sheet Opex Allocation, starting at cell AJ6',
            data          => \@rules,
            defaultFormat => $model->{multiModelSharing}
            ? 'texthard'
            : 'textcon',
            rows => $expenditureSet,
        ),
        Constant(
            name  => 'Percentage capitalised',
            lines => 'From sheet Opex Allocation, starting at cell AJ6',
            data  => [
                qw(1 1
                  .235 .235 .235 .235
                  .5257 .5257 .5257 .5257 .5257 .5257 .5257 .5257 .5257 .5257 .5257 .5257 .5257 .5257 .5257
                  0
                  .577
                  0 0 0 0 0 0 0 0 0 0)
            ],
            defaultFormat => $model->{multiModelSharing} ? '%hardnz' : '%connz',
            rows          => $expenditureSet,
            validation    => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
        ),
        Constant(
            rows          => $expenditureSet,
            name          => 'Direct cost indicator',
            defaultFormat => $model->{multiModelSharing} ? '0hardnz' : '0connz',
            data          => [
                [
                    ( map { 1 } 1 .. 6 ),
                    ( map { 0 } 1 .. 15 ),
                    ( map { 1 } 1 .. 12 )
                ]
            ]
        )
    );

    Columnset(
        name     => 'Allocation rules',
        columns  => \@c,
        location => 'Options'
    );

    \@c;

}

1;
