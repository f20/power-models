package Multiyear;

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
use Spreadsheet::WriteExcel::Utility;
use SpreadsheetModel::Shortcuts ':all';

sub setUpAssumptions {
    my ($me) = @_;
    $me->{percentageAssumptionColumns} = [];
    $me->{percentageAssumptionRowset}  = Labelset(
        list => [
            'MEAV change: 132kV',                                   #  1
            'MEAV change: 132kV/EHV',                               #  2
            'MEAV change: EHV',                                     #  3
            'MEAV change: EHV/HV',                                  #  4
            'MEAV change: 132kV/HV',                                #  5
            'MEAV change: HV network',                              #  6
            'MEAV change: HV service',                              #  7
            'MEAV change: HV/LV',                                   #  8
            'MEAV change: LV network',                              #  9
            'MEAV change: LV service',                              # 10
            'Cost change: direct costs',                            # 11
            'Cost change: indirect costs',                          # 12
            'Cost change: network rates',                           # 13
            'Cost change: transmission exit',                       # 14
            'Volume change: supercustomer metered demand units',    # 15
            'Volume change: supercustomer metered demand MPANs',    # 16
            'Volume change: site-specific metered demand units',    # 17
            'Volume change: site-specific metered demand MPANs',    # 18
            'Volume change: demand capacity',                       # 19
            'Volume change: demand excess reactive',                # 20
            'Volume change: unmetered demand units',                # 21
            'Volume change: generation units',                      # 22
            'Volume change: generation MPANs',                      # 23
            'Volume change: generation excess reactive',            # 24
        ]
    );
    $me->{overrideAssumptionColumns} = [];
    $me->{overrideAssumptionRowset} = Labelset( list => [ 'Days in year', ] );
}

sub registerModel {
    my ( $me, $model ) = @_;
    push @{ $me->{models} }, $model;
    my $assumptionZero;
    if ( ref $model->{dataset} eq 'HASH' ) {
        if ( my $sourceModel = $me->{modelByDataset}{ 0 + $model->{dataset} } )
        {
            $model->{sourceModel} = $sourceModel;
            $assumptionZero = 1;
        }
        elsif ( !$model->{sourceModel} ) {
            $me->{modelByDataset}{ 0 + $model->{dataset} } = $model;
            push @{ $me->{historical} }, $model;
            return;
        }
    }
    else {
        push @{ $me->{historical} }, $model;
        return;
    }
    push @{ $me->{scenario} }, $model;
    $me->setUpAssumptions unless $me->{percentageAssumptionColumns};
    push @{ $me->{percentageAssumptionColumns} },
      $me->{percentageAssumptionsByModel}{ 0 + $model } = Dataset(
        model         => $model,
        rows          => $me->{percentageAssumptionRowset},
        defaultFormat => '%hardpm',
        data          => [
            [
                $assumptionZero
                ? ( map { '' } @{ $me->{percentageAssumptionRowset}{list} } )
                : (
                    qw(0.02 0.02 0.02 0.02 0.02 0.02),    # MEAV EHV network
                    qw(0.02 0.02 0.02 0.02),    # MEAV HV/LV network/service
                    qw(0.02 0.02 0.02 0.02),    # direct etc.
                    qw(-0.01 0),                # super-customer
                    qw(0 0 0 0),                # site-specific
                    qw(0),                      # un-metered
                    qw(0.03 0.03 0.03),         # generation
                )
            ]
        ],
      );

    push @{ $me->{overrideAssumptionColumns} },
      $me->{overrideAssumptionsByModel}{ 0 + $model } = Dataset(
        model         => $model,
        rows          => $me->{overrideAssumptionRowset},
        defaultFormat => '0hard',
        data => [ [ map { '' } @{ $me->{overrideAssumptionRowset}{list} } ] ],
      );
}

sub percentageAssumptionsLocator {
    my ( $me, $model, $sourceModel ) = @_;
    my @assumptionsColumnLocationArray;
    sub {
        my ( $wb, $ws, $row ) = @_;
        if ( $row =~ /^[0-9]+$/s ) {
            --$row;
        }
        else {
            my $q = qr/$row/;
            ($row) =
              grep { $me->{percentageAssumptionRowset}{list}[$_] =~ /$q/; }
              0 .. $#{ $me->{percentageAssumptionRowset}{list} };
            die "$q not found in percentage assumptions table"
              unless defined $row;
        }
        unless (@assumptionsColumnLocationArray) {
            @assumptionsColumnLocationArray =
              $me->{percentageAssumptionsByModel}{ 0 + $model }
              ->wsWrite( $wb, $ws );
            $assumptionsColumnLocationArray[0] =
              q%'% . $assumptionsColumnLocationArray[0]->get_name . q%'!%;
        }
        $assumptionsColumnLocationArray[0]
          . xl_rowcol_to_cell(
            $assumptionsColumnLocationArray[1] + $row,
            $assumptionsColumnLocationArray[2],
            1, 1
          );
    };
}

sub generalOverride {
    my ( $me, $model, $wbook, $wsheet, $name ) = @_;
    my ($row) = grep { $me->{overrideAssumptionRowset}{list}[$_] =~ /$name/i; }
      0 .. $#{ $me->{overrideAssumptionRowset}{list} };
    die "$name not found in override assumptions table" unless defined $row;
    my ( $ws, $ro, $co ) =
      $me->{overrideAssumptionsByModel}{ 0 + $model }
      ->wsWrite( $wbook, $wsheet );
    q%'% . $ws->get_name . q%'!% . xl_rowcol_to_cell( $ro + $row, $co, 1, 1 );
}

1;
