package CDCM;

=head Copyright licence and disclaimer

Copyright 2009-2011 Energy Networks Association Limited and others.

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

sub setUp {

    my ($model) = @_;

    my $drmLevels = Labelset(
        name => 'Core network levels',
        list => [ split /\n/, <<EOL] );
132kV
132kV/EHV
EHV
EHV/HV
HV
HV/LV
LV circuits
EOL

    my $drmExitLevels = Labelset
      name => 'Core and transmission exit levels',
      list => [ 'GSPs', @{ $drmLevels->{list} } ];

    my $modelLife = Dataset(
        name       => 'Annualisation period (years)',
        validation => {
            validate => 'decimal',
            criteria => 'between',
            minimum  => 0,
            maximum  => 999_999,
        },
        data          => [40],
        defaultFormat => '0hard'

          #    rows => $networkComponents,
          #    data => [ map { 40 } @{ $networkComponents->{list} } ]
    );

    my $rateOfReturn = Dataset(
        name       => 'Rate of return',
        validation => {
            validate      => 'decimal',
            criteria      => 'between',
            minimum       => 0,
            maximum       => 4,
            input_title   => 'Rate of return:',
            input_message => 'Percentage',
            error_message => 'The rate of return must be'
              . ' a non-negative percentage value.'
        },
        defaultFormat => '%hard',
        data          => [0.069]
    );

    my $daysInYear = Dataset(
        name       => 'Days in the charging year',
        validation => {
            validate => 'decimal',
            criteria => 'between',
            minimum  => 365,
            maximum  => 366,
        },
        data          => [365],
        defaultFormat => '0hard'
    );

    $model->{edcmTables}[0][1] = Stack( sources => [$daysInYear] )
      if $model->{edcmTables};

    my ($daysInBefore);
    my $daysInAfter = $daysInYear;

    if ( $model->{inYear} ) {
        if ( $model->{inYear} =~ /after/i ) {
            $daysInAfter = Arithmetic(
                name =>
                  'Number of days in period for which new tariffs are to apply',
                validation => {
                    validate => 'decimal',
                    criteria => 'between',
                    minimum  => 1,
                    maximum  => 999,
                },
                data          => [0],
                defaultFormat => '0hard',
            );
        }
        elsif ( $model->{inYear} =~ /twice/i ) {
            $daysInBefore = [
                Dataset(
                    name => 'Number of days in period covered'
                      . ' by tables 1097/1098 (if any)',
                    validation => {
                        validate => 'decimal',
                        criteria => 'between',
                        minimum  => 0,
                        maximum  => 999,
                    },
                    data          => [0],
                    defaultFormat => '0hard',
                ),
                Dataset(
                    name => 'Number of days in period covered'
                      . ' by tables 1095/1096 (if any)',
                    validation => {
                        validate => 'decimal',
                        criteria => 'between',
                        minimum  => 0,
                        maximum  => 999,
                    },
                    data          => [0],
                    defaultFormat => '0hard',
                ),
            ];
            $daysInAfter = Arithmetic(
                name =>
                  'Number of days in period for which new tariffs are to apply',
                arithmetic => '=A1-A2-A3',
                arguments  => {
                    A1 => $daysInYear,
                    A2 => $daysInBefore->[0],
                    A3 => $daysInBefore->[1],
                },
                defaultFormat => '0softnz',
            );
        }
        else {
            $daysInBefore = Dataset(
                name => 'Number of days in the charging year'
                  . ' before the tariff change (if any)',
                validation => {
                    validate => 'decimal',
                    criteria => 'between',
                    minimum  => 0,
                    maximum  => 999,
                },
                data          => [0],
                defaultFormat => '0hard'
            );
            $daysInAfter = Arithmetic(
                name =>
                  'Number of days in period for which new tariffs are to apply',
                arithmetic    => '=A1-A2',
                arguments     => { A1 => $daysInYear, A2 => $daysInBefore, },
                defaultFormat => '0softnz',
            );
        }
    }

    my $annuityRate = Arithmetic(
        name          => 'Annuity rate',
        defaultFormat => '%softnz',
        arithmetic    => '=PMT(A1,A2,-1)*IF(OR(A3>366,A4<365),A5/365.25,1)',
        arguments     => {
            A1 => $rateOfReturn,
            A2 => $modelLife,
            A3 => $daysInYear,
            A4 => $daysInYear,
            A5 => $daysInYear,
        }
    );

    if ( $model->{networkRates} ) {
        my $ratesMultiplier = Dataset(
            name       => 'Network rates multiplier',
            validation => {
                validate      => 'decimal',
                criteria      => 'between',
                minimum       => 0,
                maximum       => 4,
                error_message => 'Must be'
                  . ' a non-negative percentage value.',
            },
            defaultFormat => '%hard',
            data          => [0.485]
        );
        $annuityRate = Arithmetic(
            name          => 'Annuity rate',
            defaultFormat => '%softnz',
            arithmetic    => '=(PMT(A1,A2,-1)-A4/A3)/(1-A5)',
            arguments     => {
                A1 => $rateOfReturn,
                A2 => $modelLife,
                A3 => $modelLife,
                A4 => $ratesMultiplier,
                A5 => $ratesMultiplier
            }
        );
        push @{ $model->{optionLines} }, 'Network rates explicitly included';
    }

    my $modelPowerFactor = Dataset(
        name => Label(
            'Power factor', 'Power factor for all flows in the network model'
        ),
        validation => {
            validate => 'decimal',
            criteria => 'between',
            minimum  => 0.001,
            maximum  => 1,
        },
        data => [ $model->{powerFactor} || 0.95 ]
    );

    $daysInYear, $daysInBefore, $daysInAfter, $modelLife, $annuityRate,
      $modelPowerFactor, $drmLevels, $drmExitLevels;

}

1;
