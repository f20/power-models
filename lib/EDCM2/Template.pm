package EDCM2;

=head Copyright licence and disclaimer

Copyright 2013 Franck Latrémolière, Reckon LLP and others.

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
require Spreadsheet::WriteExcel::Utility;

sub templates {
    my (
        $model,                            $tariffs,
        $importCapacity,                   $exportCapacityExempt,
        $exportCapacityChargeablePre2005,  $exportCapacityChargeable20052010,
        $exportCapacityChargeablePost2010, $tariffSoleUseMeav,
        $tariffLoc,                        $tariffCategory,
        $useProportions,                   $activeCoincidence,
        $reactiveCoincidence,              $indirectExposure,
        $nonChargeableCapacity,            $activeUnits,
        $creditableCapacity,               $tariffNetworkSupportFactor,
        $tariffDaysInYearNot,              $tariffHoursInRedNot,
        $previousChargeImport,             $previousChargeExport,
        $llfcImport,                       $llfcExport,
        $thisIsTheTariffTable,             $daysInYear,
        $hoursInRed,
    ) = @_;

    push @{ $model->{tablesTemplateImport} },
      $model->templateImport(
        $tariffs,        $llfcImport,          $thisIsTheTariffTable,
        $importCapacity, $activeCoincidence,   $daysInYear,
        $hoursInRed,     $tariffDaysInYearNot, $tariffHoursInRedNot,
      );

    push @{ $model->{tablesTemplateExport} }, $model->templateExport;
}

sub templateImport {

    my (
        $model,                $tariffs,        $llfcImport,
        $thisIsTheTariffTable, $importCapacity, $activeCoincidence,
        $daysInYear,           $hoursInRed,     $tariffDaysInYearNot,
        $tariffHoursInRedNot,
    ) = @_;

    my $index = Dataset(
        name          => 'Number',
        data          => [ [1] ],
        defaultFormat => 'thtarimport',
    );

    my @tariffComponents = map {
        Arithmetic(
            name          => $_->{name}->shortName,
            arguments     => { IV1 => $index, IV2_IV3 => $_ },
            arithmetic    => '=INDEX(IV2_IV3,IV1)',
            defaultFormat => $_->{name} =~ /k(?:VAr|W)h/
            ? '0.000copy'
            : '0.00copy',
        );
    } @{ $thisIsTheTariffTable->{columns} }[ 1 .. 4 ];

    my $agreedCapacity = Arithmetic(
        name          => 'Maximum import capacity (kVA)',
        arguments     => { IV1 => $index, IV2_IV3 => $importCapacity },
        arithmetic    => '=INDEX(IV2_IV3,IV1)',
        defaultFormat => '0hard',
    );

    my $exceededCapacity = Dataset(
        name          => 'Exceeded import capacity (kVA)',
        data          => [ [0] ],
        defaultFormat => '0hard',
    );

    foreach ( $activeCoincidence, $tariffDaysInYearNot, $tariffHoursInRedNot, )
    {
        my $df = $_->{defaultFormat} || '0.000soft';
        $df =~ s/copy|soft/hard/;
        $_ = Arithmetic(
            name          => $_->{name}->shortName,
            arguments     => { IV1 => $index, IV2_IV3 => $_ },
            arithmetic    => '=INDEX(IV2_IV3,IV1)',
            defaultFormat => $df,
        );
    }

    $_ = Stack( sources => [$_] ) foreach $daysInYear, $hoursInRed;

    my $units = Arithmetic(
        name          => 'Units consumed in super-red time band (kWh)',
        defaultFormat => '0soft',
        arithmetic    => '=IV1*IV2*(IV3-IV4)',
        arguments     => {
            IV1 => $agreedCapacity,
            IV2 => $activeCoincidence,
            IV3 => $hoursInRed,
            IV4 => $tariffHoursInRedNot,
        },
    );

    my $redPounds = Arithmetic(
        name          => 'Annual super-red charge (£)',
        defaultFormat => '0soft',
        arithmetic    => '=IV1*IV2/100',
        arguments     => { IV1 => $units, IV2 => $tariffComponents[0], },
    );

    my $fixedPounds = Arithmetic(
        name          => 'Annual fixed charge (£)',
        defaultFormat => '0soft',
        arithmetic    => '=(IV1-IV3)*IV2/100',
        arguments     => {
            IV1 => $daysInYear,
            IV2 => $tariffComponents[1],
            IV3 => $tariffDaysInYearNot,
        },
    );

    my $capacityPounds = Arithmetic(
        name          => 'Annual capacity charge (£)',
        defaultFormat => '0soft',
        arithmetic    => '=(IV1-IV3)*(IV5*IV6+IV7*IV8)/100',
        arguments     => {
            IV1 => $daysInYear,
            IV3 => $tariffDaysInYearNot,
            IV2 => $tariffComponents[2],
            IV7 => $tariffComponents[3],
            IV6 => $agreedCapacity,
            IV8 => $exceededCapacity,
        },
    );

    my $totalPounds = Arithmetic(
        name          => 'Total annual DUoS charge (£)',
        defaultFormat => '0soft',
        arithmetic    => '=IV1+IV2+IV3',
        arguments =>
          { IV1 => $redPounds, IV2 => $fixedPounds, IV3 => $capacityPounds, },
    );

    $_ = Arithmetic(
        name          => $_->{name}->shortName,
        arguments     => { IV1 => $index, IV2_IV3 => $_ },
        arithmetic    => '=INDEX(IV2_IV3,IV1)',
        defaultFormat => 'textcopy',
    ) foreach $tariffs, $llfcImport;

    my @psv;
    my $col = 1;
    foreach my $d ( @{ $model->{table935}{columns} } ) {
        if ( my $nc = $d->lastCol ) {
            push @psv, map {
                Arithmetic(
                    defaultFormat => 'codecopy',
                    name          => '',
                    arithmetic    => '="935|'
                      . ( $col + $_ )
                      . '|"&IV1&"|"&INDEX(IV3_IV4,IV2,'
                      . ( 1 + $_ ) . ')',
                    arguments => {
                        IV1     => $index,
                        IV2     => $index,
                        IV3_IV4 => $d
                    },
                  )
            } 0 .. $nc;
            $col += $nc + 1;
        }
        else {
            push @psv,
              Arithmetic(
                defaultFormat => 'codecopy',
                name          => '',
                arithmetic => '="935|' . $col . '|"&IV1&"|"&INDEX(IV3_IV4,IV2)',
                arguments  => {
                    IV1     => $index,
                    IV2     => $index,
                    IV3_IV4 => $d
                },
              ) unless $d->{name} =~ /export/i;
            ++$col;
        }
    }

    Notes(
        name  => 'ELECTRICITY DISTRIBUTION CHARGES INFORMATION FOR IMPORT',
        lines => <<'EOX'

This template is intended to illustrate the use of system charges that a distributor might levy on a
supplier under an EHV Distribution Charging Methodology (EDCM).

Charges between supplier and end customer are a bilateral contractual matter.  A supplier may apply
its own charges in addition to, or instead of, the charges that this template illustrates.

This template is for illustration only.  In case of conflict, the published statement of Distribution Use of System charges takes precedence. 
EOX
      ),

      map {
        my $t = ref $_ eq 'ARRAY' ? $_->[1] : $_;
        Columnset(
            noHeaders     => 1,
            noSpacing     => ref $_ ne 'ARRAY',
            name          => ref $_ eq 'ARRAY' ? $_->[0] : '',
            singleRowName => ref $t->{name}
            ? $t->{name}->shortName
            : $t->{name},
            columns => [$t],
            ref $_ eq 'ARRAY' && $_->[2] ? ( lines => $_->[2] ) : (),
        );
      }

      [ 'Tariff identification', $index, ], $tariffs, $llfcImport,

      [
        'Distribution Use of System (DUoS) tariff (excluding VAT)',
        $tariffComponents[0]
      ],
      @tariffComponents[ 1 .. $#tariffComponents ],

      [ 'Calendar and time band information', $daysInYear, ],
      $tariffDaysInYearNot,
      $hoursInRed, $tariffHoursInRedNot,

      [ 'Capacity and consumption', $agreedCapacity ],
      $exceededCapacity, $activeCoincidence, $units,

      [
        'Distribution Use of System (DUoS) charges (excluding VAT)', $redPounds
      ],
      $fixedPounds, $capacityPounds, $totalPounds,

      [
        'Disclosure of detailed data for advanced modelling',
        $psv[0],
        [
'The following advanced technical information is not necessary to understand your charges,',
            'but might be useful if you wish to conduct additional analysis.',
            'For further information about advanced modelling options, see:',
            'http://dcmf.co.uk/models/edcm.html'
        ]
      ],
      @psv[ 1 .. $#psv ];

}

sub templateExport {
    my ( $model, ) = @_;
    Notes(
        name  => 'ELECTRICITY DISTRIBUTION CHARGES INFORMATION FOR EXPORT',
        lines => <<'EOX');

This template is intended to illustrate the use of system charges that a distributor might levy on a
generator or supplier under an EHV Distribution Charging Methodology (EDCM).

Any charges between generator, supplier, customer and site owner are contractual matters.  They may
or may not reflect the charges that this template illustrates.

This template is for illustration only.  In case of conflict, the published statement of Distribution Use of System charges takes precedence.

This page is not done yet.
EOX
}

1;
