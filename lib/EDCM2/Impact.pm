package EDCM2;

=head Copyright licence and disclaimer

Copyright 201-2014 Franck Latrémolière, Reckon LLP and others.

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
use SpreadsheetModel::Shortcuts ':all';

sub mangleLoadFlowInputs {
    my ( $model, @inputs, ) = @_;

=head Columns

0	locations
1	locParent
2	c1
3	a1d
4	r1d
5	a1g
6	r1g

=cut

    return @inputs unless $model->{numExtraLocations};
    return @inputs if !$model->{method} || $model->{method} =~ /none/i;

    my $existingLocs = $inputs[0]{rows};
    my @newLocs =
      $model->{numExtraLocations}
      ? map { 9999 - $model->{numExtraLocations} + $_ }
      1 .. $model->{numExtraLocations}
      : ();
    my $data = [ map { '' } @newLocs ];
    my $newLocs = Labelset(
        list          => \@newLocs,
        defaultFormat => 'thloc',
    );
    my $allLocs = Labelset(
        defaultFormat => 'thloc',
        list          => [ @{ $existingLocs->{list} }, @newLocs, ]
    );

    my @newcol;
    my @allcol;
    my $doer = sub {
        my $df1 = my $df2 = $_[0]->{defaultFormat} || '0.000hard';
        $df1 =~ s/hard/input/ if $model->{lockedInputs};
        $df2 =~ s/hard/copy/;
        my $n = SpreadsheetModel::Object::_shortName( $_[0]{name} );
        my @new;
        push @newcol,
          @new = Dataset(
            cols          => $_[0]{cols},
            name          => $n,
            rows          => $newLocs,
            data          => $data,
            dataset       => $model->{dataset},
            defaultFormat => $df1,
          ) unless ref $_[0] eq 'SpreadsheetModel::Constant';
        push @allcol,
          my $result = Stack(
            name          => $n,
            defaultFormat => $df2,
            rows          => $allLocs,
            cols          => $_[0]{cols},
            sources       => [ $_[0], @new, ]
          );
        $result;
    };
    foreach (@inputs) {
        if ( ref $_ eq 'ARRAY' ) { $_ = $doer->($_) foreach @$_; }
        elsif ($_) { $_ = $doer->($_); }
    }

    push @{ $model->{impactInputTables} },
      Columnset(
        name     => 'Additional locations',
        location => 'Impact',
        number   => $model->{method} =~ /LRIC/i ? 914 : 912,
        columns  => \@newcol,
      );

    Columnset(
        name     => 'Combined location data',
        location => $model->{method} =~ /LRIC/i ? '913' : '911',
        number   => $model->{method} =~ /LRIC/i ? 923 : 921,
        columns  => \@allcol,
    );

    @inputs;

}

sub mangleTariffInputs {
    my ( $model, @columns ) = @_;

=head Columns (not compatible with DCP 189)

0	tariffs
1	importCapacity
2	exportCapacityExempt
3	exportCapacityChargeablePre2005
4	exportCapacityChargeable20052010
5	exportCapacityChargeablePost2010
6	tariffSoleUseMeav
7	tariffLoc
8	tariffCategory
9	useProportions
10	activeCoincidence
11	reactiveCoincidence
12	indirectExposure
13	nonChargeableCapacity
14	activeUnits
15	creditableCapacity
16	tariffNetworkSupportFactor
17	tariffDaysInYearNot
18	tariffHoursInPurpleNot
19	previousChargeImport
20	previousChargeExport
21	llfcImport
22	llfcExport

=cut

    $model->{numExtraTariffs} = 2
      unless ( $model->{numTariffs} || 0 ) * 2 +
      ( $model->{numExtraTariffs}  || 0 ) +
      ( $model->{numSampleTariffs} || 0 ) > 1;
    my $existingTariffs   = $columns[0]{rows};
    my @additionalTariffs = (
        $model->{numExtraTariffs} ?
          map { chr( 64 + int( $_ / 26 ) ) . chr( 65 + $_ % 26 ) . ' [ADDED]'; }
          0 .. ( $model->{numExtraLocations} - 1 )
        : (),
        $model->{numSampleTariffs} ?
          map { chr( 64 + int( $_ / 26 ) ) . chr( 65 + $_ % 26 ); }
          ( $model->{numExtraTariffs} || 0 ) .. (
            ( $model->{numExtraTariffs} || 0 ) + $model->{numSampleTariffs} - 1
          )
        : ()
    );
    my $newTariffs = Labelset(
        list => [
            ( map { "Amended $_" } @{ $existingTariffs->{list} } ),
            @additionalTariffs,
        ]
    );
    $model->{tariffSet} = Labelset(
        defaultFormat => 'thtar',
        list          => [
            ( map { ( $_, "Amended $_" ); } @{ $existingTariffs->{list} } ),
            @additionalTariffs,
        ]
    );
    my $defaultingInputMaker = sub {
        new SpreadsheetModel::Custom(
            @_,
            rows      => $newTariffs,
            custom    => ['=IV1'],
            wsPrepare => sub {
                my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) =
                  @_;
                sub {
                    my ( $x, $y ) = @_;
                    return $self->{defaultValue} || '', $format
                      if $y > $#{ $existingTariffs->{list} };
                    '', $format, $formula->[0], map {
                        $_ =>
                          Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                            $rowh->{$_} + $y,
                            $colh->{$_} + $x,
                            0, 0
                          )
                    } @$pha;
                };
            },
        );
    };

    $columns[0] = Stack(
        rows          => $model->{tariffSet},
        name          => 'Tariff name',
        defaultFormat => 'textcopy',
        sources       => [
            Arithmetic(
                name          => 'Names of original tariffs',
                defaultFormat => 'textsoft',
                arithmetic    => '=IV1&" (original)"',
                arguments     => { IV1 => $columns[0] },
                number        => 939,
            ),
            $model->{amendedTariffset} = $defaultingInputMaker->(
                name          => 'Name of amended or new tariff',
                defaultValue  => ' ',
                defaultFormat => $model->{lockedInputs}
                ? 'textinput'
                : 'texthard',
                arguments => { IV1 => $columns[0] },
            ),
        ]
    );

    foreach ( @columns[ 1 .. 18 ] ) {
        my $df1 = my $df2 = $_->{defaultFormat} || '0.000hard';
        $df1 =~ s/hard/input/ if $model->{lockedInputs};
        $df2 =~ s/hard/copy/;
        my $n = SpreadsheetModel::Object::_shortName( $_->{name} );
        $_ = Stack(
            name          => $n,
            defaultFormat => $df2,
            rows          => $model->{tariffSet},
            cols          => $_->{cols},
            sources       => [
                $_,
                $defaultingInputMaker->(
                    cols => $_->{cols},
                    name => ( $n =~ /divided by/i ? 'DNO assumption for ' : '' )
                      . $n,
                    $df1 =~ /^text/             ? ( defaultValue => ' ' )
                    : $n =~ /capacity \(kVA\)$/ ? ( defaultValue => 'VOID' )
                    : (),
                    arguments     => { IV1 => $_ },
                    defaultFormat => $df1,
                ),
            ]
        );
    }

    foreach ( @columns[ 19 .. 22 ] ) {
        $_ = Stack(
            name    => SpreadsheetModel::Object::_shortName( $_->{name} ),
            rows    => $model->{tariffSet},
            cols    => $_->{cols},
            sources => [$_]
        );
    }

    my $newTariffsActualRedDemand = $defaultingInputMaker->(
        name          => "Actual $model->{timebandName} consumption (kW/kVA)",
        arguments     => { IV1 => $columns[10]{sources}[0] },
        defaultFormat => $model->{lockedInputs}
        ? '0.000input'
        : '0.000hard',
    );

    push @{ $model->{impactInputTables} },
      Columnset(
        name    => 'Tariff scenarios: import and export capacities',
        number  => 936,
        columns => [
            $model->{amendedTariffset},
            map { $_->{sources}[1] } @columns[ 1 .. 5, 17, 18 ]
        ],
      ),
      Columnset(
        name    => 'Tariff scenarios: customer and usage',
        number  => 937,
        columns => [
            Stack(
                sources       => [ $model->{amendedTariffset} ],
                defaultFormat => 'textcopy',
            ),
            $newTariffsActualRedDemand,
            ( map { $_->{sources}[1] } @columns[ 10, 11, 13, 12, 16, 15, 14 ] ),
        ],
      ),
      Columnset(
        name    => 'Tariff scenarios: distribution network data',
        number  => 938,
        columns => [
            ( map { $_->{sources}[1] } @columns[ 7, 6, 8, 9 ] ),
            Stack(
                sources       => [ $model->{amendedTariffset} ],
                defaultFormat => 'textcopy',
            ),
        ],
      );

    push @columns,
      Stack(
        name    => "Actual $model->{timebandName} demand (kW per kVA)",
        rows    => $model->{tariffSet},
        sources => [ $columns[10]{sources}[0], $newTariffsActualRedDemand ]
      );

    Columnset(
        name     => 'Combined tariff input data',
        location => '935',
        number   => 940,
        columns  => \@columns,
    );

    @columns,
      Constant(
        name => 'Weighting of each tariff for reconciliation of totals',
        rows => $model->{tariffSet},
        data => [
            ( map { ( -1, 1 ); } @{ $existingTariffs->{list} } ),
            ( map { /^New/i ? 1 : 0; } @{ $newTariffs->{list} } ),
        ],
      );

}

sub impactFinancialSummary {
    my ( $model, $tariffs, $tariffColumns, $actualRedDemandRate,
        $revenueBitsDref, @revenueBitsG )
      = @_;

    my @revenueBitsD = (
        Stack(
            sources       => [ $revenueBitsDref->[0] ],
            defaultFormat => '0copy',
        ),
        Arithmetic(
            name => "$model->{TimebandName} charge for demand (£/year)",
            defaultFormat => '0softnz',
            arithmetic    => '=0.01*(IV9-IV7)*IV1*IV6*IV8',
            arguments     => {
                IV8 => $actualRedDemandRate,
                map { $_ => $revenueBitsDref->[1]{arguments}{$_} }
                  qw(IV1 IV6 IV7 IV9),
            }
        ),
        Stack(
            sources       => [ $revenueBitsDref->[2] ],
            defaultFormat => '0copy',
        ),
    );
    my $rev2d = Arithmetic(
        name          => 'Total for demand (£/year)',
        defaultFormat => '0soft',
        arithmetic    => '=' . join( '+', map { "IV$_" } 1 .. @revenueBitsD ),
        arguments =>
          { map { ( "IV$_" => $revenueBitsD[ $_ - 1 ] ) } 1 .. @revenueBitsD },
    );

    my @suminfocols = map {
        Stack(
            sources => [$_],
            ref $_->{defaultFormat} ? ( defaultFormat => '0copy' ) : (),
          )
    } grep { $_ } @{ $model->{summaryInformationColumns} };
    push @suminfocols, Arithmetic(
        name => "Difference due to $model->{timebandName} consumption"
          . ' and rounding errors (£/year)',
        defaultFormat => '0softnz',
        arithmetic    => join( '',
            '=IV1',
            map { $suminfocols[$_] ? ( "-IV" . ( 30 + $_ ) ) : () }
              0 .. $#suminfocols ),
        arguments => {
            IV1 => $rev2d,
            map {
                $suminfocols[$_]
                  ? ( "IV" . ( 30 + $_ ), $suminfocols[$_] )
                  : ()
            } 0 .. $#suminfocols
        }
    );

    Columnset(
        name     => 'Tariffs for demand',
        location => 'Impact',
        number   => 4991,
        columns =>
          [ map { Stack( sources => [$_] ) } @{$tariffColumns}[ 0 .. 4 ] ],
      ),
      Columnset(
        name     => 'Financial summary for demand',
        location => 'Impact',
        number   => 4992,
        columns  => [ Stack( sources => [$tariffs] ), @revenueBitsD, $rev2d, ]
      ),
      Columnset(
        name     => 'Analysis of the model\'s estimate of demand charges',
        location => 'Impact',
        number   => 4993,
        columns  => [ Stack( sources => [$tariffs] ), @suminfocols, ]
      ),
      Columnset(
        name     => 'Tariffs for generation',
        location => 'Impact',
        number   => 4996,
        columns =>
          [ map { Stack( sources => [$_] ) } @{$tariffColumns}[ 0, 5 .. 8 ] ],
      ),
      Columnset(
        name     => 'Financial summary for generation',
        location => 'Impact',
        number   => 4997,
        columns  => [
            Stack( sources => [$tariffs] ),
            map {
                Stack(
                    sources => [$_],
                    ref $_->{defaultFormat} ? ( defaultFormat => '0copy' ) : (),
                  )
            } @revenueBitsG,
        ]
      );
}

sub sheetPriority {
    $_[1] =~ /^(?:Impact|Overview|Index)$/si ? 1 : 0;
}

sub impactNotes {
    my ($model) = @_;
    push @{ $model->{impactInputTables} }, Notes( lines => <<'EOL'),
EDCM illustrative scenario analysis tool

This tool is intended to illustrate the use of system charges that a distributor might levy on a supplier under an EHV
Distribution Charging Methodology (EDCM).

Charges between supplier and customer are a bilateral contractual matter.  A supplier may apply its own charges in addition to,
or instead of, the charges that this tool tries to illustrate. Furthermore, different charges might be applied by distributors
under the CDCM, under derogations, for out-of-area or IDNO connections, or for services other than use of system.

This workbook is populated with data for one year and does not provide any forecasts.  Please speak to your distributor or
supplier if you would like to discuss possible future charges.

UNDER NO CIRCUMSTANCES SHOULD YOU HOLD ANY SUPPLIER OR ANY DISTRIBUTOR RESPONSIBLE FOR ANY DATA OR RESULTS IN THIS TOOL. CHARGES
MAY VARY.  THIS TOOL IS NOT DESIGNED FOR FORECASTING. WHILST THE AUTHORS HAVE TRIED TO ENSURE ACCURACY, THEY MAKE NO WARRANTY AS
TO THE CORRECTNESS, CURRENCY, TIMELINESS, QUALITY OR COMPLETENESS OF THIS TOOL OR ITS RELEVANCE TO YOUR CIRCUMSTANCES. FOR
INFORMATION ABOUT APPLICABLE CHARGES, CONTACT YOUR ACCOUNT MANAGER.

THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED
TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL AUTHORS OR
CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED
AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN
ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

This tool has been developed by the DCMF MIG, a subgroup of the electricity industry's Distribution Charging Methodologies
Forum, and by Reckon LLP, an advisory firm.  It builds on the charging models developed by distributors through the Energy
Networks Association. Reckon LLP is grateful for the assistance that distributors and suppliers have provided in the development
of this tool.  Mistakes are the sole responsibility of Reckon LLP. Reckon LLP does not approve of the EDCM charging
methodologies illustrated by this tool.

For any comments about this tool, contact Franck Latrémolière through http://dcmf.co.uk/franck or f20@reckon.co.uk. 

EOL
}

1;
