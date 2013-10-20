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
      ? map { 4000 + $_ } 1 .. $model->{numExtraLocations}
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
        push @newcol, @new = Dataset(
            cols          => $_[0]{cols},
            name          => $n,
            rows          => $newLocs,
            data          => $data,
            defaultFormat => $df1,
        ) unless ref $_[0] eq 'SpreadsheetModel::Constant';
        push @allcol, my $result = Stack(
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

=head Columns

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
18	tariffHoursInRedNot
19	previousChargeImport
20	previousChargeExport
21	llfcImport
22	llfcExport

=cut

    my $existingTariffs = $columns[0]{rows};
    my @additionalTariffs =
      $model->{numExtraTariffs}
      ? map { 'New tariff ' . ( 500 + $_ ) } 1 .. $model->{numExtraTariffs}
      : ();
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
                    $df1 =~ /^text/ ? ( defaultValue => ' ' ) : (),
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
        name          => 'Actual super-red consumption (kW/kVA)',
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
            $newTariffsActualRedDemand,
            map { $_->{sources}[1] } @columns[ 10, 11, 13, 12, 16, 15, 14 ]
        ],
      ),
      Columnset(
        name    => 'Tariff scenarios: distribution network data',
        number  => 938,
        columns => [ map { $_->{sources}[1] } @columns[ 7, 6, 8, 9 ] ],
      );

    push @columns,
      Stack(
        name    => 'Actual super-red demand (kW per kVA)',
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
        data => [ map { ( -1, 1 ) } @{ $existingTariffs->{list} } ],
      );

}

sub impactFinancialSummary {
    my ( $model, $tariffs, $thisIsTheTariffTable, $actualRedDemandRate,
        $revenueBitsDref, @revenueBitsG )
      = @_;

    my @revenueBitsD = (
        Stack(
            sources       => [ $revenueBitsDref->[0] ],
            defaultFormat => '0copy',
        ),
        Arithmetic(
            name          => 'Super red charge for demand (£/year)',
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
        name          => 'Difference (£/year)',
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
        columns  => [
            map { Stack( sources => [$_] ) }
              @{ $thisIsTheTariffTable->{columns} }[ 0 .. 4 ]
        ],
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
        columns  => [
            map { Stack( sources => [$_] ) }
              @{ $thisIsTheTariffTable->{columns} }[ 0, 5 .. 8 ]
        ],
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

sub frontSheets {
    qw(Impact Overview Index);
}

sub impactNotes {
    my ($model) = @_;
    push @{ $model->{impactInputTables} }, Notes( lines => <<'EOL'),
EDCM non-forecasting scenario analysis tool


This workbook has been developed by the DCMF MIG, a subgroup of the electricity industry's
Distribution Charging Methodologies Forum, and by Reckon LLP, an independent advisory firm, as a
tool to illustrate an electricity distribution use of system charging methodology.  It builds on a
charging model developed by distributors through the Energy Networks Association.

This tool is intended to illustrate the use of system charges that a distributor would levy on a
supplier in cases where the relevant charging methodology applies.  Additional or different charges
might be applied under derogations from the charging methodology, in the case of out-of-area or IDNO
connections, or for services other than use of system.

CHARGES MAY VARY.  THIS TOOL IS FOR SCENARIO ANALYSIS ONLY.  IT IS NOT DESIGNED FOR FORECASTING.

The charges that a supplier levies on its customer are a private contractual matter between
them.  If so agreed, it is legitimate for the supplier to apply its own charges in addition to,
or instead of, the charges illustrated in this tool.

Reckon LLP is grateful for the assistance that distributors and suppliers have provided in the
development of this tool, but mistakes are the sole responsibility of Reckon LLP.

IN NO CIRCUMSTANCES SHOULD YOU HOLD ANY SUPPLIER OR ANY DISTRIBUTOR RESPONSIBLE
IN ANY WAY FOR ANY DATA OR RESULTS IN THIS TOOL.  FOR INFORMATION ABOUT THE
CHARGES APPLICABLE TO YOU, CONTACT YOUR ACCOUNT MANAGER.

This workbook is populated with data for one particular year.  Forecasting
future inputs into this charging methodology is difficult.  Please speak to your
distributor or supplier if you would like to discuss possible future charges.

Whilst Reckon LLP believes the information to be correct when produced, it makes
no warranty as to the correctness, currency, timeliness, quality, content or
completeness of the information, or its relevance to your circumstances.

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

This tool illustrates an EHV Distribution Charging Methodology (EDCM) which was developed
pursued to licence conditions imposed on distributors by Ofgem (a UK Government department,
www.ofgem.gov.uk). The charging methodology was approved by Ofgem, changes to it require Ofgem's
approval, and the arrangements for the governance of the charging methodology are under Ofgem's
control. Reckon LLP does not approve of this charging methodology or of Ofgem's conduct in
relation to the development or governance of the methodology. If you have any comments or
complaints about the charging methodology or its governance, please contact Ofgem.  As of
October 2013 the relevant officials at Ofgem seemed to be:
* Hannah Nixon, Senior Partner, hannah.nixon@ofgem.gov.uk, 020 7901 7165.
* Andy Burgess, Associate Partner, Andy.Burgess@ofgem.gov.uk, 020 7901 7159.
* Iain Morgan, Senior Regulatory Economist, Iain.Morgan@ofgem.gov.uk, 020 7901 7030.

For comments on the user interface of this tool, or if you experience any practical problem in using
this tool, please contact Franck Latrémolière at Reckon LLP, f20@reckon.co.uk or 020 7841 5858.

EOL
}
