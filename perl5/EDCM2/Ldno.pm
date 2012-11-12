package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2012 Energy Networks Association Limited and contributors. All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted
provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of
conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of
conditions and the following disclaimer in the documentation and/or other materials provided
with the distribution.

THIS SOFTWARE IS PROVIDED BY ENERGY NETWORKS ASSOCIATION LIMITED AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL ENERGY
NETWORKS ASSOCIATION LIMITED OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF
THE POSSIBILITY OF SUCH DAMAGE.

=cut

=head Table numbers used in this file

1181
1182
1183

=cut

use warnings;
use strict;
use utf8;

use SpreadsheetModel::Shortcuts ':all';

sub ldnoRev {
    my ($model) = @_;

    my $endUsers = Labelset( list => [ split /\n/, <<EOL] );
Domestic Unrestricted
Domestic Two Rate
Domestic Off Peak (related MPAN)
Small Non Domestic Unrestricted
Small Non Domestic Two Rate
Small Non Domestic Off Peak (related MPAN)
LV Medium Non-Domestic
LV Sub Medium Non-Domestic
HV Medium Non-Domestic
LV HH Metered
LV Sub HH Metered
HV HH Metered
NHH UMS
LV UMS (Pseudo HH Metered)
LV Generation NHH
LV Sub Generation NHH
LV Generation Intermittent
LV Generation Non-Intermittent
LV Sub Generation Intermittent
LV Sub Generation Non-Intermittent
HV Generation Intermittent
HV Generation Non-Intermittent
EOL

    my @tariffComponentMatrix = split /\n/, <<EOL;
ynnynn
yynynn
ynnnnn
ynnynn
yynynn
ynnnnn
yynynn
yynynn
yynynn
yyyyyy
yyyyyy
yyyyyy
ynnnnn
yyynnn
ynnynn
ynnynn
ynnyny
yyyyny
ynnyny
yyyyny
ynnyny
yyyyny
EOL

    if ( $model->{DCP130} ) {
        $endUsers = Labelset( list => [ split /\n/, <<EOL] );
Domestic Unrestricted
Domestic Two Rate
Domestic Off Peak (related MPAN)
Small Non Domestic Unrestricted
Small Non Domestic Two Rate
Small Non Domestic Off Peak (related MPAN)
LV Medium Non-Domestic
LV Sub Medium Non-Domestic
HV Medium Non-Domestic
LV HH Metered
LV Sub HH Metered
HV HH Metered
NHH UMS category A
NHH UMS category B
NHH UMS category C
NHH UMS category D
LV UMS (Pseudo HH Metered)
LV Generation NHH
LV Sub Generation NHH
LV Generation Intermittent
LV Generation Non-Intermittent
LV Sub Generation Intermittent
LV Sub Generation Non-Intermittent
HV Generation Intermittent
HV Generation Non-Intermittent
EOL

        @tariffComponentMatrix = split /\n/, <<EOL;
ynnynn
yynynn
ynnnnn
ynnynn
yynynn
ynnnnn
yynynn
yynynn
yynynn
yyyyyy
yyyyyy
yyyyyy
ynnnnn
ynnnnn
ynnnnn
ynnnnn
yyynnn
ynnynn
ynnynn
ynnyny
yyyyny
ynnyny
yyyyny
ynnyny
yyyyny
EOL
    }

    my $ldnoLevels = $model->{ldnoRev} =~ /5/
      ? Labelset( list => [ split /\n/, <<EOL] )
Boundary 0000
Boundary 132kV
Boundary 132kV/EHV
Boundary EHV
Boundary HVplus
EOL
      : Labelset( list => [ split /\n/, <<EOL] );
Boundary 0000
Boundary 1000
Boundary 1100
Boundary 0100
Boundary 1110
Boundary 0110
Boundary 0010
Boundary 0001
Boundary 0002
Boundary 1001
Boundary 0011
Boundary 0111
Boundary 0101
Boundary 1101
Boundary 1111
EOL

    my $cdcmLevels = Labelset( list => [ split /\n/, <<EOL] );
LV demand
LV Sub demand or LV generation
HV demand or LV Sub generation
HV generation
EOL

    my @tariffComponents = split /\n/, <<EOL;
Unit rate 1 p/kWh
Unit rate 2 p/kWh
Unit rate 3 p/kWh
Fixed charge p/MPAN/day
Capacity charge p/kVA/day
Reactive power charge p/kVArh
EOL

    my $allTariffsByBoundaryLevelNotUsed = Labelset(
        groups => [
            map {
                my $b = $_;
                $b =~ s/\s*boundary\s*//i;
                Labelset(
                    name => "LDNO $b tariffs",
                    list => [ map { "LDNO $b: $_" } @{ $endUsers->{list} } ]
                  )
            } @{ $ldnoLevels->{list} }
        ]
    );

    my $allTariffsByEndUser = Labelset(
        groups => [
            map {
                my $e = $_;
                Labelset(
                    name => "> $e",
                    list => [
                        map {
                            local $_ = $_;
                            s/\s*boundary\s*//i;
                            "LDNO $_: $e"
                        } @{ $ldnoLevels->{list} }
                    ]
                  )
            } @{ $endUsers->{list} }
        ]
    );

    $endUsers = Labelset( list => $allTariffsByEndUser->{groups} );

    my $discounts = $model->{ldnoRev} =~ /5/
      ? Dataset(
        name          => 'LDNO discounts',
        cols          => $cdcmLevels,
        rows          => $ldnoLevels,
        defaultFormat => '%hardnz',
        data          => [
            map {
                [ map { '' } @{ $ldnoLevels->{list} } ]
            } @{ $cdcmLevels->{list} }
        ],
        number   => 1181,
        dataset  => $model->{dataset},
        appendTo => $model->{inputTables},
      )
      : Dataset(
        name          => 'LDNO discounts',
        rows          => $cdcmLevels,
        cols          => $ldnoLevels,
        defaultFormat => '%hardnz',
        data          => [
            map {
                [ map { '' } @{ $cdcmLevels->{list} } ]
            } @{ $ldnoLevels->{list} }
        ],
        number   => 1181,
        dataset  => $model->{dataset},
        appendTo => $model->{inputTables},
      );

    my @endUserTariffs = map {
        my $regexp = '^' . ( '.' x $_ ) . 'y';
        Dataset(
            name          => $tariffComponents[$_],
            defaultFormat => /day/ ? '0.00hard' : '0.000hard',
            rows          => $endUsers,
            data => [ map { /$regexp/ ? '' : undef } @tariffComponentMatrix ],
            dataset => $model->{dataset},
        );
    } 0 .. $#tariffComponents;

    Columnset(
        number   => 1182,
        dataset  => $model->{dataset},
        appendTo => $model->{inputTables},
        name     => 'CDCM end user tariffs',
        columns  => \@endUserTariffs,
    );

    my $discountsByTariff = $model->{ldnoRev} =~ /5/
      ? new SpreadsheetModel::Custom(
        name          => 'Applicable discount for each tariff',
        rows          => $allTariffsByEndUser,
        defaultFormat => '%copy',
        arithmetic    => '= IV1',
        custom        => ['=IV1'],
        objectType    => 'Special copy',
        arguments     => { IV1 => $discounts },
        wsPrepare     => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                local $_ = $allTariffsByEndUser->{list}[$y];
                $y = 0 if s/^LDNO 0000: //;
                $y = 2 if s/^LDNO 132kV\/EHV: //;
                $y = 1 if s/^LDNO 132kV: //;
                $y = 3 if s/^LDNO EHV: //;
                $y = 4 if s/^LDNO HVplus: //;
                $x =
                    /^HV Sub Gen/i ? 40
                  : /^HV Sub/i     ? 30
                  : /^HV Gen/i     ? 3
                  : /^HV/i         ? 2
                  : /^LV Sub Gen/i ? 2
                  : /^LV Sub/i     ? 1
                  : /^LV Gen/i     ? 1
                  :                  0;
                return '#VALUE!', $format if $x > 3;
                '', $format, $formula->[0],
                  IV1 => Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{IV1} + $y,
                    $colh->{IV1} + $x,
                    1, 1
                  );
            };
        }
      )
      : new SpreadsheetModel::Custom(
        name          => 'Applicable discount for each tariff',
        rows          => $allTariffsByEndUser,
        defaultFormat => '%copy',
        arithmetic    => '= IV1',
        custom        => ['=IV1'],
        objectType    => 'Special copy',
        arguments     => { IV1 => $discounts },
        wsPrepare     => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                local $_ = $allTariffsByEndUser->{list}[$y];
                $x = 0  if s/^LDNO 0000: //;
                $x = 1  if s/^LDNO 1000: //;
                $x = 2  if s/^LDNO 1100: //;
                $x = 3  if s/^LDNO 0100: //;
                $x = 4  if s/^LDNO 1110: //;
                $x = 5  if s/^LDNO 0110: //;
                $x = 6  if s/^LDNO 0010: //;
                $x = 7  if s/^LDNO 0001: //;
                $x = 8  if s/^LDNO 0002: //;
                $x = 9  if s/^LDNO 1001: //;
                $x = 10 if s/^LDNO 0011: //;
                $x = 11 if s/^LDNO 0111: //;
                $x = 12 if s/^LDNO 0101: //;
                $x = 13 if s/^LDNO 1101: //;
                $x = 14 if s/^LDNO 1111: //;
                $y =
                    /^HV Sub Gen/i ? 40
                  : /^HV Sub/i     ? 30
                  : /^HV Gen/i     ? 3
                  : /^HV/i         ? 2
                  : /^LV Sub Gen/i ? 2
                  : /^LV Sub/i     ? 1
                  : /^LV Gen/i     ? 1
                  :                  0;
                return '#VALUE!', $format if $y > 3;
                '', $format, $formula->[0],
                  IV1 => Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{IV1} + $y,
                    $colh->{IV1} + $x,
                    1, 1
                  );
            };
        }
      );

    my @explodedData = map {
        my $unexploded = $endUserTariffs[$_]{data};
        [ map { defined $_ ? $unexploded->[$_] : undef; }
              @{ $allTariffsByEndUser->{groupid} } ];
    } 0 .. $#tariffComponents;

    my @allTariffs = map {
        Arithmetic(
            name          => $tariffComponents[$_],
            defaultFormat => $tariffComponents[$_] =~ /day/
            ? '0.00softnz'
            : '0.000softnz',
            arithmetic => '=IV2*(1-IV1)',
            arguments  => {
                IV2 => $endUserTariffs[$_],
                IV1 => $discountsByTariff,
            },
            rowFormats => [
                map { defined $_ ? undef : 'unavailable' }
                  @{ $explodedData[$_] }
            ],
          )
    } 0 .. $#tariffComponents;

    return Notes( lines => 'LDNO discounted tariffs' ),
      Columnset(
        name    => 'Discounted LDNO tariffs',
        columns => \@allTariffs
      ) if $model->{ldnoRev} =~ /tar/i;

    my @volnames = split /\n/, <<EOF ;
Rate 1 units (MWh)
Rate 2 units (MWh)
Rate 3 units (MWh)
MPANs
Import capacity (kVA)
Reactive power units (MVArh)
EOF

    my @volumeData = map {
        Dataset(
            name          => $volnames[$_],
            rows          => $allTariffsByEndUser,
            data          => $explodedData[$_],
            dataset       => $model->{dataset},
            defaultFormat => $volnames[$_] =~ /M(?:W|VAr)h/
            ? '0.000hardnz'
            : '0hardnz',
        );
    } 0 .. $#tariffComponents;

    Columnset(
        number   => 1183,
        dataset  => $model->{dataset},
        appendTo => $model->{inputTables},
        name     => 'LDNO volume data',
        columns  => \@volumeData
    );

    my $revenueByTariff;

    {
        my @termsNoDays;
        my @termsWithDays;
        my %args = ( IV400 => $model->{daysInYear} );
        foreach ( 0 .. $#tariffComponents ) {
            my $pad = $_ + 1;
            $pad = "0$pad" while length $pad < 3;
            if ( $tariffComponents[$_] =~ m#/day# ) {
                push @termsWithDays, "IV2$pad*IV3$pad";
            }
            else {
                push @termsNoDays, "IV2$pad*IV3$pad";
            }
            $args{"IV2$pad"} = $allTariffs[$_];
            $args{"IV3$pad"} = $volumeData[$_];
        }
        $revenueByTariff = Arithmetic(
            name       => 'Net revenue from discounted LDNO tariffs (£/year)',
            rows       => $allTariffsByEndUser,
            arithmetic => '='
              . join( '+',
                @termsWithDays
                ? ( '0.01*IV400*(' . join( '+', @termsWithDays ) . ')' )
                : (),
                @termsNoDays ? ( '10*(' . join( '+', @termsNoDays ) . ')' )
                : ('0'),
              ),
            arguments     => \%args,
            defaultFormat => '0softnz'
          )
    }

    1 and Columnset(
        name    => 'LDNO discounted CDCM tariffs',
        columns => \@allTariffs,
    );

    my @result = (
        Notes( lines => 'LDNO revenue model' ),
        GroupBy(
            name => 'Total net revenue from discounted LDNO tariffs (£/year)',
            defaultFormat => '0softnz',
            source        => $revenueByTariff
        )
    );

    if ( $model->{ldnoRev} =~ /5/ )
    {    # reorder tariffs only if using five discounts; how weird is that?
        my $allTariffsReordered = Labelset(
            name => 'All tariffs (reordered)',
            list => [
                ( grep { /^LDNO HVplus/i } @{ $allTariffsByEndUser->{list} } ),
                ( grep { /^LDNO EHV/i } @{ $allTariffsByEndUser->{list} } ),
                (
                    grep { /^LDNO 132kV\/EHV/i }
                      @{ $allTariffsByEndUser->{list} }
                ),
                (
                    grep { /^LDNO 132kV/i && !/^LDNO 132kV\/EHV/i }
                      @{ $allTariffsByEndUser->{list} }
                ),
                ( grep { /^LDNO 0000/i } @{ $allTariffsByEndUser->{list} } ),
            ]
        );
        push @result, Columnset(
            name    => 'LDNO discounted CDCM tariffs (reordered)',
            columns => [
                map {
                    my $oldRowFormats = $_->{rowFormats};
                    my %rowFormatMap  = map {
                        ( $allTariffsByEndUser->{list}[$_] =>
                              $oldRowFormats->[$_] );
                    } 0 .. $#{ $allTariffsByEndUser->{list} };
                    Arithmetic(
                        defaultFormat => $_->{defaultFormat},
                        name          => $_->{name},
                        rows          => $allTariffsReordered,
                        arguments     => { IV1 => $_ },
                        arithmetic    => '=IV1',
                        rowFormats    => [
                            @rowFormatMap{ @{ $allTariffsReordered->{list} } }
                        ]
                    );
                } @allTariffs
            ]
        );
    }

    @result;

}

1;
