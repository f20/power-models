package EDCM2;

=head Copyright licence and disclaimer

Copyright 2009-2012 Energy Networks Association Limited and others.
Copyright 2016 Franck Latrémolière, Reckon LLP and others.

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

sub ldnoRev {
    my ($model) = @_;

    my ( @endUsers, @tariffComponentMatrix );

    my @tariffData = $model->{dcp179}
      ? split /\n/, <<EOL
ynnynn	Domestic Unrestricted
yynynn	Domestic Two Rate
ynnnnn	Domestic Off Peak (related MPAN)
ynnynn	Small Non Domestic Unrestricted
yynynn	Small Non Domestic Two Rate
ynnnnn	Small Non Domestic Off Peak (related MPAN)
yynynn	LV Medium Non-Domestic
yynynn	LV Sub Medium Non-Domestic
yynynn	HV Medium Non-Domestic
yyyynn	LV Network Domestic
yyyynn	LV Network Non-Domestic Non-CT
yyyyyy	LV HH Metered
yyyyyy	LV Sub HH Metered
yyyyyy	HV HH Metered
ynnnnn	NHH UMS category A
ynnnnn	NHH UMS category B
ynnnnn	NHH UMS category C
ynnnnn	NHH UMS category D
yyynnn	LV UMS (Pseudo HH Metered)
ynnynn	LV Generation NHH or Aggregate HH
ynnynn	LV Sub Generation NHH
ynnyny	LV Generation Intermittent
yyyyny	LV Generation Non-Intermittent
ynnyny	LV Sub Generation Intermittent
yyyyny	LV Sub Generation Non-Intermittent
ynnyny	HV Generation Intermittent
yyyyny	HV Generation Non-Intermittent
EOL
      : $model->{dcp130} ? split /\n/, <<EOL
ynnynn	Domestic Unrestricted
yynynn	Domestic Two Rate
ynnnnn	Domestic Off Peak (related MPAN)
ynnynn	Small Non Domestic Unrestricted
yynynn	Small Non Domestic Two Rate
ynnnnn	Small Non Domestic Off Peak (related MPAN)
yynynn	LV Medium Non-Domestic
yynynn	LV Sub Medium Non-Domestic
yynynn	HV Medium Non-Domestic
yyyyyy	LV HH Metered
yyyyyy	LV Sub HH Metered
yyyyyy	HV HH Metered
ynnnnn	NHH UMS category A
ynnnnn	NHH UMS category B
ynnnnn	NHH UMS category C
ynnnnn	NHH UMS category D
yyynnn	LV UMS (Pseudo HH Metered)
ynnynn	LV Generation NHH
ynnynn	LV Sub Generation NHH
ynnyny	LV Generation Intermittent
yyyyny	LV Generation Non-Intermittent
ynnyny	LV Sub Generation Intermittent
yyyyny	LV Sub Generation Non-Intermittent
ynnyny	HV Generation Intermittent
yyyyny	HV Generation Non-Intermittent
EOL
      : split /\n/, <<EOL
ynnynn	Domestic Unrestricted
yynynn	Domestic Two Rate
ynnnnn	Domestic Off Peak (related MPAN)
ynnynn	Small Non Domestic Unrestricted
yynynn	Small Non Domestic Two Rate
ynnnnn	Small Non Domestic Off Peak (related MPAN)
yynynn	LV Medium Non-Domestic
yynynn	LV Sub Medium Non-Domestic
yynynn	HV Medium Non-Domestic
yyyyyy	LV HH Metered
yyyyyy	LV Sub HH Metered
yyyyyy	HV HH Metered
ynnnnn	NHH UMS
yyynnn	LV UMS (Pseudo HH Metered)
ynnynn	LV Generation NHH
ynnynn	LV Sub Generation NHH
ynnyny	LV Generation Intermittent
yyyyny	LV Generation Non-Intermittent
ynnyny	LV Sub Generation Intermittent
yyyyny	LV Sub Generation Non-Intermittent
ynnyny	HV Generation Intermittent
yyyyny	HV Generation Non-Intermittent
EOL
      ;

    @tariffData = grep { !/\bMedium\b/i; } @tariffData if $model->{dcp270};

    foreach (@tariffData) {
        if ( my ( $a, $b ) = /^([yn]+)\s+(.+)/ ) {
            if ( $model->{dcp137} && $b =~ /HV Generation/i ) {
                push @tariffComponentMatrix, $a, $a, $a, $a;
                push @endUsers, $b, "$b Low GDA", "$b Medium GDA",
                  "$b High GDA";
            }
            else {
                push @tariffComponentMatrix, $a;
                push @endUsers,              $b;
            }
        }
    }

    my $endUsers = Labelset( list => \@endUsers );

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

    my @volnames = split /\n/, <<EOF ;
Rate 1 units (MWh)
Rate 2 units (MWh)
Rate 3 units (MWh)
MPANs
Import capacity (kVA)
Reactive power units (MVArh)
EOF

    if ( $model->{dcp161} ) {
        splice @tariffComponents, 5, 0, 'Exceeded capacity charge p/kVA/day';
        splice @volnames,         5, 0, 'Exceeded capacity (kVA)';
        s/^(....)(.)/$1$2$2/ foreach @tariffComponentMatrix;
    }

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

    my $ppu;
    $ppu = Dataset(
        name     => 'All-the-way reference p/kWh values',
        rows     => $endUsers,
        data     => [ map { 1 } @{ $endUsers->{list} } ],
        number   => 1185,
        dataset  => $model->{dataset},
        appendTo => $model->{inputTables},
    ) if $model->{ldnoRev} =~ /ppu/i;

    my $discounts = Dataset(
        name => 'LDNO discount ' . ( $ppu ? 'p/kWh' : 'percentage' ),
        cols => $cdcmLevels,
        rows => $ldnoLevels,
        $ppu ? () : ( defaultFormat => '%hardnz' ),
        data => [
            map {
                [ map { '' } @{ $ldnoLevels->{list} } ]
            } @{ $cdcmLevels->{list} }
        ],
        number     => 1181,
        dataset    => $model->{dataset},
        appendTo   => $model->{inputTables},
        validation => {
            validate      => 'decimal',
            criteria      => '>=',
            value         => 0,
            input_title   => 'LDNO discount:',
            input_message => 'At least zero',
            error_message => 'The LDNO discount must not be negative'
        },
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

    my $discountsByTariff = new SpreadsheetModel::Custom(
        name => 'Applicable discount for each tariff',
        rows => $allTariffsByEndUser,
        $ppu ? () : ( defaultFormat => '%copy' ),
        arithmetic => '= A1',
        custom     => ['=A1'],
        objectType => 'Special copy',
        arguments  => { A1 => $discounts },
        wsPrepare  => sub {
            my ( $self, $wb, $ws, $format, $formula, $pha, $rowh, $colh ) = @_;
            sub {
                my ( $x, $y ) = @_;
                local $_ = $allTariffsByEndUser->{list}[$y];
                $y = 0  if s/^LDNO 0000: //;
                $y = 2  if s/^LDNO 132kV\/EHV: //;
                $y = 1  if s/^LDNO 132kV: //;
                $y = 3  if s/^LDNO EHV: //;
                $y = 4  if s/^LDNO HVplus: //;
                $y = 1  if s/^LDNO 1000: //;
                $y = 2  if s/^LDNO 1100: //;
                $y = 3  if s/^LDNO 0100: //;
                $y = 4  if s/^LDNO 1110: //;
                $y = 5  if s/^LDNO 0110: //;
                $y = 6  if s/^LDNO 0010: //;
                $y = 7  if s/^LDNO 0001: //;
                $y = 8  if s/^LDNO 0002: //;
                $y = 9  if s/^LDNO 1001: //;
                $y = 10 if s/^LDNO 0011: //;
                $y = 11 if s/^LDNO 0111: //;
                $y = 12 if s/^LDNO 0101: //;
                $y = 13 if s/^LDNO 1101: //;
                $y = 14 if s/^LDNO 1111: //;
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
                  qr/\bA1\b/ =>
                  Spreadsheet::WriteExcel::Utility::xl_rowcol_to_cell(
                    $rowh->{A1} + $y,
                    $colh->{A1} + $x,
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
        $ppu
          ? Arithmetic(
            name          => $tariffComponents[$_],
            defaultFormat => $tariffComponents[$_] =~ /day/
            ? '0.00softnz'
            : '0.000softnz',
            arithmetic => '=IF(A31,ROUND(A2*(1-A1/A3),'
              . ( $tariffComponents[$_] =~ /day/ ? 2 : 3 )
              . '),A21)',
            arguments => {
                A1  => $discountsByTariff,
                A2  => $endUserTariffs[$_],
                A21 => $endUserTariffs[$_],
                A3  => $ppu,
                A31 => $ppu,
            },
            rowFormats => [
                map { defined $_ ? undef : 'unavailable' }
                  @{ $explodedData[$_] }
            ],
          )
          : Arithmetic(
            name          => $tariffComponents[$_],
            defaultFormat => $tariffComponents[$_] =~ /day/
            ? '0.00softnz'
            : '0.000softnz',
            arithmetic => $model->{ldnoRev} !~ /round/i ? '=A2*(1-A1)'
            : '=ROUND(A2*(1-A1),'
              . ( $tariffComponents[$_] =~ /day/ ? 2 : 3 ) . ')',
            arguments => {
                A2 => $endUserTariffs[$_],
                A1 => $discountsByTariff,
            },
            rowFormats => [
                map { defined $_ ? undef : 'unavailable' }
                  @{ $explodedData[$_] }
            ],
          );
    } 0 .. $#tariffComponents;

    return Notes( lines => 'LDNO discounted tariffs' ), undef,
      Columnset(
        name    => 'Discounted LDNO tariffs',
        columns => \@allTariffs
      ) if $model->{ldnoRev} =~ /tar/i;

    my @volumeData = map {
        Dataset(
            name          => $volnames[$_],
            rows          => $allTariffsByEndUser,
            data          => $explodedData[$_],
            dataset       => $model->{dataset},
            defaultFormat => $volnames[$_] =~ /M(?:W|VAr)h/
            ? '0.000hardnz'
            : '0hardnz',
            validation => {
                validate      => 'decimal',
                criteria      => '>=',
                value         => 0,
                input_title   => 'Volume:',
                input_message => 'At least 0',
                error_message => 'The volume must not be negative.'
            },
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
        my %args = ( A400 => $model->{daysInYear} );
        foreach ( 0 .. $#tariffComponents ) {
            my $pad = $_ + 1;
            $pad = "0$pad" while length $pad < 3;
            if ( $tariffComponents[$_] =~ m#/day# ) {
                push @termsWithDays, "A2$pad*A3$pad";
            }
            else {
                push @termsNoDays, "A2$pad*A3$pad";
            }
            $args{"A2$pad"} = $allTariffs[$_];
            $args{"A3$pad"} = $volumeData[$_];
        }
        $revenueByTariff = Arithmetic(
            name       => 'Net revenue from discounted LDNO tariffs (£/year)',
            rows       => $allTariffsByEndUser,
            arithmetic => '='
              . join( '+',
                @termsWithDays
                ? ( '0.01*A400*(' . join( '+', @termsWithDays ) . ')' )
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
                        arguments     => { A1 => $_ },
                        arithmetic    => '=A1',
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
