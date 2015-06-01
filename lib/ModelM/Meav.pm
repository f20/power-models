package ModelM;

=head Copyright licence and disclaimer

Copyright 2011 The Competitive Networks Association and others.
Copyright 2012-2014 Franck Latrémolière, Reckon LLP and others.

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

sub meavRawData {

    my ($model) = @_;
    return unless $model->{meav};
    return $model->{objects}{meav} if ref $model->{objects}{meav};
    my $rows = Labelset( name => 'MEAV rows', list => [ split /\n/, <<EOL] );
LV main overhead line km
LV service overhead
LV overhead support
LV main underground consac km
LV main underground plastic km
LV main underground paper km
LV service underground 
LV pillar indoors
LV pillar outdoors
LV board wall-mounted
LV board underground
LV fuse pole-mounted
LV fuse tower-mounted
6.6/11kV overhead open km
6.6/11kV overhead covered km
20kV overhead open pm
20kV overhead covered km
6.6/11kV overhead support
20kV overhead support
6.6/11kV underground km
20kV underground km
HV submarine km
6.6/11kV breaker pole-mounted
6.6/11kV breaker ground-mounted
6.6/11kV switch pole-mounted
6.6/11kV switch ground-mounted
6.6/11kV ring main unit
6.6/11kV other switchgear pole-mounted
6.6/11kV other switchgear ground-mounted
20kV breaker pole-mounted
20kV breaker ground-mounted
20kV switch pole-mounted
20kV switch ground-mounted
20kV ring main unit
20kV other switchgear pole-mounted
20kV other switchgear ground-mounted
6.6/11kV transformer pole-mounted
6.6/11kV transformer ground-mounted
20kV transformer pole-mounted
20kV transformer ground-mounted
33kV overhead pole line km
33kV overhead tower line km
66kV overhead pole line km
66kV overhead tower line km
33kV pole
33kV tower
66kV pole
66kV tower
33kV underground non-pressurised km
33kV underground oil km
33kV underground gas km
66kV underground non-pressurised km
66kV underground oil km
66kV underground gas km
EHV submarine km
33kV breaker indoors
33kV breaker outdoors
33kV switch ground-mounted
33kV switch pole-mounted
33kV ring main unit
33kV other switchgear
66kV breaker
66kV other switchgear
33kV transformer pole-mounted
33kV transformer ground-mounted
33kV auxiliary transformer
66kV transformer
66kV auxiliary transformer
132kV overhead pole conductor km
132kV overhead tower conductor km 
132kV pole
132kV tower
132kV tower fittings
132kV underground non-pressurised km
132kV underground oil km
132kV underground gas km
132kV submarine km
132kV breaker
132kV other switchgear
132kV transformer
132kV auxiliary transformer
132kV/EHV remote terminal unit pole-mounted
132kV/EHV remote terminal unit ground-mounted
HV remote terminal unit pole-mounted
HV remote terminal unit ground-mounted
EOL
    my @d = (
        Dataset(
            name          => 'Asset quantity',
            defaultFormat => '0hard',
            rows          => $rows,
            data          => [ map { 0 } @{ $rows->{list} } ],
            validation    => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
        ),
        Dataset(
            name          => 'Unit MEAV (£)',
            defaultFormat => '0hard',
            rows          => $rows,
            data          => [ map { 0 } @{ $rows->{list} } ],
            validation    => {
                validate => 'decimal',
                criteria => '>=',
                value    => 0,
            },
        ),
    );
    Columnset(
        name => 'MEAV data',
        lines =>
          'In a legacy Method M workbook, these data are on sheet Data-MEAV.',
        defaultFormat => '%hard',
        number        => 1355,
        dataset       => $model->{dataset},
        appendTo      => $model->{objects}{inputTables},
        columns       => \@d,
    );
    $model->{objects}{meav} = Arithmetic(
        name          => 'MEAV (£)',
        defaultFormat => 'millionsoft',
        arithmetic    => '=IV1*IV2',
        arguments     => { IV1 => $d[0], IV2 => $d[1], },
    );
}

sub meavPercentages {

    my ( $model, $allocLevelset ) = @_;

    my $meav = $model->meavRawData;
    return $model->{objects}{meavPercentages}{ 0 + $allocLevelset } ||=
      Stack(    # for Numbers for iPad which cannot do SUMPRODUCT across sheets
        sources => [
            Dataset(
                name => 'MEAV percentages',
                lines =>
                  'In a pre-DCP 118 legacy Method M workbook, these data are on'
                  . ' sheet Calc-MEAV, possibly starting at cell H6.',
                data          => [ map { 0 } @{ $allocLevelset->{list} } ],
                defaultFormat => '%hard',
                number        => 1350,
                cols       => $allocLevelset,
                dataset    => $model->{dataset},
                appendTo   => $model->{objects}{inputTables},
                validation => {
                    validate => 'decimal',
                    criteria => '>=',
                    value    => 0,
                },
            )
        ]
      ) unless $meav;

    return $model->{objects}{meavPercentages}{ 0 + $allocLevelset }{ 0 + $meav }
      if $model->{objects}{meavPercentages}{ 0 + $allocLevelset }{ 0 + $meav };

    my $mapping = Constant(
        name          => 'MEAV mapping',
        rows          => $meav->{rows},
        cols          => $allocLevelset,
        byrow         => 1,
        defaultFormat => '0con',
        data          => [
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
        ],
    );

    if ( ref( $meav->{location} ) eq 'SpreadsheetModel::Columnset' ) {
        push @{ $meav->{location}{columns} }, $mapping;
        $mapping->{location} = $meav->{location};
    }
    else {
        Columnset(
            name    => 'MEAV calculations',
            columns => [ $meav, $mapping ]
        );
    }

    my $totalMeav = SumProduct(
        name          => 'MEAV by network level (£)',
        defaultFormat => $meav->{defaultFormat},
        matrix        => $mapping,
        vector        => $meav,
    );

    $model->{objects}{meavPercentages}{ 0 + $allocLevelset }{ 0 + $meav } =
      Arithmetic(
        name          => 'MEAV percentages',
        defaultFormat => '%soft',
        cols          => $allocLevelset,
        arithmetic    => '=IV6/SUM(IV8_IV9)',
        arguments     => {
            IV6     => $totalMeav,
            IV8_IV9 => $totalMeav,
        },
      );

}

sub meavPercentageServiceLV {

    my ( $model, $lvOnly, $lvServiceOnly ) = @_;

    my $meav = $model->meavRawData;
    return Dataset(
        name          => 'MEAV: ratio of LV services to LV total',
        data          => [.5],
        defaultFormat => '%hard',
        number        => 1360,
        cols          => $lvOnly,
        rows          => $lvServiceOnly,
        dataset       => $model->{dataset},
        appendTo      => $model->{objects}{inputTables},
        validation    => {
            validate => 'decimal',
            criteria => '>=',
            value    => 0,
        },
    ) unless $meav;

    my $lvService = Constant(
        name          => 'MEAV mapping: LV services',
        rows          => $meav->{rows},
        defaultFormat => '0con',
        data          => [ undef, 1, ( map { undef } 1 .. 4 ), 1 ],
    );

    my $lvTotal = Constant(
        name          => 'MEAV mapping: LV total',
        rows          => $meav->{rows},
        defaultFormat => '0con',
        data          => [ map { 1; } 1 .. 13 ],
    );

    if ( ref( $meav->{location} ) eq 'SpreadsheetModel::Columnset' ) {
        push @{ $meav->{location}{columns} }, $lvService, $lvTotal;
        $_->{location} = $meav->{location} foreach $lvService, $lvTotal;
    }
    else {
        Columnset(
            name    => 'MEAV calculations',
            columns => [ $meav, $lvService, $lvTotal, ]
        );
    }

    Arithmetic(
        name          => 'MEAV: ratio of LV services to LV total',
        defaultFormat => '%soft',
        cols          => $lvOnly,
        rows          => $lvServiceOnly,
        arithmetic =>
          '=SUMPRODUCT(IV2_IV3,IV4_IV5)/SUMPRODUCT(IV6_IV7,IV8_IV9)',
        arguments => {
            IV2_IV3 => $lvService,
            IV4_IV5 => $meav,
            IV6_IV7 => $lvTotal,
            IV8_IV9 => $meav,
        },
    );

}

sub meavPercentagesEdcm {

    my ( $model, $edcmLevelset ) = @_;

    my $meav = $model->meavRawData or die 'EDCM method M needs raw MEAV data';

    my $mapping = Constant(
        name          => 'MEAV EDCM mapping',
        rows          => $meav->{rows},
        cols          => $edcmLevelset,
        byrow         => 1,
        defaultFormat => '0con',
        data          => [
            [],
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            undef,
            undef,
            undef,
            undef,
            undef,
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 1, 0, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 1, 0, 0, 0 ],
            [ 0, 0, 1, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 0, 1 ],
            [ 0, 0, 1, 0 ],
            [ 0, 0, 1, 0 ],
            [ 1, 0, 0, 0 ],
            [ 1, 0, 0, 0 ],
            undef,
            undef,
            undef,
            undef,
            undef,
            undef,
        ],
    );

    if ( ref( $meav->{location} ) eq 'SpreadsheetModel::Columnset' ) {
        push @{ $meav->{location}{columns} }, $mapping;
        $mapping->{location} = $meav->{location};
    }
    else {
        Columnset(
            name    => 'MEAV calculations',
            columns => [ $meav, $mapping ]
        );
    }

    my $totalMeav = SumProduct(
        name          => 'MEAV by network level (£)',
        defaultFormat => $meav->{defaultFormat},
        matrix        => $mapping,
        vector        => $meav,
    );

    Arithmetic(
        name          => 'MEAV percentages',
        defaultFormat => '%soft',
        cols          => $edcmLevelset,
        arithmetic    => '=IV6/SUM(IV8_IV9)',
        arguments     => {
            IV6     => $totalMeav,
            IV8_IV9 => $totalMeav,
        },
    );

}

1;
