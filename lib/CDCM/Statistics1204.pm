package CDCM;

=head Copyright licence and disclaimer

Copyright 2017 Franck Latrémolière, Reckon LLP and others.

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

sub makeStatisticsTables1204 {

    my ( $model, $tariffTable, $daysInYear, $nonExcludedComponents,
        $componentMap, )
      = @_;

    my ($allTariffs) = values %$tariffTable;
    $allTariffs = $allTariffs->{rows};
    $allTariffs = Labelset(
        list => [ grep { !/^(?:QNO|LDNO)/; } @{ $allTariffs->{list} } ] );

    my %args = ( A7 => $daysInYear );
    my ( @columns, @charge1, @charge2, @units );

    my %componentVolumeNameMap = (
        (
            map { ( "Unit rate $_ p/kWh", "Rate $_ units (MWh)" ) }
              1 .. $model->{maxUnitRates}
        ),
        split "\n",
        <<'EOL' );
Fixed charge p/MPAN/day
MPANs
Capacity charge p/kVA/day
Import capacity (kVA)
Unauthorised demand charge p/kVAh
Unauthorised demand (MVAh)
Exceeded capacity charge p/kVA/day
Exceeded capacity (kVA)
Generation capacity rate p/kW/day
Generation capacity (kW)
Reactive power charge p/kVArh
Reactive power units (MVArh)
EOL

    my $counter = 10;
    foreach (@$nonExcludedComponents) {
        push @columns,
          my $vol = Dataset(
            name          => $componentVolumeNameMap{$_},
            rows          => $allTariffs,
            defaultFormat => '0hard',
            data          => [ map { ''; } @{ $allTariffs->{list} } ],
          );
        if (/kWh/) {
            ++$counter;
            $args{"A$counter"} = $vol;
            push @units, "A$counter";
        }
        ++$counter;
        my $prod = "A$counter";
        $args{"A$counter"} = $vol;
        ++$counter;
        $prod .= "*A$counter";
        $args{"A$counter"} = $tariffTable->{$_};
        if (/day/) {
            push @charge2, $prod;
        }
        else {
            push @charge1, $prod;
        }
    }

    Columnset(
        name     => 'Volume assumptions for illustrative unit rates',
        columns  => \@columns,
        number   => 1204,
        appendTo => $model->{inputTables},
        dataset  => $model->{dataset},
    );

    my $ppu = Arithmetic(
        name          => Label( 'p/kWh', 'Illustrative average p/kWh' ),
        rows          => $allTariffs,
        defaultFormat => '0.000soft',
        arithmetic    => '=('
          . join( '+', @charge1, '0.001*A7*(' . join( '+', @charge2 ) ) . '))/('
          . join( '+', @units ) . ')',
        arguments => \%args,
    );

    $model->{sharedData}->addStats( 'Illustrative average p/kWh', $model, $ppu )
      if $model->{sharedData};

    push @{ $model->{statisticsTables} }, $ppu;

}

1;
