package Impact;

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
use SpreadsheetModel::Shortcuts ':all';
use SpreadsheetModel::MatrixSheet;

sub processTable3701 {

    my ( $model, $dno, $title, $baselineData, $scenarioData, ) = @_;
    my $bd = $baselineData->{3701} or return;
    my $sd = $scenarioData->{3701} or return;
    my @tariffs = @{ $sd->[0] };
    shift @tariffs;
    my $tariffSet = Labelset( list => \@tariffs );
    my %bTariffNo =
      map { $bd->[0][$_] => $_; } 0 .. $#{ $bd->[0] };
    my @bTariffMap = map { $bTariffNo{$_}; } @tariffs;

    my @baselineTariffs = map {
        my $col = $_;
        Constant(
            name          => $bd->[$col][0],
            defaultFormat => $bd->[$col][0] =~ /day/
            ? '0.00copy'
            : '0.000copy',
            rows => $tariffSet,
            data => [ map { $_ ? $bd->[$col][$_] : undef; } @bTariffMap ],
        );
      } grep { defined $bd->[$_][0] && $bd->[$_][0] !~ /LLFC|PC|checksum/; }
      1 .. $#$bd;

    my @scenarioTariffs = map {
        my $col = $_;
        Constant(
            name          => $sd->[$col][0],
            defaultFormat => $sd->[$col][0] =~ /day/
            ? '0.00copy'
            : '0.000copy',
            rows => $tariffSet,
            data => [ map { $sd->[$col][$_]; } 1 .. @tariffs ],
        );
      } grep { defined $sd->[$_][0] && $sd->[$_][0] !~ /LLFC|PC|checksum/; }
      1 .. $#$sd;

    SpreadsheetModel::MatrixSheet->new( verticalSpace => 2 )->addDatasetGroup(
        name    => 'Baseline tariffs',
        columns => \@baselineTariffs,
      )->addDatasetGroup(
        name    => 'Scenario tariffs',
        columns => \@scenarioTariffs,
      );

    push @{ $model->{sheetNames} }, $dno;
    push @{ $model->{sheetTables} },
      [ Notes( name => $title ), @baselineTariffs, @scenarioTariffs ];

}

1;
