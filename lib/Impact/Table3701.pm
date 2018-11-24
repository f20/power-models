package Impact;

# Copyright 2017 Franck Latrémolière, Reckon LLP and others.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice,
# this list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
# this list of conditions and the following disclaimer in the documentation
# and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
# EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
# THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

use warnings;
use strict;
use utf8;
use SpreadsheetModel::Shortcuts ':all';

sub processTable3701 {

    my ( $model, $baselineData, $scenarioData, $sheetName, $sheetTitle, ) = @_;
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

    SpreadsheetModel::MatrixSheet->new(
        noLines       => 1,
        noDoubleNames => 1,
        noNumbers     => 1,
      )->addDatasetGroup(
        name    => 'Baseline tariffs',
        columns => \@baselineTariffs,
      )->addDatasetGroup(
        name    => 'Scenario tariffs',
        columns => \@scenarioTariffs,
      );

    push @{ $model->{worksheetsAndClosures} }, $sheetName => sub {
        my ( $wsheet, $wbook ) = @_;
        $wsheet->set_column( 0, 0,   48 );
        $wsheet->set_column( 1, 254, 16 );
        $wsheet->freeze_panes( 5, 1 );
        $_->wsWrite( $wbook, $wsheet )
          foreach Notes( name => $sheetTitle ), @baselineTariffs,
          @scenarioTariffs;
    };

}

1;
