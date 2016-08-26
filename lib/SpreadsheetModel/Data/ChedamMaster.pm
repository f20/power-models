package SpreadsheetModel::Data::Chedam;

=head Copyright licence and disclaimer

Copyright 2013 Franck Latrémolière and others. All rights reserved.

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

sub runFromDatabase {
    require SpreadsheetModel::Data::Database;
    require SpreadsheetModel::Data::DotDiagrams;
    require SpreadsheetModel::Data::ChedamDataLocator;
    require SpreadsheetModel::Data::ChedamToDot;
    my ( $dataReader, $bookTableIndexHash ) = Compilation->makeDatabaseReader;
    SpreadsheetModel::Data::DotDiagrams::writeDotDiagrams(
        map { $_->calculate->toDot } map {
            my $filename = $_;
            map { $dataReader->( $bookTableIndexHash->{$filename}{bid}, $_ ); }
              exists $bookTableIndexHash->{$filename}{1703}
              ? (
                SpreadsheetModel::Data::Chedam->locateHidamModelled($filename),
                SpreadsheetModel::Data::Chedam->locateHidamAdjMMD($filename),
                SpreadsheetModel::Data::Chedam->locateHidamActualCap($filename),
                SpreadsheetModel::Data::Chedam->locateHidamActualMD($filename),
              )
              : (),
              exists $bookTableIndexHash->{$filename}{1017}
              ? ( SpreadsheetModel::Data::Chedam->locateDrm($filename), )
              : (),
        } sort keys %$bookTableIndexHash
    );
}

sub calculate {
    my ($data) = @_;
    if ( !exists $data->{MW_BSP} && defined $data->{Div_GSP} ) {
        $data->{MW_BSP} = ( 1 + $data->{Div_BSP} ) * $data->{SMD_132kV_EHV};
        $data->{MW_Dis} = ( 1 + $data->{Div_Dis} ) * $data->{SMD_HV_LV};
        $data->{MW_GSP_132kV} = ( 1 + $data->{Div_GSP} ) * $data->{SMD_132kV};
        $data->{MW_GSP_EHV}   = 0;
        $data->{MW_GSP_HV}    = 0;
        $data->{MW_Pri_132kV} =
          ( 1 + $data->{Div_BSP} ) * $data->{SMD_132kV_HV};
        $data->{MW_Pri_EHV} = ( 1 + $data->{Div_Pri} ) * $data->{SMD_EHV_HV};
    }
    $data;
}

1;
